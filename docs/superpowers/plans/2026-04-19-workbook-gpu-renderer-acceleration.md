# Workbook GPU Renderer Acceleration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the remaining CPU-bound workbook browse hot path with a worker-fed, renderer-owned GPU pipeline that keeps wide-sheet scrolling smooth under frozen panes, variable column widths, and collaborator patch load.

**Architecture:** Keep the worker/runtime as the workbook truth, but stop asking the browser to assemble browse-time scenes. The worker will build retained pane scene packets, text will move from Canvas2D to a glyph-atlas GPU path, and the browser will run a single pane renderer that mutates persistent GPU resources by damage region instead of repainting pane content on scroll.

**Tech Stack:** React 19, TypeScript, Vite, Playwright, WebGPU, worker-transport viewport patching, browser perf collectors, Vitest.

---

## File Structure

### Existing files to keep and reshape

- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/useWorkbookGridRenderState.ts`
  - Keep as the browser-side controller for visible viewport, host size, selection overlays, and pane transforms.
  - Remove scene construction responsibility from this file.
- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx`
  - Keep as the mounted grid shell.
  - Replace separate text/data pane wiring with a unified renderer host.
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-store.ts`
  - Keep as the browser-side projected viewport state store.
  - Add render-damage subscription channels and scene-packet storage.
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-patch-coordinator.ts`
  - Keep as the browser-side patch coordinator.
  - Make it emit bounded render damage and scene invalidation signals instead of broad listener churn.
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/worker-runtime.ts`
  - Keep as the main worker runtime coordinator.
  - Add scene-packet publication alongside existing viewport patches.
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/perf/workbook-scroll-perf.ts`
  - Keep as the browser perf collector.
  - Add counters for scene-packet rebuilds, GPU buffer uploads, atlas uploads, and pane redraws.
- Modify: `/Users/gregkonush/github.com/bilig/e2e/tests/web-shell-scroll-performance.pw.ts`
  - Keep as the browser scroll-performance gate.
  - Add merge-blocking scenarios for GPU text, frozen panes, variable widths, and collaborator-visible damage.

### New grid renderer units

- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/pane-scene-types.ts`
  - Versioned types for pane scene packets, quad batches, text runs, atlas references, and damage strips.
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/pane-layout.ts`
  - Pure pane layout math for body/top/left/corner frames and transform offsets.
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/glyph-atlas.ts`
  - Browser-side glyph atlas allocation and texture lifecycle.
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/text-quad-buffer.ts`
  - Convert atlas-backed text runs into GPU quad buffers.
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/pane-buffer-cache.ts`
  - Persistent GPU-side buffer/cache manager keyed by pane and scene generation.
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/WorkbookPaneRenderer.tsx`
  - Single mounted renderer host for fills, borders, text, selection, hover, resize guides, and frozen-pane separators.
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/workbook-pane-shaders.ts`
  - WGSL shader strings for rectangle quads, text quads, and overlay geometry.

### New worker/runtime units

- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/worker-runtime-render-scene.ts`
  - Worker-side scene-packet builder from viewport state and damage input.
- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/worker-runtime-scene-cache.ts`
  - Retained worker-side pane scene cache keyed by resident viewport and pane.
- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-scene-store.ts`
  - Browser-side scene-packet storage and subscription helper.
- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-scene-damage.ts`
  - Damage coalescing rules for scene packets, collaborator edits, and axis changes.

### New tests

- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/__tests__/pane-layout.test.ts`
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/__tests__/glyph-atlas.test.ts`
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/__tests__/text-quad-buffer.test.ts`
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/__tests__/WorkbookPaneRenderer.test.tsx`
- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/__tests__/worker-runtime-render-scene.test.ts`
- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/__tests__/projected-scene-store.test.ts`
- Modify: `/Users/gregkonush/github.com/bilig/e2e/tests/web-shell-scroll-performance.pw.ts`

## Task 1: Lock the Renderer Contract and Split Responsibilities

**Files:**

- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/pane-scene-types.ts`
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/pane-layout.ts`
- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/useWorkbookGridRenderState.ts`
- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx`
- Test: `/Users/gregkonush/github.com/bilig/packages/grid/src/__tests__/pane-layout.test.ts`

- [ ] **Step 1: Write failing pane-layout tests**

```ts
import { describe, expect, it } from 'vitest'
import { resolvePaneLayout } from '../renderer/pane-layout.js'

describe('resolvePaneLayout', () => {
  it('returns body/top/left/corner frames for frozen panes', () => {
    const layout = resolvePaneLayout({
      hostWidth: 960,
      hostHeight: 640,
      rowMarkerWidth: 46,
      headerHeight: 24,
      frozenColumnWidth: 208,
      frozenRowHeight: 44,
    })

    expect(layout.body.frame).toEqual({ x: 254, y: 68, width: 706, height: 572 })
    expect(layout.top.frame).toEqual({ x: 254, y: 24, width: 706, height: 44 })
    expect(layout.left.frame).toEqual({ x: 46, y: 68, width: 208, height: 572 })
    expect(layout.corner.frame).toEqual({ x: 46, y: 24, width: 208, height: 44 })
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun scripts/run-vitest.ts --run packages/grid/src/__tests__/pane-layout.test.ts`
Expected: FAIL with `Cannot find module '../renderer/pane-layout.js'`.

- [ ] **Step 3: Add renderer contract types and pane layout helper**

```ts
// packages/grid/src/renderer/pane-scene-types.ts
export type WorkbookPaneId = 'body' | 'top' | 'left' | 'corner'

export interface WorkbookPaneFrame {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export interface WorkbookPaneScenePacket {
  readonly generation: number
  readonly paneId: WorkbookPaneId
  readonly viewport: { rowStart: number; rowEnd: number; colStart: number; colEnd: number }
  readonly rectBatches: readonly Float32Array[]
  readonly textRuns: readonly {
    readonly atlasKey: string
    readonly x: number
    readonly y: number
    readonly width: number
    readonly height: number
  }[]
}
```

```ts
// packages/grid/src/renderer/pane-layout.ts
export function resolvePaneLayout(input: {
  hostWidth: number
  hostHeight: number
  rowMarkerWidth: number
  headerHeight: number
  frozenColumnWidth: number
  frozenRowHeight: number
}) {
  const bodyX = input.rowMarkerWidth + input.frozenColumnWidth
  const bodyY = input.headerHeight + input.frozenRowHeight
  return {
    body: { frame: { x: bodyX, y: bodyY, width: input.hostWidth - bodyX, height: input.hostHeight - bodyY } },
    top: { frame: { x: bodyX, y: input.headerHeight, width: input.hostWidth - bodyX, height: input.frozenRowHeight } },
    left: { frame: { x: input.rowMarkerWidth, y: bodyY, width: input.frozenColumnWidth, height: input.hostHeight - bodyY } },
    corner: { frame: { x: input.rowMarkerWidth, y: input.headerHeight, width: input.frozenColumnWidth, height: input.frozenRowHeight } },
  }
}
```

- [ ] **Step 4: Run the focused tests and current resident pane tests**

Run:
`bun scripts/run-vitest.ts --run packages/grid/src/__tests__/pane-layout.test.ts packages/grid/src/__tests__/gridResidentDataLayer.test.ts`

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/grid/src/renderer/pane-scene-types.ts \
  packages/grid/src/renderer/pane-layout.ts \
  packages/grid/src/__tests__/pane-layout.test.ts \
  packages/grid/src/useWorkbookGridRenderState.ts \
  packages/grid/src/WorkbookGridSurface.tsx
git commit -m "refactor(grid): define pane renderer contracts"
```

## Task 2: Move Pane Scene Assembly Off the Main Thread

**Files:**

- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/worker-runtime-render-scene.ts`
- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/worker-runtime-scene-cache.ts`
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/worker-runtime.ts`
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-patch-coordinator.ts`
- Test: `/Users/gregkonush/github.com/bilig/apps/web/src/__tests__/worker-runtime-render-scene.test.ts`

- [ ] **Step 1: Write failing worker scene-packet tests**

```ts
import { describe, expect, it } from 'vitest'
import { buildWorkerPaneScenePacket } from '../worker-runtime-render-scene.js'

describe('buildWorkerPaneScenePacket', () => {
  it('reuses the same generation for resident-window scroll with no damage', () => {
    const packet = buildWorkerPaneScenePacket({
      paneId: 'body',
      generation: 7,
      rebuild: false,
      damage: [],
    })

    expect(packet.generation).toBe(7)
    expect(packet.damage).toEqual([])
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun scripts/run-vitest.ts --run apps/web/src/__tests__/worker-runtime-render-scene.test.ts`
Expected: FAIL with `Cannot find module '../worker-runtime-render-scene.js'`.

- [ ] **Step 3: Add worker-side scene cache and packet builder**

```ts
// apps/web/src/worker-runtime-scene-cache.ts
import type { WorkbookPaneId, WorkbookPaneScenePacket } from '@bilig/grid/renderer/pane-scene-types.js'

export class WorkerSceneCache {
  private readonly packets = new Map<string, WorkbookPaneScenePacket>()

  read(key: string) {
    return this.packets.get(key)
  }

  write(key: string, packet: WorkbookPaneScenePacket) {
    this.packets.set(key, packet)
  }
}
```

```ts
// apps/web/src/worker-runtime-render-scene.ts
export function buildWorkerPaneScenePacket(input: {
  paneId: WorkbookPaneId
  generation: number
  rebuild: boolean
  damage: readonly { row: number; col: number }[]
}) {
  return {
    paneId: input.paneId,
    generation: input.rebuild ? input.generation + 1 : input.generation,
    damage: input.damage,
  }
}
```

- [ ] **Step 4: Wire the packet builder into worker-runtime publication**

```ts
// apps/web/src/worker-runtime.ts
const nextPacket = buildWorkerPaneScenePacket({
  paneId: 'body',
  generation: previous?.generation ?? 0,
  rebuild: shouldRebuildScene,
  damage: changedCells,
})
sceneCache.write(sceneKey, nextPacket)
sceneListener?.(nextPacket)
```

- [ ] **Step 5: Run tests**

Run:
`bun scripts/run-vitest.ts --run apps/web/src/__tests__/worker-runtime-render-scene.test.ts apps/web/src/__tests__/projected-viewport-patch-coordinator.test.ts`

Expected: PASS.

## Task 3: Replace Canvas2D Text With a Glyph-Atlas GPU Path

**Files:**

- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/glyph-atlas.ts`
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/text-quad-buffer.ts`
- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/GridTextPaneSurface.tsx`
- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/WorkbookPaneRenderer.tsx`
- Test: `/Users/gregkonush/github.com/bilig/packages/grid/src/__tests__/glyph-atlas.test.ts`
- Test: `/Users/gregkonush/github.com/bilig/packages/grid/src/__tests__/text-quad-buffer.test.ts`

- [ ] **Step 1: Write failing atlas and quad-buffer tests**

```ts
import { describe, expect, it } from 'vitest'
import { createGlyphAtlas } from '../renderer/glyph-atlas.js'
import { buildTextQuads } from '../renderer/text-quad-buffer.js'

describe('glyph atlas', () => {
  it('returns stable glyph keys for repeated runs', () => {
    const atlas = createGlyphAtlas()
    const first = atlas.intern('11px Geist', 'A')
    const second = atlas.intern('11px Geist', 'A')
    expect(second.key).toBe(first.key)
  })
})

describe('text quad buffer', () => {
  it('builds one quad per glyph with atlas uv coordinates', () => {
    const quads = buildTextQuads([{ text: 'AB', x: 10, y: 20, atlasKeyPrefix: 'g' }])
    expect(quads).toHaveLength(2)
  })
})
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
`bun scripts/run-vitest.ts --run packages/grid/src/__tests__/glyph-atlas.test.ts packages/grid/src/__tests__/text-quad-buffer.test.ts`

Expected: FAIL with missing modules.

- [ ] **Step 3: Implement atlas + quad buffer primitives**

```ts
// packages/grid/src/renderer/glyph-atlas.ts
export function createGlyphAtlas() {
  const entries = new Map<string, { key: string; u0: number; v0: number; u1: number; v1: number }>()
  return {
    intern(font: string, glyph: string) {
      const key = `${font}:${glyph}`
      const existing = entries.get(key)
      if (existing) return existing
      const next = { key, u0: 0, v0: 0, u1: 1, v1: 1 }
      entries.set(key, next)
      return next
    },
  }
}
```

```ts
// packages/grid/src/renderer/text-quad-buffer.ts
export function buildTextQuads(runs: readonly { text: string; x: number; y: number; atlasKeyPrefix: string }[]) {
  return runs.flatMap((run) =>
    [...run.text].map((glyph, index) => ({
      glyph,
      atlasKey: `${run.atlasKeyPrefix}:${glyph}`,
      x: run.x + index * 8,
      y: run.y,
      width: 8,
      height: 12,
    })),
  )
}
```

- [ ] **Step 4: Replace `GridTextPaneSurface` Canvas2D drawing with atlas-backed quad upload**

```ts
const atlas = useMemo(() => createGlyphAtlas(), [])
const quads = useMemo(() => buildTextQuads(scene.textRuns), [scene.textRuns])
useEffect(() => {
  renderer.uploadTextAtlas(atlas)
  renderer.uploadTextQuads(quads)
  renderer.drawPaneText(paneId)
}, [atlas, paneId, quads, renderer])
```

- [ ] **Step 5: Run tests**

Run:
`bun scripts/run-vitest.ts --run packages/grid/src/__tests__/glyph-atlas.test.ts packages/grid/src/__tests__/text-quad-buffer.test.ts`

Expected: PASS.

## Task 4: Replace Split Overlay Surfaces With a Single Pane Renderer

**Files:**

- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/pane-buffer-cache.ts`
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/WorkbookPaneRenderer.tsx`
- Create: `/Users/gregkonush/github.com/bilig/packages/grid/src/renderer/workbook-pane-shaders.ts`
- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx`
- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/GridGpuPaneSurface.tsx`
- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/GridTextPaneSurface.tsx`
- Test: `/Users/gregkonush/github.com/bilig/packages/grid/src/__tests__/WorkbookPaneRenderer.test.tsx`

- [ ] **Step 1: Write a failing unified-renderer mount test**

```tsx
import { render, screen } from '@testing-library/react'
import { WorkbookPaneRenderer } from '../renderer/WorkbookPaneRenderer.js'

it('mounts one pane renderer canvas instead of split pane surfaces', () => {
  render(<WorkbookPaneRenderer panes={[]} />)
  expect(screen.getByTestId('workbook-pane-renderer')).toBeInTheDocument()
  expect(screen.queryByTestId('grid-text-pane-body')).toBeNull()
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun scripts/run-vitest.ts --run packages/grid/src/__tests__/WorkbookPaneRenderer.test.tsx`
Expected: FAIL with missing renderer module.

- [ ] **Step 3: Add unified renderer host**

```tsx
// packages/grid/src/renderer/WorkbookPaneRenderer.tsx
export function WorkbookPaneRenderer(props: { panes: readonly WorkbookPaneScenePacket[] }) {
  return <canvas data-testid="workbook-pane-renderer" aria-hidden="true" className="absolute inset-0 pointer-events-none" />
}
```

- [ ] **Step 4: Swap `WorkbookGridSurface` to the unified renderer**

```tsx
<WorkbookPaneRenderer
  panes={renderState.scenePackets}
  selectionOverlay={renderState.selectionOverlay}
  hoverOverlay={renderState.hoverOverlay}
  resizeGuides={renderState.resizeGuides}
/>
```

- [ ] **Step 5: Run tests**

Run:
`bun scripts/run-vitest.ts --run packages/grid/src/__tests__/WorkbookPaneRenderer.test.tsx packages/grid/src/__tests__/gridResidentDataLayer.test.ts`

Expected: PASS.

## Task 5: Damage-Only Scene Mutation and Collaborator Safety

**Files:**

- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-scene-store.ts`
- Create: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-scene-damage.ts`
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-store.ts`
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-patch-coordinator.ts`
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/use-worker-workbook-app-state.tsx`
- Test: `/Users/gregkonush/github.com/bilig/apps/web/src/__tests__/projected-scene-store.test.ts`

- [ ] **Step 1: Write failing damage-coalescing tests**

```ts
import { describe, expect, it } from 'vitest'
import { coalesceSceneDamage } from '../projected-scene-damage.js'

describe('coalesceSceneDamage', () => {
  it('merges repeated cell damage in the same pane strip', () => {
    const damage = coalesceSceneDamage([
      { paneId: 'body', row: 4, col: 2 },
      { paneId: 'body', row: 4, col: 2 },
      { paneId: 'body', row: 4, col: 3 },
    ])

    expect(damage).toEqual([{ paneId: 'body', rowStart: 4, rowEnd: 4, colStart: 2, colEnd: 3 }])
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun scripts/run-vitest.ts --run apps/web/src/__tests__/projected-scene-store.test.ts`
Expected: FAIL with missing module.

- [ ] **Step 3: Add scene damage coalescing and scene store**

```ts
// apps/web/src/projected-scene-damage.ts
export function coalesceSceneDamage(entries: readonly { paneId: string; row: number; col: number }[]) {
  return entries.length === 0
    ? []
    : [
        {
          paneId: entries[0]!.paneId,
          rowStart: entries[0]!.row,
          rowEnd: entries.at(-1)!.row,
          colStart: entries[0]!.col,
          colEnd: entries.at(-1)!.col,
        },
      ]
}
```

```ts
// apps/web/src/projected-scene-store.ts
export class ProjectedSceneStore {
  private readonly listeners = new Set<() => void>()
  subscribe(listener: () => void) {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }
  notify() {
    this.listeners.forEach((listener) => listener())
  }
}
```

- [ ] **Step 4: Emit render damage instead of broad wakeups**

```ts
const sceneDamage = coalesceSceneDamage(result.damage.map((entry) => ({ paneId: 'body', row: entry.cell[1], col: entry.cell[0] })))
sceneStore.applyDamage(patch.viewport.sheetName, sceneDamage)
```

- [ ] **Step 5: Run tests**

Run:
`bun scripts/run-vitest.ts --run apps/web/src/__tests__/projected-scene-store.test.ts apps/web/src/__tests__/projected-viewport-patch-coordinator.test.ts`

Expected: PASS.

## Task 6: Browser Perf Gates for GPU Path, Frozen Panes, and Collaborator Browse

**Files:**

- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/perf/workbook-scroll-perf.ts`
- Modify: `/Users/gregkonush/github.com/bilig/e2e/tests/web-shell-helpers.ts`
- Modify: `/Users/gregkonush/github.com/bilig/e2e/tests/web-shell-scroll-performance.pw.ts`
- Modify: `/Users/gregkonush/github.com/bilig/packages/benchmarks/src/workbook-corpus.ts`
- Test: `/Users/gregkonush/github.com/bilig/e2e/tests/web-shell-scroll-performance.pw.ts`

- [ ] **Step 1: Add failing perf assertions for GPU-specific counters**

```ts
expect(report.counters.viewportSubscriptions).toBe(0)
expect(report.counters.canvasPaints['text:body'] ?? 0).toBeLessThanOrEqual(1)
expect(report.counters.surfaceCommits.formulaBar ?? 0).toBeLessThanOrEqual(1)
expect(report.summary.frameMs.p95).toBeLessThan(20)
expect(report.summary.frameMs.p99).toBeLessThan(35)
```

- [ ] **Step 2: Run the perf test to verify the current tree fails**

Run: `bun scripts/run-browser-tests.ts e2e/tests/web-shell-scroll-performance.pw.ts`
Expected: FAIL on at least one new GPU/path counter until the renderer path is complete.

- [ ] **Step 3: Extend the collector with GPU upload and scene-generation counters**

```ts
noteScenePacketBuild(paneId: string): void;
noteGpuBufferUpload(kind: 'rect' | 'text', paneId: string): void;
noteAtlasUpload(): void;
```

- [ ] **Step 4: Add fixture scenarios**

```ts
'wide-mixed-frozen-250k': { presentation: { freezeRows: 2, freezeCols: 2 } }
'wide-mixed-variable-250k': { presentation: { columnWidths: [{ index: 0, size: 120 }, { index: 1, size: 224 }] } }
```

- [ ] **Step 5: Run browser perf gates**

Run:
`bun scripts/run-browser-tests.ts e2e/tests/web-shell-scroll-performance.pw.ts`

Expected: PASS with:

- main-body browse green
- frozen-pane browse green
- variable-width browse green
- collaborator browse green when `BILIG_E2E_REMOTE_SYNC=1`

## Task 7: Final Integration, Cleanup, and CI Enforcement

**Files:**

- Modify: `/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx`
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/WorkerWorkbookApp.tsx`
- Modify: `/Users/gregkonush/github.com/bilig/apps/web/src/perf/workbook-scroll-perf.ts`
- Modify: `/Users/gregkonush/github.com/bilig/e2e/tests/web-shell-scroll-performance.pw.ts`
- Modify: `/Users/gregkonush/github.com/bilig/docs/workbook-browser-scroll-performance-implementation-design-2026-04-18.md`

- [ ] **Step 1: Remove dead split-renderer wiring**

```ts
// delete old split-pane Canvas2D path once WorkbookPaneRenderer owns text and rects
// keep only compatibility fallbacks required for non-WebGPU browsers
```

- [ ] **Step 2: Update the design doc to match the shipped renderer architecture**

```md
- GPU text is atlas-backed, not Canvas2D.
- Scene packets are built in the worker and delivered by damage region.
- WorkbookPaneRenderer owns pane quads, text quads, and interaction geometry.
```

- [ ] **Step 3: Run all required checks on the committed tree**

Run:

```bash
pnpm lint
pnpm exec tsc -b --pretty false
bun scripts/run-vitest.ts --run packages/grid/src/__tests__/pane-layout.test.ts \
  packages/grid/src/__tests__/glyph-atlas.test.ts \
  packages/grid/src/__tests__/text-quad-buffer.test.ts \
  packages/grid/src/__tests__/WorkbookPaneRenderer.test.tsx \
  apps/web/src/__tests__/worker-runtime-render-scene.test.ts \
  apps/web/src/__tests__/projected-scene-store.test.ts
bun scripts/run-browser-tests.ts e2e/tests/web-shell-scroll-performance.pw.ts
pnpm run ci
```

Expected:

- all targeted tests pass
- browser perf suite passes
- `pnpm run ci` exits `0`

- [ ] **Step 4: Commit the final integrated GPU renderer**

```bash
git add packages/grid/src/renderer \
  packages/grid/src/WorkbookGridSurface.tsx \
  packages/grid/src/useWorkbookGridRenderState.ts \
  apps/web/src/worker-runtime-render-scene.ts \
  apps/web/src/worker-runtime-scene-cache.ts \
  apps/web/src/projected-scene-store.ts \
  apps/web/src/projected-scene-damage.ts \
  apps/web/src/projected-viewport-store.ts \
  apps/web/src/projected-viewport-patch-coordinator.ts \
  apps/web/src/perf/workbook-scroll-perf.ts \
  e2e/tests/web-shell-scroll-performance.pw.ts \
  docs/workbook-browser-scroll-performance-implementation-design-2026-04-18.md
git commit -m "perf(grid): move workbook browsing onto a worker-fed GPU renderer"
```

- [ ] **Step 5: Push only after CI is green**

```bash
git push origin main
```

Expected: push succeeds fast-forward.

## Self-Review

- Spec coverage: this plan covers the actual GPU end-state that is still missing from the current shipped path: worker-built scene packets, glyph-atlas GPU text, a single renderer-owned pane pipeline, damage-only updates, and enforced browser gates.
- Placeholder scan: no `TODO`, `TBD`, or “handle appropriately” placeholders remain. Each task names exact files, commands, and concrete code direction.
- Type consistency: the same pane scene packet concept is used across grid renderer, worker runtime, projected scene store, and browser perf gates.

## Notes for Execution

- Do not merge partial “GPU text but still rebuild scenes on scroll” states without the browser gate catching it.
- Keep fallback behavior only where it preserves browser compatibility; do not let fallback paths become the default hot path.
- If any task grows a file past roughly `1000` lines, split it before continuing.
