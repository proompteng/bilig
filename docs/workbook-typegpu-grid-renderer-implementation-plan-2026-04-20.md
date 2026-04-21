# Workbook TypeGPU Grid Renderer Implementation Plan

Date: `2026-04-20`

Status: `ready for implementation`

Grounding:

- current mounted workbook grid in `packages/grid`
- current TypeGPU renderer in `packages/grid/src/renderer`
- current worker-backed viewport/runtime path in `apps/web`
- current browser scroll perf harness in `e2e/tests/web-shell-scroll-performance.pw.ts`
- current TypeGPU visual readback harness in `e2e/tests/web-shell-typegpu.pw.ts`

Related documents:

- `docs/workbook-browser-scroll-performance-implementation-design-2026-04-18.md`
- `docs/glide-removal-renderer-v2.md`
- `docs/performance-budgets.md`
- `docs/react-spectrum-ui-philosophy.md`
- `docs/testing-and-benchmarks.md`

## Purpose

This document defines the implementation plan for replacing the current workbook grid
rendering path with a retained, accurate, high-performance WebGPU renderer built through
TypeGPU.

The target is not a small optimization pass. The target is a game-engine style grid:

- scroll behaves like camera movement through a large retained world
- GPU resources survive scroll frames
- workers prepare scene data
- the main thread submits frames
- React owns semantics, not per-frame motion
- visual accuracy is proven by readback tests
- performance is enforced by budgets, not by subjective feel

The plan is complete when the workbook grid can scroll vertically, horizontally, and
diagonally across large sheets with stable frame pacing, no visible jitter, no layout shift,
and accurate text/header/selection rendering.

## Current Starting Point

The current implementation already has useful pieces, but they are not assembled as a
renderer engine.

Existing useful pieces:

- `WorkbookGridSurface.tsx` owns the mounted workbook grid surface.
- `useWorkbookGridRenderState.ts` resolves viewport, selection, overlay, and resident pane
  render state.
- `WorkbookPaneRenderer.tsx` mounts a TypeGPU canvas and draws panes.
- `typegpu-renderer.ts` owns TypeGPU shader, pipeline, buffer, uniform, and atlas helpers.
- `gridResidentDataLayer.ts` can build resident pane scenes.
- `apps/web/src/worker-runtime-render-scene.ts` can build resident scenes in the worker.
- `projected-scene-store.ts` can subscribe to worker resident pane scene packets.
- `web-shell-scroll-performance.pw.ts` and `web-shell-typegpu.pw.ts` provide the right
  browser-level test surfaces.

Current blockers:

- `resolveColumnAnchor` and `resolveRowAnchor` linearly scan axes from zero to the target
  scroll offset. Vertical scroll can degrade badly near deep rows.
- Scroll state is synchronized through one `requestAnimationFrame`, then the renderer draws
  through another `requestAnimationFrame`. This creates avoidable frame latency.
- Canvas sizing and `GPUCanvasContext.configure()` still live inside the draw path.
- Worker resident scenes are subscribed, but local main-thread scene building is still the
  primary render source.
- Scene payloads are object-heavy and not shaped for transferable, typed render packets.
- Text layout and GPU text atlas behavior are not yet a single authoritative layout system.
- The perf collector does not yet prove TypeGPU submit/upload/configure stability.
- Visual accuracy tests prove only a narrow text-visibility regression, not full grid
  fidelity.

## Final Architecture

The renderer is split into five explicit systems.

### 1. Workbook Semantics

Owner: React and workbook state.

Responsibilities:

- active sheet
- selected cell and selection range
- editing state
- row/column sizes and hidden state
- freeze panes
- style and value revisions
- accessibility and keyboard semantics

Non-responsibilities:

- per-frame scroll camera updates
- TypeGPU resource lifetime
- steady-state pane draw scheduling
- scene buffer upload during scroll

### 2. Axis Engine

Owner: `AxisIndex`.

Responsibilities:

- offset to row/column anchor
- row/column index to offset
- hidden axis entries
- sparse size overrides
- prefix sums
- binary search
- viewport restore
- pointer hit testing support

This is the source of truth for scroll math.

### 3. Grid Camera

Owner: `GridCamera`.

Responsibilities:

- raw scroll position
- resolved anchor row and column
- intra-cell offsets
- visible viewport
- resident tile window
- DPR
- host size
- frozen pane offsets
- velocity and direction
- tile prefetch hints

The camera is updated directly from scroll input and sampled by the render scheduler.

### 4. Scene Residency

Owner: worker-backed `GridTileResidency`.

Responsibilities:

- visible tile
- warm neighbor tiles
- body pane tile
- top frozen strip tile
- left frozen strip tile
- corner tile
- tile revisions
- stale-but-valid scene fallback
- tile miss accounting

Workers build typed render packets. The main thread consumes packets.

### 5. TypeGPU Render Backend

Owner: renderer modules under `packages/grid/src/renderer`.

Responsibilities:

- WebGPU device/context lifecycle
- canvas configure on init/resize/DPR only
- persistent pipelines
- persistent buffers
- persistent bind groups
- glyph atlas texture
- camera uniforms
- pane scissor rects
- render pass encoding
- GPU counters

The steady scroll frame updates camera uniforms and submits. It does not rebuild scenes,
resize the canvas, configure the context, or allocate buffers.

## Success Criteria

These are implementation gates, not aspirations.

### Steady In-Tile Scroll

During a warmed scroll inside an already resident tile:

- `0` React commits outside the renderer bridge
- `0` worker viewport resubscriptions
- `0` worker scene rebuilds
- `0` `GPUCanvasContext.configure()` calls
- `0` GPU buffer allocations
- `0` atlas uploads
- camera uniform upload only
- p95 frame interval `< 16.7ms`
- p99 frame interval `< 24ms`
- no long task over `50ms`
- no layout shift
- frozen panes remain visually locked
- selection and fill handle remain attached to cells

### Tile Boundary Scroll

When crossing a resident tile boundary:

- visible tile is already available or a stale valid tile renders for at most one frame
- no blank panes
- no missing text
- no visible jump
- tile miss count stays under the scenario budget
- scene packet refreshes are bounded by crossed tile count

### Visual Accuracy

TypeGPU readback tests must prove:

- grid line alignment within `1` physical pixel
- text clip rects within `1` physical pixel
- column header labels stay visible and aligned after horizontal scroll
- row header labels stay visible and aligned after vertical scroll
- body text stays visible and clipped correctly after diagonal scroll
- active cell border aligns with axis math
- range selection fill and border align with cell bounds
- frozen rows and columns align with the body pane
- wide text, clipped text, wrapped text, bold text, italic text, underline, and number text
  match the shared text layout model

## Implementation Workstreams

## Workstream 1: Renderer Contract And Counters

Files to add:

- `packages/grid/src/renderer/grid-render-contract.ts`
- `packages/grid/src/renderer/grid-render-counters.ts`

Types to define:

- `GridCameraSnapshot`
- `GridAxisSnapshot`
- `GridRenderFrame`
- `GridRenderStats`
- `GridGpuCounters`
- `GridResidentTileKey`
- `GridResidentScenePacket`
- `GridPaneRenderPacket`
- `GridTextLayoutPacket`

Counter API:

- `noteTypeGpuConfigure()`
- `noteTypeGpuSubmit()`
- `noteTypeGpuDrawCall(count)`
- `noteTypeGpuUniformWrite(bytes, label)`
- `noteTypeGpuBufferWrite(bytes, label)`
- `noteTypeGpuBufferAllocation(bytes, label)`
- `noteTypeGpuAtlasUpload(bytes)`
- `noteTypeGpuSurfaceResize(width, height, dpr)`
- `noteTypeGpuTileMiss(tileKey)`
- `noteTypeGpuScenePacketApplied(packetKey)`

Perf collector changes:

- Extend `apps/web/src/perf/workbook-scroll-perf.ts`.
- Add GPU counters to `WorkbookScrollPerfCounters`.
- Add input-to-draw timing samples.
- Add per-frame submit and upload counters.
- Export raw counter samples to Playwright artifacts.

Acceptance:

- Existing scroll perf tests keep passing.
- New counters appear in the JSON artifact.
- TypeGPU path increments `gpu:*` counters directly from renderer code.
- Canvas fallback counters remain separate from TypeGPU counters.

## Workstream 2: AxisIndex

Files to add:

- `packages/grid/src/gridAxisIndex.ts`
- `packages/grid/src/__tests__/gridAxisIndex.test.ts`

Files to change:

- `packages/grid/src/workbookGridViewport.ts`
- `packages/grid/src/gridPointer.ts`
- `packages/grid/src/useWorkbookGridRenderState.ts`

`AxisIndex` API:

```ts
export interface AxisEntryOverride {
  readonly index: number
  readonly size: number
  readonly hidden?: boolean
}

export interface AxisAnchor {
  readonly index: number
  readonly offset: number
}

export interface AxisIndex {
  readonly axisLength: number
  readonly defaultSize: number
  resolveOffset(index: number): number
  resolveAnchor(scrollOffset: number): AxisAnchor
  resolveSize(index: number): number
  resolveSpan(start: number, endExclusive: number): number
  resolveVisibleCount(start: number, viewportSize: number, overscanPx: number): number
}
```

Implementation rules:

- Fast path for no overrides: direct division and multiplication.
- Sparse overrides are sorted once per axis revision.
- Prefix sums store cumulative delta from default size.
- Hidden rows/columns are represented as zero-size overrides.
- Anchor lookup uses binary search over override ranges.
- All results are clamped to protocol max rows/columns.

Tests:

- default column offset/index round trip
- default row offset/index round trip near `MAX_ROWS`
- variable column widths
- variable row heights
- hidden rows
- hidden columns
- frozen panes
- viewport restore
- fuzz tests for small axis lengths against a simple linear oracle

Acceptance:

- No scroll path loops over rows or columns from zero to the target index.
- Vertical scrolling near row `1_048_576` does not degrade relative to row `1`.
- Existing viewport tests pass.

## Workstream 3: GridCamera

Files to add:

- `packages/grid/src/gridCamera.ts`
- `packages/grid/src/__tests__/gridCamera.test.ts`

Responsibilities:

- Convert native `scrollLeft` and `scrollTop` to axis anchors.
- Maintain visible viewport.
- Maintain resident viewport.
- Maintain frozen pane offsets.
- Maintain camera offsets in pane-local coordinates.
- Track velocity and direction.
- Emit semantic changes only when the resident window changes.

`GridCamera` API:

```ts
export interface GridCameraInput {
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly viewportWidth: number
  readonly viewportHeight: number
  readonly dpr: number
}

export interface GridCameraSnapshot {
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly tx: number
  readonly ty: number
  readonly visibleViewport: Viewport
  readonly residentViewport: Viewport
  readonly velocityX: number
  readonly velocityY: number
  readonly dpr: number
}
```

Acceptance:

- `GridCamera` replaces ad hoc scroll transform calculation.
- `getCellLocalBounds` and renderer pane offsets use the same camera snapshot.
- Editor overlay, fill handle, and selection use the same camera source as the canvas.

## Workstream 4: Single Frame Scheduler

Files to add:

- `packages/grid/src/renderer/grid-render-scheduler.ts`
- `packages/grid/src/__tests__/grid-render-scheduler.test.ts`

Files to change:

- `packages/grid/src/useWorkbookGridRenderState.ts`
- `packages/grid/src/renderer/WorkbookPaneRenderer.tsx`
- `packages/grid/src/WorkbookGridSurface.tsx`

Scheduler behavior:

1. Scroll event fires.
2. Scroll handler updates `GridCamera` immediately.
3. Scheduler records input timestamp and marks frame dirty.
4. One `requestAnimationFrame` draws the latest camera snapshot.
5. React state updates only if the resident tile window changed or interaction state requires
   semantic React updates.

Remove:

- scroll sync RAF followed by renderer draw RAF
- renderer-owned RAF that is disconnected from the camera scheduler

Acceptance:

- Input-to-draw latency is at most one frame.
- Scroll and render share the same camera snapshot.
- Perf counter records the scroll event timestamp and draw timestamp.
- Existing overlays do not drift from canvas content.

## Workstream 5: Persistent TypeGPU Surface

Files to add:

- `packages/grid/src/renderer/typegpu-surface-manager.ts`

Files to change:

- `packages/grid/src/renderer/WorkbookPaneRenderer.tsx`
- `packages/grid/src/renderer/typegpu-renderer.ts`

Surface manager responsibilities:

- initialize device
- initialize canvas context
- configure context
- track canvas CSS size
- track canvas pixel size
- track DPR
- reconfigure only on init/resize/DPR change
- recover from device loss

Draw path rules:

- no canvas `width` assignment
- no canvas `height` assignment
- no `canvas.style.width` assignment
- no `canvas.style.height` assignment
- no `context.configure`

Acceptance:

- Perf test can assert `configureCount === 0` during steady scroll.
- Resize test can assert `configureCount > 0` only during resize.
- Surface size changes do not cause layout shift.

## Workstream 6: Persistent TypeGPU Resources

Files to add:

- `packages/grid/src/renderer/typegpu-resource-cache.ts`
- `packages/grid/src/renderer/typegpu-draw-pass.ts`

Files to change:

- `packages/grid/src/renderer/pane-buffer-cache.ts`
- `packages/grid/src/renderer/typegpu-renderer.ts`
- `packages/grid/src/renderer/WorkbookPaneRenderer.tsx`

Resource model:

- one immutable unit quad buffer
- persistent rect buffers per pane/tile
- persistent text buffers per pane/tile
- persistent decoration buffers per pane/tile
- persistent camera uniform buffers
- persistent pane bind groups
- persistent glyph atlas texture
- explicit resource version fields

Upload rules:

- scene buffers upload only when scene packet version changes
- camera uniform uploads on scroll
- atlas uploads only when atlas version changes
- capacity grows geometrically
- no capacity shrink during active session
- stale unused resources are pruned outside the scroll frame

Acceptance:

- `bufferAllocations === 0` during steady scroll.
- `vertexUploadBytes === 0` during steady in-tile scroll.
- `uniformWriteBytes` equals camera/pane uniform budget during scroll.
- Draw pass can render body, top, left, corner, and overlay panes from retained buffers.

## Workstream 7: Worker Typed Scene Packets

Files to add:

- `packages/grid/src/renderer/grid-scene-packet.ts`
- `apps/web/src/worker-runtime-render-packet.ts`

Files to change:

- `packages/grid/src/renderer/pane-scene-types.ts`
- `packages/grid/src/gridResidentDataLayer.ts`
- `apps/web/src/worker-runtime-render-scene.ts`
- `apps/web/src/projected-scene-store.ts`
- `apps/web/src/resident-pane-scene-types.ts`

Typed packet contents:

- packet key
- generation
- sheet name
- axis revision
- style revision
- value revision
- selection revision
- pane id
- tile key
- viewport
- surface size
- rect instance buffer
- text instance buffer
- decoration rect buffer
- optional debug metadata

Transfer rules:

- worker transfers `ArrayBuffer` ownership to main thread
- main thread validates packet shape
- packet identity determines buffer uploads
- old packet remains renderable while next packet is in flight

Acceptance:

- Worker scene packets are the primary render input.
- Main-thread local scene build is fallback only.
- In-tile scroll does not request new packets.
- Tile-boundary scroll requests bounded packets.

## Workstream 8: Tile Residency And Streaming

Files to add:

- `packages/grid/src/gridTileResidency.ts`
- `packages/grid/src/__tests__/gridTileResidency.test.ts`

Tile model:

- body tile
- top frozen strip tile
- left frozen strip tile
- corner tile
- warm neighbor tiles

Policy:

- keep visible tile resident
- prefetch the next tile in scroll direction
- prefetch diagonal neighbor during diagonal motion
- use hysteresis before tile swap
- never block render waiting for a tile
- stale valid tile can remain for one frame
- blank tile is a test failure

Acceptance:

- Tile boundary tests have no blank readback regions.
- Tile miss counter stays under budget.
- Scene packet refresh count is proportional to tile crossings, not frames.

## Workstream 9: Shared Text Layout And Atlas Accuracy

Files to add:

- `packages/grid/src/renderer/gridTextLayout.ts`
- `packages/grid/src/__tests__/gridTextLayout.test.ts`

Files to change:

- `packages/grid/src/gridTextScene.ts`
- `packages/grid/src/GridTextOverlay.tsx`
- `packages/grid/src/renderer/text-quad-buffer.ts`
- `packages/grid/src/renderer/glyph-atlas.ts`
- `packages/grid/src/renderer/typegpu-renderer.ts`

Layout requirements:

- one authoritative text layout output for TypeGPU and canvas fallback
- grapheme-aware segmentation where text is split
- exact clip rects
- exact align behavior
- exact vertical baseline behavior
- wrapped text support
- underline and strike support
- high-DPR snapping rules
- overflow behavior for adjacent empty cells where supported
- deterministic atlas versioning

Acceptance:

- Canvas fallback and TypeGPU consume the same layout packets.
- TypeGPU readback proves text remains visible and correctly clipped after scroll.
- Wide labels and narrow clipped cells match the layout model.
- Atlas growth cannot occur inside steady scroll after warmup.

## Workstream 10: Explicit Render Layers

Files to add:

- `packages/grid/src/renderer/grid-render-layers.ts`

Layer order:

1. sheet background
2. cell fills
3. grid lines
4. conditional and semantic fills
5. selection fill
6. body text
7. text decorations
8. cell borders
9. active cell border
10. fill handle
11. resize guides
12. frozen pane separators
13. editor and accessibility DOM overlays

Acceptance:

- Layer order is declared in code, not incidental.
- Body, header, frozen, and overlay panes use the same ordering rules.
- Active cell and selection readback are stable after scroll and resize.

## Workstream 11: React Boundary Refactor

Files to add:

- `packages/grid/src/useGridCameraState.ts`
- `packages/grid/src/useGridSelectionState.ts`
- `packages/grid/src/useGridResizeState.ts`
- `packages/grid/src/useGridOverlayState.ts`
- `packages/grid/src/useGridSceneResidency.ts`
- `packages/grid/src/useGridViewportSubscriptions.ts`

Files to change:

- `packages/grid/src/useWorkbookGridRenderState.ts`
- `packages/grid/src/WorkbookGridSurface.tsx`

Refactor rules:

- keep public `useWorkbookGridRenderState` API stable during extraction
- extract without changing behavior first
- move camera ownership out of the large hook
- keep selection/editing semantics in React
- keep per-frame render scheduling outside React

Acceptance:

- `useWorkbookGridRenderState.ts` no longer owns renderer frame pacing.
- No source file grows beyond the repository size guideline.
- Tests cover the extracted hooks or pure helpers.

## Workstream 12: Browser Perf Gates

Files to change:

- `e2e/tests/web-shell-scroll-performance.pw.ts`
- `e2e/tests/web-shell-helpers.ts`
- `apps/web/src/perf/workbook-scroll-perf.ts`

Workloads:

- `wide-horizontal-steady`
- `deep-vertical-steady`
- `diagonal-steady`
- `inertial-wheel-scroll`
- `variable-row-height-vertical`
- `variable-column-width-horizontal`
- `frozen-pane-diagonal`
- `tile-boundary-horizontal`
- `tile-boundary-vertical`
- `editing-overlay-scroll`
- `resize-then-scroll`

Assertions:

- p95 frame interval `< 16.7ms`
- p99 frame interval `< 24ms`
- no long task over `50ms`
- no React commits during steady scroll
- no context configure during steady scroll
- no buffer allocation during steady scroll
- no vertex upload during in-tile scroll
- no scene rebuild during in-tile scroll
- tile miss budget respected during boundary tests
- frozen panes remain aligned

Artifacts:

- raw frame samples
- raw long task samples
- GPU counters
- scene packet counters
- tile counters
- browser viewport
- fixture id
- device scale factor
- readback artifact path where relevant

## Workstream 13: TypeGPU Accuracy Gates

Files to change:

- `e2e/tests/web-shell-typegpu.pw.ts`

Readback scenarios:

- initial grid render
- vertical scroll text and row header visibility
- horizontal scroll text and column header visibility
- diagonal scroll body/header/frozen alignment
- wide clipped labels
- narrow clipped labels
- wrapped text
- right-aligned numbers
- selected cell border
- selected range fill and border
- frozen row and column separators
- high-DPR render
- resize then scroll
- tile boundary crossing

Acceptance:

- Tests assert concrete pixel/color/region evidence.
- Tests save readback artifacts for failures.
- Text/header/body regions do not rely on broad "dark pixel exists" checks only; final tests
  include geometry and alignment probes.

## Commit Plan

Use focused commits in this order:

1. `test(grid): add typegpu scroll counters and renderer artifacts`
2. `feat(grid): add axis index for fast viewport anchors`
3. `feat(grid): route scroll through grid camera scheduler`
4. `feat(grid): keep typegpu surface configuration outside draw`
5. `feat(grid): retain typegpu pane resources across scroll frames`
6. `feat(grid): use worker typed scene packets for grid panes`
7. `feat(grid): stream grid tiles with warm neighbor residency`
8. `fix(grid): unify text layout for typegpu and canvas fallback`
9. `refactor(grid): split workbook render state by camera and residency`
10. `test(grid): gate webgpu grid accuracy and scroll budgets`

Each commit must leave targeted tests green. The final committed tree must pass full CI.

## Verification Commands

Run targeted verification after each relevant workstream:

```bash
pnpm --filter @bilig/grid test
pnpm exec playwright test e2e/tests/web-shell-typegpu.pw.ts
pnpm exec playwright test e2e/tests/web-shell-scroll-performance.pw.ts
```

Final verification:

```bash
pnpm run ci
```

## Execution Checklist

- [ ] Add renderer contract and counters.
- [ ] Extend scroll perf collector with TypeGPU counters.
- [ ] Add `AxisIndex` and replace linear scroll anchor scans.
- [ ] Add `GridCamera`.
- [ ] Replace two-stage RAF scroll path with one render scheduler.
- [ ] Move TypeGPU configure and resize work out of draw.
- [ ] Add persistent TypeGPU resource cache.
- [ ] Promote worker scene packets to primary render input.
- [ ] Convert scene packets to transferable typed arrays.
- [ ] Add tile residency with warm neighbors and hysteresis.
- [ ] Unify text layout for TypeGPU and canvas fallback.
- [ ] Declare explicit render layer order.
- [ ] Split large render state hook into focused modules.
- [ ] Add vertical, horizontal, diagonal, inertial, frozen-pane, tile-boundary, and resize perf gates.
- [ ] Add TypeGPU readback accuracy tests.
- [ ] Run targeted tests.
- [ ] Run full `pnpm run ci`.

