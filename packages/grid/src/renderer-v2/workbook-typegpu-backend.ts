import type { WorkbookGridScrollSnapshot } from '../workbookGridScrollStore.js'
import type { WorkbookRenderPaneState } from '../renderer/pane-scene-types.js'
import { noteTypeGpuTileMiss } from '../renderer/grid-render-counters.js'
import { createGlyphAtlas } from './typegpu-atlas-manager.js'
import { WorkbookPaneBufferCache, type WorkbookPaneBufferEntry } from './pane-buffer-cache.js'
import {
  createTypeGpuRenderer,
  destroyTypeGpuRenderer,
  syncTypeGpuAtlasResources,
  type TypeGpuRendererArtifacts,
} from './typegpu-backend.js'
import { drawTypeGpuPanes, type TypeGpuDrawSurface } from './typegpu-render-pass.js'
import { syncTypeGpuPaneResources } from './typegpu-buffer-pool.js'
import { createTypeGpuSurfaceState, syncTypeGpuCanvasSurface, type TypeGpuSurfaceState } from './typegpu-surface.js'
import { buildTileGpuCacheKey, TileGpuCache, syncTileGpuCacheFromPanes } from './tile-gpu-cache.js'

export interface WorkbookTypeGpuBackend {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly surfaceState: TypeGpuSurfaceState
  readonly tileCache: TileGpuCache
}

export async function createWorkbookTypeGpuBackend(canvas: HTMLCanvasElement): Promise<WorkbookTypeGpuBackend | null> {
  const artifacts = await createTypeGpuRenderer(canvas)
  if (!artifacts) {
    return null
  }
  return {
    artifacts,
    atlas: createGlyphAtlas(),
    paneBuffers: new WorkbookPaneBufferCache(),
    surfaceState: createTypeGpuSurfaceState(),
    tileCache: new TileGpuCache(),
  }
}

export function destroyWorkbookTypeGpuBackend(backend: WorkbookTypeGpuBackend): void {
  backend.paneBuffers.dispose()
  destroyTypeGpuRenderer(backend.artifacts)
}

export function syncWorkbookTypeGpuSurface(input: {
  readonly backend: WorkbookTypeGpuBackend
  readonly canvas: HTMLCanvasElement
  readonly size: TypeGpuDrawSurface
}): void {
  syncTypeGpuCanvasSurface({
    artifacts: input.backend.artifacts,
    canvas: input.canvas,
    size: input.size,
    state: input.backend.surfaceState,
  })
}

export function drawWorkbookTypeGpuFrame(input: {
  readonly backend: WorkbookTypeGpuBackend
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly preloadPanes?: readonly WorkbookRenderPaneState[] | undefined
  readonly syncPreloadPanes?: boolean | undefined
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuDrawSurface
}): void {
  const retainPanes = input.preloadPanes?.length ? [...input.preloadPanes, ...input.panes] : input.panes
  const resourcePanes = input.syncPreloadPanes === false ? input.panes : retainPanes
  syncTileGpuCacheFromPanes({
    cache: input.backend.tileCache,
    panes: resourcePanes,
  })
  const staleRetainPanes = resolveTypeGpuDrawPanes({
    paneBuffers: input.backend.paneBuffers,
    panes: input.panes,
    tileCache: input.backend.tileCache,
  })
  const resourceRetainPanes = mergePaneLists(retainPanes, staleRetainPanes)
  syncTypeGpuPaneResources({
    artifacts: input.backend.artifacts,
    atlas: input.backend.atlas,
    paneBuffers: input.backend.paneBuffers,
    panes: resourcePanes,
    retainPanes: resourceRetainPanes,
  })
  syncTypeGpuAtlasResources(input.backend.artifacts, input.backend.atlas)
  const drawPanes = resolveTypeGpuDrawPanes({
    onTileMiss: noteTypeGpuTileMiss,
    paneBuffers: input.backend.paneBuffers,
    panes: input.panes,
    tileCache: input.backend.tileCache,
  })
  markVisibleTilePanes(input.backend.tileCache, drawPanes)
  drawTypeGpuPanes({
    artifacts: input.backend.artifacts,
    paneBuffers: input.backend.paneBuffers,
    panes: drawPanes,
    scrollSnapshot: input.scrollSnapshot,
    surface: input.surface,
  })
}

function markVisibleTilePanes(cache: TileGpuCache, panes: readonly WorkbookRenderPaneState[]): void {
  cache.markVisible(new Set(panes.map((pane) => buildTileGpuCacheKey(pane.packedScene))))
}

export function resolveTypeGpuDrawPanes(input: {
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly tileCache: TileGpuCache
  readonly onTileMiss?: ((tileKey: string) => void) | undefined
}): readonly WorkbookRenderPaneState[] {
  return input.panes.map((pane) => {
    const packedScene = pane.packedScene
    const exactKey = buildTileGpuCacheKey(packedScene)
    const exact = input.paneBuffers.peek(exactKey)
    if (exact && isPaneDrawReady(exact, pane)) {
      return pane
    }
    const stale = input.tileCache.findStaleValid(packedScene.key, { excludeKey: exactKey })
    if (stale) {
      const staleEntry = input.paneBuffers.peek(stale.key)
      if (staleEntry && isPaneDrawReady(staleEntry, { ...pane, packedScene: stale.packet })) {
        return { ...pane, packedScene: stale.packet }
      }
    }
    input.onTileMiss?.(exactKey)
    return pane
  })
}

function isPaneDrawReady(entry: WorkbookPaneBufferEntry, pane: WorkbookRenderPaneState): boolean {
  const packedScene = pane.packedScene
  const rectReady =
    packedScene.rectCount === 0 ? entry.rectSignature !== null : entry.rectBuffer !== null && entry.rectCount >= packedScene.rectCount
  const textReady =
    packedScene.textCount === 0
      ? entry.textSignature !== null
      : entry.textBuffer !== null && entry.textSignature !== null && entry.textCount > 0
  return rectReady && textReady
}

function mergePaneLists(
  primary: readonly WorkbookRenderPaneState[],
  secondary: readonly WorkbookRenderPaneState[],
): readonly WorkbookRenderPaneState[] {
  const result: WorkbookRenderPaneState[] = []
  const seen = new Set<string>()
  for (const pane of [...primary, ...secondary]) {
    const key = buildTileGpuCacheKey(pane.packedScene)
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    result.push(pane)
  }
  return result
}
