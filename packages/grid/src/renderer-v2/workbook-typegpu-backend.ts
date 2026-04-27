import type { WorkbookGridScrollSnapshot } from '../workbookGridScrollStore.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { WorkbookRenderPaneState } from './pane-scene-types.js'
import { noteTypeGpuTileMiss } from './grid-render-counters.js'
import { createGlyphAtlas } from './typegpu-atlas-manager.js'
import { WorkbookPaneBufferCache, type WorkbookPaneBufferEntry } from './pane-buffer-cache.js'
import {
  createTypeGpuRenderer,
  destroyTypeGpuRenderer,
  syncTypeGpuAtlasResources,
  type TypeGpuRendererArtifacts,
} from './typegpu-backend.js'
import { drawTypeGpuPanes, drawTypeGpuTilePanes, type TypeGpuDrawSurface } from './typegpu-render-pass.js'
import {
  WORKBOOK_DYNAMIC_OVERLAY_BUFFER_KEY,
  resolveWorkbookHeaderBufferKey,
  resolveWorkbookPaneBufferKey,
  resolveWorkbookTilePaneBufferKey,
  syncTypeGpuHeaderResources,
  syncTypeGpuOverlayResources,
  syncTypeGpuPaneResources,
  syncTypeGpuTilePaneResources,
} from './typegpu-buffer-pool.js'
import { createTypeGpuSurfaceState, syncTypeGpuCanvasSurface, type TypeGpuSurfaceState } from './typegpu-surface.js'
import { buildTileGpuCacheKey, TileGpuCache, syncTileGpuCacheFromPanes } from './tile-gpu-cache.js'
import type { GridScenePacketV2 } from './scene-packet-v2.js'
import type { DynamicGridOverlayBatchV3 } from '../renderer-v3/dynamic-overlay-batch.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import { TileResidencyV3 } from '../renderer-v3/tile-residency.js'

export interface WorkbookTypeGpuBackend {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly surfaceState: TypeGpuSurfaceState
  readonly tileCache: TileGpuCache
  readonly tileResidency: TileResidencyV3<GridRenderTile, null>
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
    tileResidency: new TileResidencyV3<GridRenderTile, null>(),
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
  readonly headerPanes?: readonly GridHeaderPaneState[] | undefined
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly preloadPanes?: readonly WorkbookRenderPaneState[] | undefined
  readonly tilePanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly preloadTilePanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly overlay?: DynamicGridOverlayBatchV3 | null | undefined
  readonly syncPreloadPanes?: boolean | undefined
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuDrawSurface
}): void {
  if (input.tilePanes) {
    drawWorkbookTypeGpuTileFrame({
      backend: input.backend,
      headerPanes: input.headerPanes,
      overlay: input.overlay,
      preloadTilePanes: input.preloadTilePanes,
      scrollSnapshot: input.scrollSnapshot,
      surface: input.surface,
      syncPreloadPanes: input.syncPreloadPanes,
      tilePanes: input.tilePanes,
    })
    return
  }

  const retainPanes = input.preloadPanes?.length ? [...input.preloadPanes, ...input.panes] : input.panes
  const resourcePanes = input.syncPreloadPanes === false ? input.panes : retainPanes
  const headerPanes = input.headerPanes ?? []
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
    retainBufferKeys: [...headerPanes.map(resolveWorkbookHeaderBufferKey), ...(input.overlay ? [WORKBOOK_DYNAMIC_OVERLAY_BUFFER_KEY] : [])],
  })
  syncTypeGpuHeaderResources({
    artifacts: input.backend.artifacts,
    atlas: input.backend.atlas,
    headerPanes,
    paneBuffers: input.backend.paneBuffers,
  })
  syncTypeGpuOverlayResources({
    artifacts: input.backend.artifacts,
    overlay: input.overlay ?? null,
    paneBuffers: input.backend.paneBuffers,
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
    headerPanes,
    overlay: input.overlay ?? null,
    paneBuffers: input.backend.paneBuffers,
    panes: drawPanes,
    scrollSnapshot: input.scrollSnapshot,
    surface: input.surface,
  })
}

function drawWorkbookTypeGpuTileFrame(input: {
  readonly backend: WorkbookTypeGpuBackend
  readonly headerPanes?: readonly GridHeaderPaneState[] | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly preloadTilePanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly overlay?: DynamicGridOverlayBatchV3 | null | undefined
  readonly syncPreloadPanes?: boolean | undefined
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuDrawSurface
}): void {
  const retainPanes = input.preloadTilePanes?.length ? [...input.preloadTilePanes, ...input.tilePanes] : input.tilePanes
  const resourcePanes = input.syncPreloadPanes === false ? input.tilePanes : retainPanes
  const headerPanes = input.headerPanes ?? []
  syncRenderTileResidencyFromPanes({
    panes: resourcePanes,
    residency: input.backend.tileResidency,
  })
  syncTypeGpuTilePaneResources({
    artifacts: input.backend.artifacts,
    atlas: input.backend.atlas,
    paneBuffers: input.backend.paneBuffers,
    panes: resourcePanes,
    retainPanes,
    retainBufferKeys: [...headerPanes.map(resolveWorkbookHeaderBufferKey), ...(input.overlay ? [WORKBOOK_DYNAMIC_OVERLAY_BUFFER_KEY] : [])],
  })
  syncTypeGpuHeaderResources({
    artifacts: input.backend.artifacts,
    atlas: input.backend.atlas,
    headerPanes,
    paneBuffers: input.backend.paneBuffers,
  })
  syncTypeGpuOverlayResources({
    artifacts: input.backend.artifacts,
    overlay: input.overlay ?? null,
    paneBuffers: input.backend.paneBuffers,
  })
  syncTypeGpuAtlasResources(input.backend.artifacts, input.backend.atlas)
  const drawPanes = resolveTypeGpuDrawTilePanes({
    onTileMiss: (tileKey) => noteTypeGpuTileMiss(String(tileKey)),
    paneBuffers: input.backend.paneBuffers,
    panes: input.tilePanes,
    residency: input.backend.tileResidency,
  })
  input.backend.tileResidency.markVisible(drawPanes.map((pane) => pane.tile.tileId))
  drawTypeGpuTilePanes({
    artifacts: input.backend.artifacts,
    headerPanes,
    overlay: input.overlay ?? null,
    paneBuffers: input.backend.paneBuffers,
    scrollSnapshot: input.scrollSnapshot,
    surface: input.surface,
    tilePanes: drawPanes,
  })
}

function syncRenderTileResidencyFromPanes(input: {
  readonly residency: TileResidencyV3<GridRenderTile, null>
  readonly panes: readonly WorkbookRenderTilePaneState[]
}): void {
  for (const pane of input.panes) {
    const tile = pane.tile
    input.residency.upsert({
      axisSeqX: tile.version.axisX,
      axisSeqY: tile.version.axisY,
      byteSizeCpu: estimateRenderTileCpuBytes(tile),
      byteSizeGpu: estimateRenderTileGpuBytes(tile),
      colTile: tile.coord.colTile,
      dprBucket: tile.coord.dprBucket,
      freezeSeq: tile.version.freeze,
      key: tile.tileId,
      packet: tile,
      rectSeq: Math.max(tile.version.values, tile.version.styles, tile.version.axisX, tile.version.axisY),
      resources: null,
      rowTile: tile.coord.rowTile,
      sheetOrdinal: tile.coord.sheetId,
      state: 'ready',
      styleSeq: tile.version.styles,
      textSeq: tile.version.text,
      valueSeq: tile.version.values,
    })
  }
  input.residency.markVisible(input.panes.map((pane) => pane.tile.tileId))
  input.residency.evictToSize(256)
}

export function resolveTypeGpuDrawTilePanes(input: {
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly residency: TileResidencyV3<GridRenderTile, null>
  readonly onTileMiss?: ((tileKey: number) => void) | undefined
}): readonly WorkbookRenderTilePaneState[] {
  return input.panes.map((pane) => {
    const entry = input.residency.getExact(pane.tile.tileId)
    const exact = input.paneBuffers.peek(resolveWorkbookTilePaneBufferKey(pane))
    if (entry?.packet && exact && isTilePaneDrawReady(exact, pane)) {
      return { ...pane, tile: entry.packet }
    }
    input.onTileMiss?.(pane.tile.tileId)
    return pane
  })
}

function isTilePaneDrawReady(entry: WorkbookPaneBufferEntry, pane: WorkbookRenderTilePaneState): boolean {
  const tile = pane.tile
  const rectReady = tile.rectCount === 0 ? entry.rectSignature !== null : entry.rectBuffer !== null && entry.rectCount >= tile.rectCount
  const textReady =
    tile.textCount === 0 ? entry.textSignature !== null : entry.textBuffer !== null && entry.textSignature !== null && entry.textCount > 0
  return rectReady && textReady
}

function estimateRenderTileCpuBytes(tile: GridRenderTile): number {
  let textBytes = 0
  for (const run of tile.textRuns) {
    textBytes += run.text.length * 2 + run.font.length * 2 + run.color.length * 2 + 80
  }
  return tile.rectInstances.byteLength + tile.textMetrics.byteLength + textBytes
}

function estimateRenderTileGpuBytes(tile: GridRenderTile): number {
  return tile.rectInstances.byteLength + tile.textMetrics.byteLength
}

function markVisibleTilePanes(cache: TileGpuCache, panes: readonly WorkbookRenderPaneState[]): void {
  cache.beginVisibilityPass()
  let marked = 0
  for (const pane of panes) {
    if (cache.markVisibleKey(buildTileGpuCacheKey(pane.packedScene))) {
      marked += 1
    }
  }
  cache.finishVisibilityPass(marked)
}

export function resolveTypeGpuDrawPanes(input: {
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly tileCache: TileGpuCache
  readonly onTileMiss?: ((tileKey: string) => void) | undefined
}): readonly WorkbookRenderPaneState[] {
  return input.panes.map((pane) => {
    const packedScene = pane.packedScene
    const exactTileKey = buildTileGpuCacheKey(packedScene)
    const exact = input.paneBuffers.peek(resolveWorkbookPaneBufferKey(pane))
    if (exact && isPaneDrawReady(exact, pane)) {
      return pane
    }
    const stale = input.tileCache.findStaleValid(packedScene.key, { excludeKey: exactTileKey })
    if (stale && hasCompatiblePaneOrigin(stale.packet, packedScene)) {
      const stalePane = { ...pane, packedScene: stale.packet }
      const staleEntry = input.paneBuffers.peek(resolveWorkbookPaneBufferKey(stalePane))
      if (staleEntry && isPaneDrawReady(staleEntry, stalePane)) {
        return stalePane
      }
    }
    input.onTileMiss?.(exactTileKey)
    return pane
  })
}

function hasCompatiblePaneOrigin(candidate: GridScenePacketV2, desired: GridScenePacketV2): boolean {
  return candidate.viewport.rowStart === desired.viewport.rowStart && candidate.viewport.colStart === desired.viewport.colStart
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
    const key = resolveWorkbookPaneBufferKey(pane)
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    result.push(pane)
  }
  return result
}
