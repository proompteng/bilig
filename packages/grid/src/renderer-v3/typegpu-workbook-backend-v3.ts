import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { WorkbookGridScrollSnapshot } from '../workbookGridScrollStore.js'
import { noteTypeGpuTileMiss } from '../grid-render-counters.js'
import { createGlyphAtlas } from './typegpu-atlas-manager.js'
import {
  createTypeGpuRenderer,
  destroyTypeGpuRenderer,
  syncTypeGpuAtlasResources,
  type TypeGpuRendererArtifacts,
} from './typegpu-primitives.js'
import { createTypeGpuSurfaceState, syncTypeGpuCanvasSurface, type TypeGpuSurfaceState } from './typegpu-surface.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { GridRenderTile } from './render-tile-source.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { TileResidencyV3 } from './tile-residency.js'
import {
  TypeGpuLayerResourceCacheV3,
  pruneTypeGpuLayerResourcesV3,
  syncTypeGpuHeaderResourcesV3,
  syncTypeGpuOverlayResourcesV3,
} from './typegpu-layer-buffer-pool.js'
import {
  TypeGpuTileResourceCacheV3,
  resolveWorkbookTileContentBufferKeyV3,
  syncTypeGpuTilePaneResourcesV3,
  type TypeGpuTileContentResourceEntryV3,
} from './typegpu-tile-buffer-pool.js'
import { drawTypeGpuTilePanesV3, type TypeGpuTileDrawSurface } from './typegpu-tile-render-pass.js'

export interface WorkbookTypeGpuBackendV3 {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly surfaceState: TypeGpuSurfaceState
  readonly tileResources: TypeGpuTileResourceCacheV3
  readonly tileResidency: TileResidencyV3<GridRenderTile, null>
}

export async function createWorkbookTypeGpuBackendV3(canvas: HTMLCanvasElement): Promise<WorkbookTypeGpuBackendV3 | null> {
  const artifacts = await createTypeGpuRenderer(canvas)
  if (!artifacts) {
    return null
  }
  return {
    artifacts,
    atlas: createGlyphAtlas(),
    layerResources: new TypeGpuLayerResourceCacheV3(artifacts),
    surfaceState: createTypeGpuSurfaceState(),
    tileResources: new TypeGpuTileResourceCacheV3(artifacts),
    tileResidency: new TileResidencyV3<GridRenderTile, null>(),
  }
}

export function destroyWorkbookTypeGpuBackendV3(backend: WorkbookTypeGpuBackendV3): void {
  backend.tileResources.dispose()
  backend.layerResources.dispose()
  destroyTypeGpuRenderer(backend.artifacts)
}

export function syncWorkbookTypeGpuSurfaceV3(input: {
  readonly backend: WorkbookTypeGpuBackendV3
  readonly canvas: HTMLCanvasElement
  readonly size: TypeGpuTileDrawSurface
}): void {
  syncTypeGpuCanvasSurface({
    artifacts: input.backend.artifacts,
    canvas: input.canvas,
    size: input.size,
    state: input.backend.surfaceState,
  })
}

export function drawWorkbookTypeGpuTileFrameV3(input: {
  readonly backend: WorkbookTypeGpuBackendV3
  readonly headerPanes?: readonly GridHeaderPaneState[] | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly preloadTilePanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly overlay?: DynamicGridOverlayBatchV3 | null | undefined
  readonly syncPreloadPanes?: boolean | undefined
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuTileDrawSurface
}): void {
  const retainPanes = input.preloadTilePanes?.length ? [...input.preloadTilePanes, ...input.tilePanes] : input.tilePanes
  const resourcePanes = input.syncPreloadPanes === false ? input.tilePanes : retainPanes
  const headerPanes = input.headerPanes ?? []
  syncRenderTileResidencyFromPanesV3({
    panes: resourcePanes,
    residency: input.backend.tileResidency,
  })
  syncTypeGpuTilePaneResourcesV3({
    artifacts: input.backend.artifacts,
    atlas: input.backend.atlas,
    panes: resourcePanes,
    retainPanes,
    tileResources: input.backend.tileResources,
  })
  pruneTypeGpuLayerResourcesV3({
    headerPanes,
    layerResources: input.backend.layerResources,
    overlay: input.overlay ?? null,
  })
  syncTypeGpuHeaderResourcesV3({
    artifacts: input.backend.artifacts,
    atlas: input.backend.atlas,
    headerPanes,
    layerResources: input.backend.layerResources,
  })
  syncTypeGpuOverlayResourcesV3({
    layerResources: input.backend.layerResources,
    overlay: input.overlay ?? null,
  })
  syncTypeGpuAtlasResources(input.backend.artifacts, input.backend.atlas)
  const drawPanes = resolveTypeGpuDrawTilePanesV3({
    onTileMiss: (tileKey) => noteTypeGpuTileMiss(String(tileKey)),
    panes: input.tilePanes,
    residency: input.backend.tileResidency,
    tileResources: input.backend.tileResources,
  })
  input.backend.tileResidency.markVisible(drawPanes.map((pane) => pane.tile.tileId))
  drawTypeGpuTilePanesV3({
    artifacts: input.backend.artifacts,
    headerPanes,
    layerResources: input.backend.layerResources,
    overlay: input.overlay ?? null,
    scrollSnapshot: input.scrollSnapshot,
    surface: input.surface,
    tileResources: input.backend.tileResources,
    tilePanes: drawPanes,
  })
}

function syncRenderTileResidencyFromPanesV3(input: {
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

export function resolveTypeGpuDrawTilePanesV3(input: {
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly residency: TileResidencyV3<GridRenderTile, null>
  readonly tileResources: Pick<TypeGpuTileResourceCacheV3, 'peekContent'>
  readonly onTileMiss?: ((tileKey: number) => void) | undefined
}): readonly WorkbookRenderTilePaneState[] {
  return input.panes.map((pane) => {
    const entry = input.residency.getExact(pane.tile.tileId)
    const exact = input.tileResources.peekContent(resolveWorkbookTileContentBufferKeyV3(pane))
    if (entry?.packet && exact && isTileContentDrawReady(exact, pane)) {
      return { ...pane, tile: entry.packet }
    }
    input.onTileMiss?.(pane.tile.tileId)
    return pane
  })
}

function isTileContentDrawReady(entry: TypeGpuTileContentResourceEntryV3, pane: WorkbookRenderTilePaneState): boolean {
  const tile = pane.tile
  const rectReady = tile.rectCount === 0 ? entry.rectSignature !== null : entry.rectHandle !== null && entry.rectCount >= tile.rectCount
  const textReady =
    tile.textCount === 0 ? entry.textSignature !== null : entry.textHandle !== null && entry.textSignature !== null && entry.textCount > 0
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
