import { describe, expect, test, vi } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
} from '../renderer-v2/scene-packet-v2.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import {
  TypeGpuTileResourceCacheV3,
  resolveWorkbookTileContentBufferKeyV3,
  resolveWorkbookTilePlacementBufferKeyV3,
} from '../renderer-v3/typegpu-tile-buffer-pool.js'
import { resolveTypeGpuDrawTilePanesV3 } from '../renderer-v3/typegpu-workbook-backend-v3.js'
import { TileResidencyV3 } from '../renderer-v3/tile-residency.js'

function createRenderTile(valueVersion: number, tileId = 101): GridRenderTile {
  return {
    bounds: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 7,
    },
    lastBatchId: valueVersion,
    lastCameraSeq: valueVersion,
    rectCount: 0,
    rectInstances: new Float32Array(GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT),
    textCount: 0,
    textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
    textRuns: [],
    tileId,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 0,
      styles: valueVersion,
      text: valueVersion,
      values: valueVersion,
    },
  }
}

function createTilePane(tile: GridRenderTile): WorkbookRenderTilePaneState {
  return {
    contentOffset: { x: 0, y: 0 },
    frame: { height: 220, width: 480, x: 0, y: 0 },
    generation: tile.lastBatchId,
    paneId: 'body',
    scrollAxes: { x: true, y: true },
    surfaceSize: { height: 220, width: 480 },
    tile,
    viewport: tile.bounds,
  }
}

function upsertRenderTile(residency: TileResidencyV3<GridRenderTile, null>, tile: GridRenderTile): void {
  residency.upsert({
    axisSeqX: tile.version.axisX,
    axisSeqY: tile.version.axisY,
    byteSizeCpu: 1,
    byteSizeGpu: 1,
    colTile: tile.coord.colTile,
    dprBucket: tile.coord.dprBucket,
    freezeSeq: tile.version.freeze,
    key: tile.tileId,
    packet: tile,
    rectSeq: tile.version.values,
    resources: null,
    rowTile: tile.coord.rowTile,
    sheetOrdinal: tile.coord.sheetId,
    state: 'ready',
    styleSeq: tile.version.styles,
    textSeq: tile.version.text,
    valueSeq: tile.version.values,
  })
}

describe('workbook typegpu backend v3 tile path', () => {
  test('draws V3 tile panes from the numeric tile residency path', () => {
    const tile = createRenderTile(2)
    const pane = createTilePane(tile)
    const tileResources = new TypeGpuTileResourceCacheV3()
    const residency = new TileResidencyV3<GridRenderTile, null>()
    upsertRenderTile(residency, tile)
    const entry = tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(pane))
    entry.rectSignature = 'rect:2'
    entry.textSignature = 'text:2'

    expect(resolveTypeGpuDrawTilePanesV3({ panes: [pane], residency, tileResources })[0]?.tile).toBe(tile)
  })

  test('shares V3 content resources but keeps placement resources distinct for frozen placements', () => {
    const tile = createRenderTile(2)
    const bodyPane = createTilePane(tile)
    const frozenPane: WorkbookRenderTilePaneState = {
      ...bodyPane,
      frame: { height: 48, width: 480, x: 46, y: 24 },
      paneId: 'top:0:0',
      scrollAxes: { x: true, y: false },
    }

    expect(resolveWorkbookTileContentBufferKeyV3(bodyPane)).toBe(resolveWorkbookTileContentBufferKeyV3(frozenPane))
    expect(resolveWorkbookTilePlacementBufferKeyV3(bodyPane)).not.toBe(resolveWorkbookTilePlacementBufferKeyV3(frozenPane))
  })

  test('prunes V3 tile content and placement resources independently', () => {
    const cache = new TypeGpuTileResourceCacheV3()
    const tile = createRenderTile(2)
    const bodyPane = createTilePane(tile)
    const frozenPane: WorkbookRenderTilePaneState = {
      ...bodyPane,
      paneId: 'top:0:0',
      scrollAxes: { x: true, y: false },
    }
    const contentKey = resolveWorkbookTileContentBufferKeyV3(bodyPane)
    const bodyPlacementKey = resolveWorkbookTilePlacementBufferKeyV3(bodyPane)
    const frozenPlacementKey = resolveWorkbookTilePlacementBufferKeyV3(frozenPane)

    const content = cache.getContent(contentKey)
    const bodyPlacement = cache.getPlacement(bodyPlacementKey)
    cache.getPlacement(frozenPlacementKey)

    cache.pruneExcept({
      contentKeys: new Set([contentKey]),
      placementKeys: new Set([bodyPlacementKey]),
    })
    expect(cache.peekContent(contentKey)).toBe(content)
    expect(cache.peekPlacement(bodyPlacementKey)).toBe(bodyPlacement)
    expect(cache.peekPlacement(frozenPlacementKey)).toBeNull()

    cache.pruneExcept({ contentKeys: new Set(), placementKeys: new Set() })
    expect(cache.peekContent(contentKey)).toBeNull()
    expect(cache.peekPlacement(bodyPlacementKey)).toBeNull()
  })

  test('reports a V3 tile miss when tile resources are not draw-ready', () => {
    const tile = createRenderTile(2)
    const pane = createTilePane(tile)
    const onTileMiss = vi.fn()

    expect(
      resolveTypeGpuDrawTilePanesV3({
        onTileMiss,
        panes: [pane],
        residency: new TileResidencyV3<GridRenderTile, null>(),
        tileResources: new TypeGpuTileResourceCacheV3(),
      })[0]?.tile,
    ).toBe(tile)
    expect(onTileMiss).toHaveBeenCalledWith(tile.tileId)
  })
})
