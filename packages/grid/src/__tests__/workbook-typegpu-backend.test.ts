import { describe, expect, test, vi } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  createGridTileKeyV2,
  type GridScenePacketV2,
} from '../renderer-v2/scene-packet-v2.js'
import { WorkbookPaneBufferCache, type WorkbookPaneBufferEntry } from '../renderer-v2/pane-buffer-cache.js'
import { TileGpuCache, buildTileGpuCacheKey } from '../renderer-v2/tile-gpu-cache.js'
import { resolveWorkbookPaneBufferKey, resolveWorkbookTilePaneBufferKey } from '../renderer-v2/typegpu-buffer-pool.js'
import { resolveTypeGpuDrawPanes, resolveTypeGpuDrawTilePanes } from '../renderer-v2/workbook-typegpu-backend.js'
import type { WorkbookRenderPaneState } from '../renderer-v2/pane-scene-types.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import { TileResidencyV3 } from '../renderer-v3/tile-residency.js'

function createPacket(valueVersion: number, rowStart = 0, colStart = 0): GridScenePacketV2 {
  const viewport = { colEnd: colStart + 127, colStart, rowEnd: rowStart + 31, rowStart }
  return {
    borderRectCount: 0,
    cameraSeq: valueVersion,
    fillRectCount: 0,
    generatedAt: valueVersion,
    generation: valueVersion,
    key: createGridTileKeyV2({
      paneId: 'body',
      sheetName: 'Sheet1',
      valueVersion,
      styleVersion: valueVersion,
      viewport,
    }),
    magic: GRID_SCENE_PACKET_V2_MAGIC,
    paneId: 'body',
    rectCount: 0,
    rectInstances: new Float32Array(GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT),
    rects: new Float32Array(GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT),
    rectSignature: 'rect-empty',
    requestSeq: valueVersion,
    sheetName: 'Sheet1',
    surfaceSize: { height: 220, width: 480 },
    textCount: 0,
    textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
    textRuns: [],
    textSignature: 'text-empty',
    version: GRID_SCENE_PACKET_V2_VERSION,
    viewport,
  }
}

function createPane(packet: GridScenePacketV2): WorkbookRenderPaneState {
  return {
    contentOffset: { x: 0, y: 0 },
    frame: { height: 220, width: 480, x: 0, y: 0 },
    generation: packet.generation,
    packedScene: packet,
    paneId: 'body',
    scrollAxes: { x: true, y: true },
    surfaceSize: packet.surfaceSize,
    viewport: packet.viewport,
  }
}

function markReady(cache: WorkbookPaneBufferCache, pane: WorkbookRenderPaneState): WorkbookPaneBufferEntry {
  const entry = cache.get(resolveWorkbookPaneBufferKey(pane))
  const packet = pane.packedScene
  entry.rectSignature = `rect:${packet.generation}`
  entry.textSignature = `text:${packet.generation}`
  return entry
}

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

describe('workbook typegpu backend tile fallback', () => {
  test('draws the exact pane when its buffer resources are ready', () => {
    const packet = createPacket(2)
    const pane = createPane(packet)
    const paneBuffers = new WorkbookPaneBufferCache()
    const tileCache = new TileGpuCache()
    tileCache.upsert(packet)
    markReady(paneBuffers, pane)

    expect(resolveTypeGpuDrawPanes({ paneBuffers, panes: [pane], tileCache })[0]?.packedScene).toBe(packet)
  })

  test('does not substitute a retained pane from an older data revision', () => {
    const stalePacket = createPacket(1)
    const exactPacket = createPacket(2)
    const pane = createPane(exactPacket)
    const paneBuffers = new WorkbookPaneBufferCache()
    const tileCache = new TileGpuCache()
    const onTileMiss = vi.fn()
    tileCache.upsert(stalePacket)
    tileCache.upsert(exactPacket)
    markReady(paneBuffers, createPane(stalePacket))

    expect(resolveTypeGpuDrawPanes({ onTileMiss, paneBuffers, panes: [pane], tileCache })[0]?.packedScene).toBe(exactPacket)
    expect(onTileMiss).toHaveBeenCalledWith(buildTileGpuCacheKey(exactPacket))
  })

  test('reports a tile miss when neither exact nor stale-valid resources are draw-ready', () => {
    const exactPacket = createPacket(2)
    const pane = createPane(exactPacket)
    const onTileMiss = vi.fn()

    expect(
      resolveTypeGpuDrawPanes({
        onTileMiss,
        paneBuffers: new WorkbookPaneBufferCache(),
        panes: [pane],
        tileCache: new TileGpuCache(),
      })[0]?.packedScene,
    ).toBe(exactPacket)
    expect(onTileMiss).toHaveBeenCalledWith(buildTileGpuCacheKey(exactPacket))
  })

  test('does not substitute an overlapping stale pane with a different local origin', () => {
    const stalePacket = createPacket(2, 0, 0)
    const exactPacket = createPacket(2, 4, 0)
    const pane = createPane(exactPacket)
    const paneBuffers = new WorkbookPaneBufferCache()
    const tileCache = new TileGpuCache()
    const onTileMiss = vi.fn()
    tileCache.upsert(stalePacket)
    markReady(paneBuffers, createPane(stalePacket))

    expect(resolveTypeGpuDrawPanes({ onTileMiss, paneBuffers, panes: [pane], tileCache })[0]?.packedScene).toBe(exactPacket)
    expect(onTileMiss).toHaveBeenCalledWith(buildTileGpuCacheKey(exactPacket))
  })

  test('keeps resource keys distinct for separate placements of the same content tile', () => {
    const packet = createPacket(2)
    const bodyPane = createPane(packet)
    const frozenPane: WorkbookRenderPaneState = {
      ...bodyPane,
      frame: { height: 48, width: 480, x: 46, y: 24 },
      paneId: 'top:0:0',
      scrollAxes: { x: true, y: false },
    }

    expect(buildTileGpuCacheKey(bodyPane.packedScene)).toBe(buildTileGpuCacheKey(frozenPane.packedScene))
    expect(resolveWorkbookPaneBufferKey(bodyPane)).not.toBe(resolveWorkbookPaneBufferKey(frozenPane))
  })

  test('draws V3 tile panes from the numeric tile residency path', () => {
    const tile = createRenderTile(2)
    const pane = createTilePane(tile)
    const paneBuffers = new WorkbookPaneBufferCache()
    const residency = new TileResidencyV3<GridRenderTile, null>()
    upsertRenderTile(residency, tile)
    const entry = paneBuffers.get(resolveWorkbookTilePaneBufferKey(pane))
    entry.rectSignature = 'rect:2'
    entry.textSignature = 'text:2'

    expect(resolveTypeGpuDrawTilePanes({ paneBuffers, panes: [pane], residency })[0]?.tile).toBe(tile)
  })

  test('shares V3 resource keys for frozen placements of the same content tile', () => {
    const tile = createRenderTile(2)
    const bodyPane = createTilePane(tile)
    const frozenPane: WorkbookRenderTilePaneState = {
      ...bodyPane,
      frame: { height: 48, width: 480, x: 46, y: 24 },
      paneId: 'top:0:0',
      scrollAxes: { x: true, y: false },
    }

    expect(resolveWorkbookTilePaneBufferKey(bodyPane)).toBe(resolveWorkbookTilePaneBufferKey(frozenPane))
  })

  test('reports a V3 tile miss when tile resources are not draw-ready', () => {
    const tile = createRenderTile(2)
    const pane = createTilePane(tile)
    const onTileMiss = vi.fn()

    expect(
      resolveTypeGpuDrawTilePanes({
        onTileMiss,
        paneBuffers: new WorkbookPaneBufferCache(),
        panes: [pane],
        residency: new TileResidencyV3<GridRenderTile, null>(),
      })[0]?.tile,
    ).toBe(tile)
    expect(onTileMiss).toHaveBeenCalledWith(tile.tileId)
  })
})
