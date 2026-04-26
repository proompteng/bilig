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
import { resolveTypeGpuDrawPanes } from '../renderer-v2/workbook-typegpu-backend.js'
import type { WorkbookRenderPaneState } from '../renderer-v2/pane-scene-types.js'

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

function markReady(cache: WorkbookPaneBufferCache, packet: GridScenePacketV2): WorkbookPaneBufferEntry {
  const entry = cache.get(buildTileGpuCacheKey(packet))
  entry.rectSignature = `rect:${packet.generation}`
  entry.textSignature = `text:${packet.generation}`
  return entry
}

describe('workbook typegpu backend tile fallback', () => {
  test('draws the exact pane when its buffer resources are ready', () => {
    const packet = createPacket(2)
    const pane = createPane(packet)
    const paneBuffers = new WorkbookPaneBufferCache()
    const tileCache = new TileGpuCache()
    tileCache.upsert(packet)
    markReady(paneBuffers, packet)

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
    markReady(paneBuffers, stalePacket)

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
    markReady(paneBuffers, stalePacket)

    expect(resolveTypeGpuDrawPanes({ onTileMiss, paneBuffers, panes: [pane], tileCache })[0]?.packedScene).toBe(exactPacket)
    expect(onTileMiss).toHaveBeenCalledWith(buildTileGpuCacheKey(exactPacket))
  })
})
