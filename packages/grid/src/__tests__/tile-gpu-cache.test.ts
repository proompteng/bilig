import { describe, expect, test } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  type GridScenePacketV2,
} from '../renderer-v2/scene-packet-v2.js'
import { TileGpuCache, buildTileGpuCacheKey, syncTileGpuCacheFromPanes } from '../renderer-v2/tile-gpu-cache.js'

function createPacket(generation: number, colStart = 0): GridScenePacketV2 {
  return {
    generation,
    borderRectCount: 0,
    fillRectCount: 0,
    magic: GRID_SCENE_PACKET_V2_MAGIC,
    paneId: 'body',
    rectCount: 0,
    rects: new Float32Array(GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT),
    sheetName: 'Sheet1',
    surfaceSize: { height: 200, width: 400 },
    textCount: 0,
    textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
    version: GRID_SCENE_PACKET_V2_VERSION,
    viewport: { colEnd: colStart + 127, colStart, rowEnd: 31, rowStart: 0 },
  }
}

describe('TileGpuCache', () => {
  test('keeps newest generation for a tile key', () => {
    const cache = new TileGpuCache()
    const newer = createPacket(2)
    cache.upsert(newer)
    cache.upsert(createPacket(1))

    expect(cache.get(buildTileGpuCacheKey(newer))?.packet.generation).toBe(2)
  })

  test('evicts least-recent non-visible tiles first', () => {
    const cache = new TileGpuCache()
    const first = cache.upsert(createPacket(1, 0))
    const second = cache.upsert(createPacket(1, 128))
    const third = cache.upsert(createPacket(1, 256))
    cache.markVisible(new Set([second.key]))
    cache.get(third.key)
    cache.evictTo(2)

    expect(cache.get(first.key)).toBeNull()
    expect(cache.get(second.key)).not.toBeNull()
    expect(cache.get(third.key)).not.toBeNull()
  })

  test('rejects invalid packets', () => {
    const cache = new TileGpuCache()
    expect(() => cache.upsert({ ...createPacket(1), rects: new Float32Array(1) })).toThrow(/rect buffer too small/u)
  })

  test('syncs renderer panes into visible cache entries', () => {
    const cache = new TileGpuCache()
    const packet = createPacket(1)
    syncTileGpuCacheFromPanes({ cache, maxEntries: 8, panes: [{ packedScene: packet }] })

    expect(cache.get(buildTileGpuCacheKey(packet))?.visible).toBe(true)
  })
})
