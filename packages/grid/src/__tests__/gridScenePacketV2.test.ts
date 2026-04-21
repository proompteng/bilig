import { describe, expect, test } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  type GridScenePacketV2,
} from '../renderer-v2/scene-packet-v2.js'
import { validateGridScenePacketV2 } from '../renderer-v2/scene-packet-validator.js'

function createPacket(overrides: Partial<GridScenePacketV2> = {}): GridScenePacketV2 {
  return {
    generation: 1,
    magic: GRID_SCENE_PACKET_V2_MAGIC,
    paneId: 'body',
    rectCount: 1,
    rects: new Float32Array(GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT),
    sheetName: 'Sheet1',
    surfaceSize: { height: 220, width: 480 },
    textCount: 1,
    textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
    version: GRID_SCENE_PACKET_V2_VERSION,
    viewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
    ...overrides,
  }
}

describe('grid scene packet v2', () => {
  test('accepts well-formed typed scene packets', () => {
    const packet = createPacket()
    packet.rects.set([0, 0, 104, 22, 1, 1, 1, 1])
    packet.textMetrics.set([4, 3, 80, 16, 0, 0, 0, 0])

    expect(validateGridScenePacketV2(packet)).toEqual({ ok: true })
  })

  test('rejects malformed packet metadata and buffers', () => {
    expect(validateGridScenePacketV2(createPacket({ rects: new Float32Array(1) }))).toEqual({
      ok: false,
      reason: 'rect buffer too small',
    })
    expect(validateGridScenePacketV2(createPacket({ viewport: { colEnd: 1, colStart: 2, rowEnd: 1, rowStart: 0 } }))).toEqual({
      ok: false,
      reason: 'empty viewport',
    })
    expect(validateGridScenePacketV2(createPacket({ rects: new Float32Array([0, 0, Number.NaN, 22, 1, 1, 1, 1]) }))).toEqual({
      ok: false,
      reason: 'bad rect size',
    })
  })
})
