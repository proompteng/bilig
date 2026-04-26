import { describe, expect, test } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  createGridTileKeyV2,
  isStaleValidGridTileKeyV2,
  serializeGridTileKeyV2,
  type GridScenePacketV2,
} from '../renderer-v2/scene-packet-v2.js'
import { validateGridScenePacketV2 } from '../renderer-v2/scene-packet-validator.js'

function createPacket(overrides: Partial<GridScenePacketV2> = {}): GridScenePacketV2 {
  const viewport = { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 }
  return {
    generation: 1,
    cameraSeq: 11,
    borderRectCount: 0,
    fillRectCount: 1,
    generatedAt: 1234,
    key: createGridTileKeyV2({ paneId: 'body', sheetName: 'Sheet1', viewport }),
    magic: GRID_SCENE_PACKET_V2_MAGIC,
    paneId: 'body',
    rectCount: 1,
    rectInstances: new Float32Array(GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT),
    rects: new Float32Array(GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT),
    rectSignature: 'rect-one',
    requestSeq: 10,
    sheetName: 'Sheet1',
    surfaceSize: { height: 220, width: 480 },
    textCount: 1,
    textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
    textRuns: [
      {
        align: 'left',
        clipHeight: 16,
        clipWidth: 80,
        clipX: 4,
        clipY: 3,
        color: '#1f2933',
        font: '400 11px sans-serif',
        fontSize: 11,
        height: 16,
        strike: false,
        text: 'A1',
        underline: false,
        width: 80,
        wrap: false,
        x: 4,
        y: 3,
      },
    ],
    textSignature: 'text-one',
    version: GRID_SCENE_PACKET_V2_VERSION,
    viewport,
    ...overrides,
  }
}

describe('grid scene packet v2', () => {
  test('accepts well-formed typed scene packets', () => {
    const packet = createPacket()
    packet.rectInstances.set([0, 0, 104, 22, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 480, 220])
    packet.rects.set([0, 0, 104, 22, 1, 1, 1, 1])
    packet.textMetrics.set([4, 3, 80, 16, 0, 0, 0, 0])

    expect(validateGridScenePacketV2(packet)).toEqual({ ok: true })
  })

  test('rejects malformed packet metadata and buffers', () => {
    expect(validateGridScenePacketV2(createPacket({ rects: new Float32Array(1) }))).toEqual({
      ok: false,
      reason: 'rect buffer too small',
    })
    expect(validateGridScenePacketV2(createPacket({ rectInstances: new Float32Array(1) }))).toEqual({
      ok: false,
      reason: 'rect instance buffer too small',
    })
    expect(validateGridScenePacketV2(createPacket({ viewport: { colEnd: 1, colStart: 2, rowEnd: 1, rowStart: 0 } }))).toEqual({
      ok: false,
      reason: 'empty viewport',
    })
    expect(validateGridScenePacketV2(createPacket({ rects: new Float32Array([0, 0, Number.NaN, 22, 1, 1, 1, 1]) }))).toEqual({
      ok: false,
      reason: 'bad rect size',
    })
    expect(validateGridScenePacketV2(createPacket({ requestSeq: -1 }))).toEqual({
      ok: false,
      reason: 'bad request sequence',
    })
    expect(validateGridScenePacketV2(createPacket({ cameraSeq: -1 }))).toEqual({
      ok: false,
      reason: 'bad camera sequence',
    })
    expect(validateGridScenePacketV2(createPacket({ generatedAt: Number.NaN }))).toEqual({
      ok: false,
      reason: 'bad generated timestamp',
    })
  })

  test('rejects mismatched tile packet keys', () => {
    expect(
      validateGridScenePacketV2(
        createPacket({
          key: createGridTileKeyV2({
            paneId: 'body',
            sheetName: 'Sheet2',
            viewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
          }),
        }),
      ),
    ).toEqual({ ok: false, reason: 'tile key sheet mismatch' })
    expect(
      validateGridScenePacketV2(
        createPacket({
          key: createGridTileKeyV2({
            paneId: 'left',
            sheetName: 'Sheet1',
            viewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
          }),
        }),
      ),
    ).toEqual({ ok: false, reason: 'tile key pane mismatch' })
  })

  test('serializes full revision identity and rejects stale-valid reuse across data revisions', () => {
    const base = createGridTileKeyV2({
      paneId: 'body',
      sheetName: 'Sheet1',
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      valueVersion: 1,
      styleVersion: 1,
    })
    const desired = { ...base, valueVersion: 4, styleVersion: 3 }

    expect(serializeGridTileKeyV2(desired)).not.toBe(serializeGridTileKeyV2(base))
    expect(isStaleValidGridTileKeyV2(base, desired)).toBe(false)
    expect(isStaleValidGridTileKeyV2(base, { ...base, rowStart: 4, rowEnd: 12 })).toBe(true)
    expect(isStaleValidGridTileKeyV2({ ...base, axisVersionX: 1 }, desired)).toBe(false)
  })
})
