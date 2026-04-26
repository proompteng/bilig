import { describe, expect, test } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  createGridTileKeyV2,
  type GridScenePacketV2,
} from '../renderer-v2/scene-packet-v2.js'
import { resolveGridRectPacketSignature, resolveGridTextPacketSignature } from '../renderer-v2/typegpu-buffer-pool.js'

function createPacket(overrides: Partial<GridScenePacketV2> = {}): GridScenePacketV2 {
  const viewport = { colEnd: 1, colStart: 0, rowEnd: 1, rowStart: 0 }
  return {
    borderRectCount: 0,
    cameraSeq: 1,
    fillRectCount: 1,
    generatedAt: 1,
    generation: 1,
    key: createGridTileKeyV2({ paneId: 'body', sheetName: 'Sheet1', viewport }),
    magic: GRID_SCENE_PACKET_V2_MAGIC,
    paneId: 'body',
    rectCount: 1,
    rectInstances: new Float32Array([0, 0, 104, 22, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 200, 100]),
    rects: new Float32Array([0, 0, 104, 22, 1, 0, 0, 1]),
    rectSignature: 'rect-one',
    requestSeq: 1,
    sheetName: 'Sheet1',
    surfaceSize: { height: 100, width: 200 },
    textCount: 0,
    textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
    textRuns: [],
    textSignature: 'text-empty',
    version: GRID_SCENE_PACKET_V2_VERSION,
    viewport,
    ...overrides,
  }
}

function rectSignature(packet: GridScenePacketV2): string {
  return resolveGridRectPacketSignature({
    frameHeight: 100,
    frameWidth: 200,
    packedScene: packet,
  })
}

describe('typegpu resource cache signatures', () => {
  test('keeps equivalent packed text packets stable without relying on object identity', () => {
    const first = createPacket({
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 22,
          clipWidth: 104,
          clipX: 0,
          clipY: 0,
          color: '#111111',
          font: '400 11px sans-serif',
          fontSize: 11,
          height: 22,
          strike: false,
          text: 'A1',
          underline: false,
          width: 104,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
      textSignature: 'text-a1',
    })
    const second = createPacket({
      textCount: first.textCount,
      textRuns: first.textRuns.map((run) => ({ ...run })),
      textSignature: first.textSignature,
    })
    const changed = createPacket({
      textCount: first.textCount,
      textRuns: first.textRuns.map((run) => ({ ...run, text: 'A2' })),
      textSignature: 'text-a2',
    })

    expect(resolveGridTextPacketSignature(first)).toBe(resolveGridTextPacketSignature(second))
    expect(resolveGridTextPacketSignature(changed)).not.toBe(resolveGridTextPacketSignature(first))
  })

  test('keeps resource signatures stable across packet sequence churn', () => {
    const base = createPacket()
    const newerPacketWithSameContent = createPacket({
      cameraSeq: 22,
      generatedAt: 22,
      generation: 22,
      requestSeq: 22,
    })

    expect(rectSignature(newerPacketWithSameContent)).toBe(rectSignature(base))
    expect(resolveGridTextPacketSignature(newerPacketWithSameContent)).toBe(resolveGridTextPacketSignature(base))
  })

  test('includes packed rect and decoration content in rect signatures', () => {
    const base = createPacket()
    const changedRects = createPacket({ rectSignature: 'rect-two' })
    const changedDecorations = createPacket({ textSignature: 'text-underlined' })

    expect(rectSignature(changedRects)).not.toBe(rectSignature(base))
    expect(rectSignature(changedDecorations)).not.toBe(rectSignature(base))
    expect(
      resolveGridRectPacketSignature({
        decorationRects: [{ color: '#111111', height: 1, width: 20, x: 4, y: 18 }],
        frameHeight: 100,
        frameWidth: 200,
        packedScene: base,
      }),
    ).not.toBe(rectSignature(base))
    expect(base.rects.length).toBe(GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT)
    expect(base.rectInstances.length).toBe(GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT)
  })
})
