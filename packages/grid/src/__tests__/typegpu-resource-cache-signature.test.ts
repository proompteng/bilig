import { describe, expect, test } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
} from '../renderer-v2/scene-packet-v2.js'
import {
  resolveGridRectSceneSignature,
  resolveGridTextSceneSignature,
  shouldDeferPaneTextUpload,
} from '../renderer-v2/typegpu-buffer-pool.js'

describe('typegpu resource cache signatures', () => {
  test('keeps equivalent text scenes stable without relying on object identity', () => {
    const first = {
      items: [
        {
          align: 'left' as const,
          clipInsetBottom: 0,
          clipInsetLeft: 0,
          clipInsetRight: 0,
          clipInsetTop: 0,
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
    }
    const second = { items: first.items.map((item) => ({ ...item })) }
    const changed = { items: [{ ...first.items[0], text: 'A2' }] }

    expect(resolveGridTextSceneSignature(first)).toBe(resolveGridTextSceneSignature(second))
    expect(resolveGridTextSceneSignature(changed)).not.toBe(resolveGridTextSceneSignature(first))
  })

  test('includes pane dimensions and decoration rects in rect signatures', () => {
    const scene = {
      borderRects: [],
      fillRects: [{ color: { a: 1, b: 0, g: 0, r: 1 }, height: 22, width: 104, x: 0, y: 0 }],
    }

    expect(
      resolveGridRectSceneSignature({
        frame: { height: 100, width: 200, x: 0, y: 0 },
        scene,
      }),
    ).not.toBe(
      resolveGridRectSceneSignature({
        decorationRects: [{ color: '#111111', height: 1, width: 20, x: 4, y: 18 }],
        frame: { height: 100, width: 200, x: 0, y: 0 },
        scene,
      }),
    )
  })

  test('uses packed rect data in rect signatures', () => {
    const scene = {
      borderRects: [],
      fillRects: [],
    }
    const basePacket = {
      borderRectCount: 0,
      fillRectCount: 1,
      generation: 1,
      magic: GRID_SCENE_PACKET_V2_MAGIC,
      paneId: 'body' as const,
      rectCount: 1,
      rectInstances: new Float32Array([0, 0, 10, 10, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 100]),
      rects: new Float32Array([0, 0, 10, 10, 1, 0, 0, 1]),
      sheetName: 'Sheet1',
      surfaceSize: { height: 100, width: 100 },
      textCount: 0,
      textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
      version: GRID_SCENE_PACKET_V2_VERSION,
      viewport: { colEnd: 1, colStart: 0, rowEnd: 1, rowStart: 0 },
    }
    const changedPacket = {
      ...basePacket,
      rectInstances: new Float32Array([0, 0, 11, 10, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 100]),
      rects: new Float32Array([0, 0, 11, 10, 1, 0, 0, 1]),
    }
    const newerPacketWithSameRects = {
      ...basePacket,
      generation: 2,
    }

    expect(
      resolveGridRectSceneSignature({
        frame: { height: 100, width: 200, x: 0, y: 0 },
        packedScene: basePacket,
        scene,
      }),
    ).not.toBe(
      resolveGridRectSceneSignature({
        frame: { height: 100, width: 200, x: 0, y: 0 },
        packedScene: changedPacket,
        scene,
      }),
    )
    expect(
      resolveGridRectSceneSignature({
        frame: { height: 100, width: 200, x: 0, y: 0 },
        packedScene: basePacket,
        scene,
      }),
    ).toBe(
      resolveGridRectSceneSignature({
        frame: { height: 100, width: 200, x: 0, y: 0 },
        packedScene: newerPacketWithSameRects,
        scene,
      }),
    )
    expect(basePacket.rects.length).toBe(GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT)
    expect(basePacket.rectInstances.length).toBe(GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT)
  })

  test('does not defer the first text upload for a newly visible packed pane', () => {
    expect(
      shouldDeferPaneTextUpload({
        currentTextCount: 0,
        currentTextSignature: null,
        deferTextUploads: true,
        hasPackedScene: true,
        hasTextBuffer: false,
        nextTextItemCount: 12,
      }),
    ).toBe(false)
  })

  test('defers packed pane text updates only when a resident text resource can keep drawing', () => {
    expect(
      shouldDeferPaneTextUpload({
        currentTextCount: 12,
        currentTextSignature: 'resident-text',
        deferTextUploads: true,
        hasPackedScene: true,
        hasTextBuffer: true,
        nextTextItemCount: 12,
      }),
    ).toBe(true)

    expect(
      shouldDeferPaneTextUpload({
        currentTextCount: 0,
        currentTextSignature: 'empty-text',
        deferTextUploads: true,
        hasPackedScene: true,
        hasTextBuffer: false,
        nextTextItemCount: 12,
      }),
    ).toBe(false)
  })

  test('keeps non-packed panes on the immediate text upload path', () => {
    expect(
      shouldDeferPaneTextUpload({
        currentTextCount: 12,
        currentTextSignature: 'resident-text',
        deferTextUploads: true,
        hasPackedScene: false,
        hasTextBuffer: true,
        nextTextItemCount: 12,
      }),
    ).toBe(false)
  })
})
