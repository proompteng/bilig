import { describe, expect, test } from 'vitest'
import { resolveGridRectSceneSignature, resolveGridTextSceneSignature } from '../renderer/typegpu-resource-cache.js'

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
})
