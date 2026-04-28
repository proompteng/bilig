import { describe, expect, test } from 'vitest'
import { resolveTextClipRect, resolveTextDecorationRects, resolveTextLineLayouts } from '../renderer-v3/line-text-layout.js'
import type { GlyphAtlasEntry } from '../renderer-v3/typegpu-atlas-manager.js'

const atlas = {
  intern(font: string, glyph: string): GlyphAtlasEntry {
    return {
      advance: glyph.length * 10,
      baseline: 10,
      font,
      glyph,
      glyphId: glyph.codePointAt(0) ?? 0,
      height: 12,
      key: `${font}:${glyph}`,
      originOffsetX: 0,
      pageId: 1,
      u0: 0,
      u1: 1,
      v0: 0,
      v1: 1,
      width: glyph.length * 10,
      x: 0,
      y: 0,
    }
  },
}

describe('gridTextLayout', () => {
  test('resolves right-aligned single-line text', () => {
    const lines = resolveTextLineLayouts(
      {
        align: 'right',
        font: '400 10px sans-serif',
        height: 20,
        text: '123',
        width: 100,
        x: 10,
        y: 5,
      },
      atlas,
    )

    expect(lines).toEqual([{ text: '123', width: 30, x: 72, y: 10 }])
  })

  test('resolves clip rectangles from explicit insets', () => {
    const run = {
      clipHeight: 18,
      clipWidth: 80,
      clipX: 12,
      clipY: 9,
      text: 'abc',
      x: 10,
      y: 5,
    }
    const lines = resolveTextLineLayouts(run, atlas)

    expect(resolveTextClipRect(run, lines)).toEqual({ height: 18, width: 80, x: 12, y: 9 })
  })

  test('wraps by measured width', () => {
    const lines = resolveTextLineLayouts(
      {
        font: '400 10px sans-serif',
        text: 'alpha beta',
        width: 66,
        wrap: true,
        x: 0,
        y: 0,
      },
      atlas,
    )

    expect(lines.map((line) => line.text)).toEqual(['alpha', 'beta'])
  })

  test('resolves decorations from the same line layout model', () => {
    expect(
      resolveTextDecorationRects(
        {
          color: '#123456',
          font: '400 10px sans-serif',
          height: 20,
          text: 'abc',
          underline: true,
          width: 100,
          x: 10,
          y: 5,
        },
        atlas,
      ),
    ).toEqual([{ color: '#123456', height: 1, width: 30, x: 18, y: 13.6 }])
  })
})
