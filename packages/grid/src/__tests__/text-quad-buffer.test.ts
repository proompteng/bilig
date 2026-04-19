import { describe, expect, it } from 'vitest'
import { buildTextQuads } from '../renderer/text-quad-buffer.js'

const atlas = {
  intern(font: string, glyph: string) {
    void font
    return {
      key: `atlas:${glyph}`,
      font,
      glyph,
      x: 0,
      y: 0,
      width: 8,
      height: 12,
      advance: 8,
      baseline: 10,
      u0: 0,
      v0: 0,
      u1: 1,
      v1: 1,
    }
  },
}

describe('text-quad-buffer', () => {
  it('builds one quad per glyph with atlas uv coordinates', () => {
    const quads = buildTextQuads(
      [
        {
          text: 'AB',
          x: 10,
          y: 20,
          font: '400 11px Geist',
        },
      ],
      atlas,
    )

    expect(quads).toHaveLength(2)
    expect(quads[0]).toMatchObject({
      atlasKey: 'atlas:A',
      x: 18,
      y: 22.5,
      width: 8,
      height: 12,
    })
    expect(quads[1]).toMatchObject({
      atlasKey: 'atlas:B',
      x: 26,
      y: 22.5,
    })
  })
})
