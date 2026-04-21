import { describe, expect, it } from 'vitest'
import { buildTextDecorationRectsFromScene, buildTextQuads, buildTextQuadsFromScene } from '../renderer-v2/line-text-quad-buffer.js'

const atlas = {
  intern(font: string, glyph: string) {
    void font
    const advance = Math.max(0, glyph.length * 8)
    return {
      key: `atlas:${glyph}`,
      font,
      glyph,
      x: 0,
      y: 0,
      originOffsetX: 0,
      width: advance,
      height: 12,
      advance,
      baseline: 10,
      u0: 0,
      v0: 0,
      u1: 1,
      v1: 1,
    }
  },
}

describe('text-quad-buffer', () => {
  it('builds one quad per resolved line with atlas uv coordinates', () => {
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

    expect(quads).toHaveLength(1)
    expect(quads[0]).toMatchObject({
      atlasKey: 'atlas:AB',
      x: 18,
      y: 22.5,
      width: 16,
      height: 12,
      clipX: 10,
      clipY: 20,
      clipWidth: 32,
      clipHeight: 16,
    })
    expect(quads[0]?.glyph).toBe('AB')
  })

  it('wraps and centers lines inside the clipped layout box', () => {
    const quads = buildTextQuads(
      [
        {
          text: 'AAA BBB',
          x: 0,
          y: 10,
          width: 40,
          height: 28,
          wrap: true,
          align: 'center',
          font: '400 10px Geist',
          fontSize: 10,
        },
      ],
      atlas,
    )

    expect(quads).toHaveLength(2)
    expect(quads[0]).toMatchObject({ x: 8, y: 14 })
    expect(quads[1]).toMatchObject({ x: 8, y: 26 })
  })

  it('packs clipped scene items into gpu text instance buffers', () => {
    const { floats, quadCount } = buildTextQuadsFromScene(
      [
        {
          x: 0,
          y: 0,
          width: 40,
          height: 20,
          clipInsetTop: 2,
          clipInsetRight: 6,
          clipInsetBottom: 0,
          clipInsetLeft: 10,
          text: 'AB',
          align: 'left',
          wrap: false,
          color: '#112233',
          font: '400 10px Geist',
          fontSize: 10,
          underline: false,
          strike: false,
        },
      ],
      atlas,
    )

    expect(quadCount).toBe(1)
    expect(floats[0]).toBe(18)
    expect(floats[1]).toBe(6)
    expect(floats[12]).toBe(10)
    expect(floats[13]).toBe(2)
    expect(floats[14]).toBe(34)
    expect(floats[15]).toBe(20)
    expect(floats[8]).toBeCloseTo(0x11 / 255)
    expect(floats[9]).toBeCloseTo(0x22 / 255)
    expect(floats[10]).toBeCloseTo(0x33 / 255)
    expect(floats[11]).toBe(1)
  })

  it('builds underline and strike rects from clipped scene items', () => {
    const rects = buildTextDecorationRectsFromScene(
      [
        {
          x: 0,
          y: 0,
          width: 40,
          height: 20,
          clipInsetTop: 2,
          clipInsetRight: 6,
          clipInsetBottom: 0,
          clipInsetLeft: 10,
          text: 'AB',
          align: 'left',
          wrap: false,
          color: '#112233',
          font: '400 10px Geist',
          fontSize: 10,
          underline: true,
          strike: true,
        },
      ],
      atlas,
    )

    expect(rects).toHaveLength(2)
    expect(rects[0]).toMatchObject({ x: 18, width: 16, height: 1, color: '#112233' })
    expect(rects[0]?.y).toBeCloseTo(9.6)
    expect(rects[1]).toMatchObject({ x: 18, width: 16, height: 1, color: '#112233' })
    expect(rects[1]?.y).toBeCloseTo(4.2)
  })
})
