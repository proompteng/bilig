import { describe, expect, it } from 'vitest'
import {
  buildTextDecorationRectsFromScene,
  buildTextQuads,
  buildTextQuadsFromRunsWithSpans,
  buildTextQuadsFromScene,
} from '../renderer-v3/line-text-quad-buffer.js'

const atlas = {
  intern(font: string, glyph: string) {
    void font
    const advance = Math.max(0, glyph.length * 8)
    return {
      key: `atlas:${glyph}`,
      glyphId: glyph.codePointAt(0) ?? 0,
      pageId: 1,
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
  it('builds one quad per resolved glyph with atlas uv coordinates', () => {
    const internedGlyphs: string[] = []
    const recordingAtlas = {
      intern(font: string, glyph: string) {
        internedGlyphs.push(glyph)
        return atlas.intern(font, glyph)
      },
    }
    const quads = buildTextQuads(
      [
        {
          text: 'AB',
          x: 10,
          y: 20,
          font: '400 11px Geist',
        },
      ],
      recordingAtlas,
    )

    expect(quads).toHaveLength(2)
    expect(quads[0]).toMatchObject({
      atlasKey: 'atlas:A',
      x: 18,
      y: 24.4,
      width: 8,
      height: 12,
      clipX: 18,
      clipY: 23,
      clipWidth: 16,
      clipHeight: 16,
    })
    expect(quads[0]?.glyph).toBe('A')
    expect(quads[1]).toMatchObject({ atlasKey: 'atlas:B', glyph: 'B', x: 26, y: 24.4 })
    expect(internedGlyphs).not.toContain('AB')
  })

  it('wraps and centers lines inside the clipped layout box', () => {
    const quads = buildTextQuads(
      [
        {
          text: 'AAA BBB',
          x: 0,
          y: 10,
          width: 40,
          height: 40,
          wrap: true,
          align: 'center',
          font: '400 10px Geist',
          fontSize: 10,
        },
      ],
      atlas,
    )

    expect(quads).toHaveLength(6)
    expect(quads[0]).toMatchObject({ x: 8, y: 13 })
    expect(quads[3]).toMatchObject({ x: 8, y: 25 })
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

    expect(quadCount).toBe(2)
    expect(floats[0]).toBe(8)
    expect(floats[1]).toBe(4)
    expect(floats[16]).toBe(16)
    expect(floats[17]).toBe(4)
    expect(floats[12]).toBe(10)
    expect(floats[13]).toBe(3)
    expect(floats[14]).toBe(32)
    expect(floats[15]).toBe(17)
    expect(floats[8]).toBeCloseTo(0x11 / 255)
    expect(floats[9]).toBeCloseTo(0x22 / 255)
    expect(floats[10]).toBeCloseTo(0x33 / 255)
    expect(floats[11]).toBe(1)
  })

  it('records text-run to gpu-quad spans for partial text uploads', () => {
    const payload = buildTextQuadsFromRunsWithSpans(
      [
        { text: 'AB', x: 0, y: 0, font: '400 10px Geist', fontSize: 10 },
        { text: 'C', x: 40, y: 0, font: '400 10px Geist', fontSize: 10 },
      ],
      atlas,
    )

    expect(payload.quadCount).toBe(3)
    expect(payload.runSpans).toEqual([
      { offset: 0, length: 2 },
      { offset: 2, length: 1 },
    ])
    expect(payload.glyphIds).toEqual(['A', 'B', 'C'].map((glyph) => glyph.codePointAt(0)))
    expect(payload.runGlyphIds).toEqual([['A', 'B'].map((glyph) => glyph.codePointAt(0)), ['C'].map((glyph) => glyph.codePointAt(0))])
  })

  it('reuses clean text-run payloads while rebuilding dirty runs', () => {
    const internedGlyphs: string[] = []
    const recordingAtlas = {
      getVersion: () => 0,
      intern(font: string, glyph: string) {
        internedGlyphs.push(glyph)
        return atlas.intern(font, glyph)
      },
    }
    const first = buildTextQuadsFromRunsWithSpans(
      [
        { text: 'AB', x: 0, y: 0, font: '400 10px Geist', fontSize: 10 },
        { text: 'C', x: 40, y: 0, font: '400 10px Geist', fontSize: 10 },
      ],
      recordingAtlas,
    )

    internedGlyphs.length = 0
    const second = buildTextQuadsFromRunsWithSpans(
      [
        { text: 'AB', x: 0, y: 0, font: '400 10px Geist', fontSize: 10 },
        { text: 'D', x: 40, y: 0, font: '400 10px Geist', fontSize: 10 },
      ],
      recordingAtlas,
      undefined,
      {
        dirtyRunSpans: [{ offset: 1, length: 1 }],
        previousRunPayloads: first.runPayloads,
      },
    )

    expect(second.quadCount).toBe(3)
    expect(second.runSpans).toEqual([
      { offset: 0, length: 2 },
      { offset: 2, length: 1 },
    ])
    expect(internedGlyphs).not.toContain('A')
    expect(internedGlyphs).not.toContain('B')
    expect(internedGlyphs).toContain('D')
  })

  it('does not reuse text-run payloads across atlas version changes', () => {
    let atlasVersion = 1
    const internedGlyphs: string[] = []
    const recordingAtlas = {
      getVersion: () => atlasVersion,
      intern(font: string, glyph: string) {
        internedGlyphs.push(glyph)
        return atlas.intern(font, glyph)
      },
    }
    const first = buildTextQuadsFromRunsWithSpans([{ text: 'AB', x: 0, y: 0, font: '400 10px Geist', fontSize: 10 }], recordingAtlas)

    atlasVersion = 2
    internedGlyphs.length = 0
    const second = buildTextQuadsFromRunsWithSpans(
      [{ text: 'AB', x: 0, y: 0, font: '400 10px Geist', fontSize: 10 }],
      recordingAtlas,
      undefined,
      { previousRunPayloads: first.runPayloads },
    )

    expect(second.runPayloads[0]?.atlasVersion).toBe(2)
    expect(internedGlyphs).toContain('A')
    expect(internedGlyphs).toContain('B')
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
    expect(rects[0]).toMatchObject({ x: 10, width: 14, height: 1, color: '#112233' })
    expect(rects[0]?.y).toBeCloseTo(14.7)
    expect(rects[1]).toMatchObject({ x: 10, width: 14, height: 1, color: '#112233' })
    expect(rects[1]?.y).toBeCloseTo(10.5)
  })
})
