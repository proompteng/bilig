// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { createGlyphAtlas } from '../renderer/glyph-atlas.js'

describe('glyph-atlas', () => {
  it('returns stable glyph keys for repeated runs', () => {
    const atlas = createGlyphAtlas()
    const first = atlas.intern('400 11px Geist', 'A')
    const second = atlas.intern('400 11px Geist', 'A')

    expect(second.key).toBe(first.key)
    expect(second.x).toBe(first.x)
    expect(second.y).toBe(first.y)
  })

  it('tracks atlas version after adding glyphs', () => {
    const atlas = createGlyphAtlas()
    expect(atlas.getVersion()).toBe(0)
    atlas.intern('400 11px Geist', 'A')
    expect(atlas.getVersion()).toBeGreaterThan(0)
  })
})
