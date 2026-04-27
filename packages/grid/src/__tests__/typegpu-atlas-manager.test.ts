// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { createGlyphAtlas } from '../renderer-v3/typegpu-atlas-manager.js'

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

  it('tracks dirty atlas pages for V3 glyph inserts', () => {
    const atlas = createGlyphAtlas()

    atlas.intern('400 11px Geist', 'A')

    const stats = atlas.getDirtyPageStats()
    expect(stats.dirtyPageCount).toBeGreaterThan(0)
    expect(stats.dirtyUploadBytes).toBeGreaterThan(0)

    const pages = atlas.drainDirtyPages()
    expect(pages.length).toBe(stats.dirtyPageCount)
    expect(pages.every((page) => page.byteSize === page.width * page.height * 4)).toBe(true)
    expect(atlas.getDirtyPageStats().dirtyPageCount).toBe(0)
  })
})
