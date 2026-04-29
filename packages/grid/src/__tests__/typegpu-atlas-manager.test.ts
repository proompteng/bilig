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
    expect(atlas.getGlyphGeometryVersion()).toBe(0)
    atlas.intern('400 11px Geist', 'A')
    expect(atlas.getVersion()).toBeGreaterThan(0)
    expect(atlas.getGlyphGeometryVersion()).toBe(0)
  })

  it('bumps glyph geometry version only when atlas texture growth remaps UVs', () => {
    const originalOffscreenCanvas = globalThis.OffscreenCanvas
    Object.defineProperty(globalThis, 'OffscreenCanvas', {
      configurable: true,
      value: class {
        height: number
        width: number
        constructor(width: number, height: number) {
          this.width = width
          this.height = height
        }
        getContext() {
          return null
        }
      },
    })
    try {
      const atlas = createGlyphAtlas({ initialHeight: 512, initialWidth: 512 })
      let glyph = 0
      const firstGeometryVersion = atlas.getGlyphGeometryVersion()

      while (atlas.getGlyphGeometryVersion() === firstGeometryVersion && glyph < 4000) {
        atlas.intern('400 48px Geist', `glyph-${glyph}`)
        glyph += 1
      }

      expect(atlas.getGlyphGeometryVersion()).toBeGreaterThan(firstGeometryVersion)
    } finally {
      Object.defineProperty(globalThis, 'OffscreenCanvas', {
        configurable: true,
        value: originalOffscreenCanvas,
      })
    }
  })

  it('tracks dirty atlas pages for V3 glyph inserts', () => {
    const atlas = createGlyphAtlas()

    const entry = atlas.intern('400 11px Geist', 'A')

    const stats = atlas.getDirtyPageStats()
    expect(atlas.getTextAtlasPagesSeq()).toBeGreaterThan(0)
    expect(stats.dirtyPageCount).toBeGreaterThan(0)
    expect(stats.dirtyUploadBytes).toBeGreaterThan(0)
    expect(atlas.getTextAtlasPagesStats()).toMatchObject({
      dirtyPageCount: stats.dirtyPageCount,
      glyphCount: 1,
    })
    expect(atlas.resolveGlyphRecord(entry.glyphId)).toMatchObject({
      glyphId: entry.glyphId,
      pageId: entry.pageId,
    })

    const pages = atlas.drainDirtyPages()
    expect(pages.length).toBe(stats.dirtyPageCount)
    expect(pages.every((page) => page.byteSize === page.width * page.height * 4)).toBe(true)
    expect(atlas.getDirtyPageStats().dirtyPageCount).toBe(0)
    expect(atlas.getTextAtlasPagesStats().dirtyPageCount).toBe(0)
  })

  it('assigns stable glyph identities without ref-counting repeated reads as new glyph registrations', () => {
    const atlas = createGlyphAtlas()

    const first = atlas.intern('400 11px Geist', 'A')
    atlas.drainDirtyPages()
    const second = atlas.intern('400 11px Geist', 'A')

    expect(second.glyphId).toBe(first.glyphId)
    expect(second.pageId).toBe(first.pageId)
    expect(atlas.getTextAtlasPagesStats()).toMatchObject({
      dirtyPageCount: 0,
      glyphCount: 1,
    })
  })
})
