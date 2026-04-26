import { describe, expect, it } from 'vitest'
import { GlyphKeyRegistryV3 } from '../renderer-v3/glyph-key.js'
import { TextAtlasPagesV3 } from '../renderer-v3/text-atlas-pages.js'

describe('renderer-v3 text atlas pages', () => {
  it('interns glyph keys and keeps glyph IDs stable', () => {
    const registry = new GlyphKeyRegistryV3()

    const first = registry.intern({ dprBucket: 1, fontInternId: 7, glyph: 'A' })
    const second = registry.intern({ dprBucket: 1, fontInternId: 7, glyph: 'A' })
    const third = registry.intern({ dprBucket: 2, fontInternId: 7, glyph: 'A' })

    expect(second).toBe(first)
    expect(third).not.toBe(first)
    expect(registry.resolve(first)).toEqual({ dprBucket: 1, fontInternId: 7, glyph: 'A' })
    expect(registry.stats().glyphCount).toBe(2)
  })

  it('marks only pages touched by new glyph records as dirty uploads', () => {
    const atlas = new TextAtlasPagesV3()
    atlas.upsertPage({ height: 256, pageId: 1, width: 256 })
    atlas.upsertPage({ height: 128, pageId: 2, width: 128 })

    const glyph = atlas.registerGlyph({ glyphId: 4, pageId: 1, u0: 0, u1: 0.1, v0: 0, v1: 0.1 })

    expect(glyph.refCount).toBe(1)
    expect(atlas.stats()).toMatchObject({
      dirtyPageCount: 1,
      dirtyUploadBytes: 256 * 256 * 4,
      glyphCount: 1,
      pageCount: 2,
    })

    const uploads = atlas.drainDirtyPages()
    expect(uploads.map((upload) => upload.page.pageId)).toEqual([1])
    expect(atlas.stats().dirtyUploadBytes).toBe(0)

    atlas.registerGlyph({ glyphId: 5, pageId: 2, u0: 0.2, u1: 0.3, v0: 0.2, v1: 0.3 })
    expect(atlas.drainDirtyPages().map((upload) => upload.page.pageId)).toEqual([2])
  })

  it('keeps existing glyph UVs immutable and reference counted', () => {
    const atlas = new TextAtlasPagesV3()
    atlas.upsertPage({ height: 64, pageId: 1, width: 64 })

    const glyph = atlas.registerGlyph({ glyphId: 8, pageId: 1, u0: 0, u1: 1, v0: 0, v1: 1 })
    atlas.drainDirtyPages()

    expect(atlas.registerGlyph({ glyphId: 8, pageId: 1, u0: 0, u1: 1, v0: 0, v1: 1 })).toBe(glyph)
    expect(glyph.refCount).toBe(2)
    expect(atlas.stats().dirtyPageCount).toBe(0)
    expect(() => atlas.registerGlyph({ glyphId: 8, pageId: 1, u0: 0, u1: 0.5, v0: 0, v1: 1 })).toThrow(/immutable/)

    expect(atlas.releaseGlyph(8)).toBe(false)
    expect(atlas.resolveGlyph(8)).toBe(glyph)
    expect(atlas.releaseGlyph(8)).toBe(true)
    expect(atlas.resolveGlyph(8)).toBeNull()
  })

  it('respects dirty page upload budgets without blanking all pages', () => {
    const atlas = new TextAtlasPagesV3()
    atlas.upsertPage({ height: 64, pageId: 1, width: 64, dirty: true })
    atlas.upsertPage({ height: 64, pageId: 2, width: 64, dirty: true })

    const uploads = atlas.drainDirtyPages(64 * 64 * 4)

    expect(uploads.map((upload) => upload.page.pageId)).toEqual([1])
    expect(atlas.stats().dirtyPageCount).toBe(1)
    expect(atlas.drainDirtyPages().map((upload) => upload.page.pageId)).toEqual([2])
  })
})
