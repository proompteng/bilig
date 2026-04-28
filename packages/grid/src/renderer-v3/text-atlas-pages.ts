import type { GlyphIdV3 } from './glyph-key.js'

export interface AtlasPageV3 {
  readonly pageId: number
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  seq: number
  dirty: boolean
  lastUsedGeneration: number
}

export interface GlyphRecordV3 {
  readonly glyphId: GlyphIdV3
  readonly pageId: number
  u0: number
  v0: number
  u1: number
  v1: number
  refCount: number
}

export interface AtlasDirtyPageUploadV3 {
  readonly page: AtlasPageV3
  readonly byteSize: number
}

export interface TextAtlasPagesStatsV3 {
  readonly atlasSeq: number
  readonly pageCount: number
  readonly glyphCount: number
  readonly dirtyPageCount: number
  readonly dirtyUploadBytes: number
}

export class TextAtlasPagesV3 {
  private readonly pages = new Map<number, AtlasPageV3>()
  private readonly glyphs = new Map<GlyphIdV3, GlyphRecordV3>()
  private generation = 0
  private atlasSeq = 0

  get seq(): number {
    return this.atlasSeq
  }

  stats(): TextAtlasPagesStatsV3 {
    let dirtyPageCount = 0
    let dirtyUploadBytes = 0
    for (const page of this.pages.values()) {
      if (!page.dirty) {
        continue
      }
      dirtyPageCount += 1
      dirtyUploadBytes += pageUploadBytes(page)
    }
    return {
      atlasSeq: this.atlasSeq,
      dirtyPageCount,
      dirtyUploadBytes,
      glyphCount: this.glyphs.size,
      pageCount: this.pages.size,
    }
  }

  upsertPage(input: {
    readonly pageId: number
    readonly x?: number | undefined
    readonly y?: number | undefined
    readonly width: number
    readonly height: number
    readonly dirty?: boolean | undefined
  }): AtlasPageV3 {
    const existing = this.pages.get(input.pageId)
    if (existing) {
      const geometryChanged =
        existing.x !== (input.x ?? existing.x) ||
        existing.y !== (input.y ?? existing.y) ||
        existing.width !== input.width ||
        existing.height !== input.height
      if (geometryChanged) {
        throw new Error(`Atlas page ${input.pageId} geometry is immutable`)
      }
      if (input.dirty) {
        this.markPageDirty(input.pageId)
      }
      return existing
    }

    const page: AtlasPageV3 = {
      dirty: input.dirty ?? false,
      height: Math.max(1, Math.ceil(input.height)),
      lastUsedGeneration: 0,
      pageId: input.pageId,
      seq: 0,
      width: Math.max(1, Math.ceil(input.width)),
      x: input.x ?? 0,
      y: input.y ?? 0,
    }
    this.pages.set(page.pageId, page)
    if (page.dirty) {
      this.bumpPageSeq(page)
    }
    return page
  }

  registerGlyph(input: {
    readonly glyphId: GlyphIdV3
    readonly pageId: number
    readonly u0: number
    readonly v0: number
    readonly u1: number
    readonly v1: number
  }): GlyphRecordV3 {
    const page = this.pages.get(input.pageId)
    if (!page) {
      throw new Error(`Atlas page ${input.pageId} has not been registered`)
    }

    const existing = this.glyphs.get(input.glyphId)
    if (existing) {
      if (
        existing.pageId !== input.pageId ||
        existing.u0 !== input.u0 ||
        existing.v0 !== input.v0 ||
        existing.u1 !== input.u1 ||
        existing.v1 !== input.v1
      ) {
        throw new Error(`Glyph ${input.glyphId} atlas location is immutable`)
      }
      existing.refCount += 1
      this.touchPage(page)
      return existing
    }

    const record: GlyphRecordV3 = {
      glyphId: input.glyphId,
      pageId: input.pageId,
      refCount: 1,
      u0: input.u0,
      u1: input.u1,
      v0: input.v0,
      v1: input.v1,
    }
    this.glyphs.set(record.glyphId, record)
    this.touchPage(page)
    page.dirty = true
    this.bumpPageSeq(page)
    return record
  }

  resolveGlyph(glyphId: GlyphIdV3): GlyphRecordV3 | null {
    const glyph = this.glyphs.get(glyphId) ?? null
    if (!glyph) {
      return null
    }
    const page = this.pages.get(glyph.pageId)
    if (page) {
      this.touchPage(page)
    }
    return glyph
  }

  updateGlyphLocation(input: {
    readonly glyphId: GlyphIdV3
    readonly pageId: number
    readonly u0: number
    readonly v0: number
    readonly u1: number
    readonly v1: number
  }): GlyphRecordV3 {
    const page = this.pages.get(input.pageId)
    if (!page) {
      throw new Error(`Atlas page ${input.pageId} has not been registered`)
    }
    const existing = this.glyphs.get(input.glyphId)
    if (!existing) {
      return this.registerGlyph(input)
    }
    if (existing.pageId !== input.pageId) {
      throw new Error(`Glyph ${input.glyphId} atlas page is immutable`)
    }
    // Atlas texture growth can change normalized UVs while the physical page and
    // glyph identity remain stable. Update the record and mark only that page
    // dirty so existing tile dependencies can be redrawn without full-atlas
    // invalidation.
    existing.u0 = input.u0
    existing.u1 = input.u1
    existing.v0 = input.v0
    existing.v1 = input.v1
    page.dirty = true
    this.bumpPageSeq(page)
    this.touchPage(page)
    return existing
  }

  releaseGlyph(glyphId: GlyphIdV3): boolean {
    const glyph = this.glyphs.get(glyphId)
    if (!glyph) {
      return false
    }
    glyph.refCount = Math.max(0, glyph.refCount - 1)
    if (glyph.refCount > 0) {
      return false
    }
    this.glyphs.delete(glyphId)
    return true
  }

  markPageDirty(pageId: number): void {
    const page = this.pages.get(pageId)
    if (!page) {
      throw new Error(`Atlas page ${pageId} has not been registered`)
    }
    if (page.dirty) {
      return
    }
    page.dirty = true
    this.bumpPageSeq(page)
  }

  drainDirtyPages(maxBytes = Number.POSITIVE_INFINITY): readonly AtlasDirtyPageUploadV3[] {
    const uploads: AtlasDirtyPageUploadV3[] = []
    let remaining = Math.max(0, maxBytes)
    for (const page of this.pages.values()) {
      if (!page.dirty) {
        continue
      }
      const byteSize = pageUploadBytes(page)
      if (uploads.length > 0 && byteSize > remaining) {
        break
      }
      page.dirty = false
      remaining -= byteSize
      uploads.push({ byteSize, page })
    }
    return uploads
  }

  private bumpPageSeq(page: AtlasPageV3): void {
    this.atlasSeq += 1
    page.seq = this.atlasSeq
  }

  private touchPage(page: AtlasPageV3): void {
    this.generation += 1
    page.lastUsedGeneration = this.generation
  }
}

function pageUploadBytes(page: Pick<AtlasPageV3, 'width' | 'height'>): number {
  return page.width * page.height * 4
}
