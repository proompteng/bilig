export interface GlyphAtlasLocationV2 {
  readonly pageId: number
  readonly generation: number
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export class GlyphAtlasV2 {
  private generation = 0
  private readonly glyphs = new Map<string, GlyphAtlasLocationV2>()
  private readonly missing = new Set<string>()

  getGeneration(): number {
    return this.generation
  }

  resolve(key: string): GlyphAtlasLocationV2 | null {
    return this.glyphs.get(key) ?? null
  }

  queueMissing(key: string): void {
    if (!this.glyphs.has(key)) {
      this.missing.add(key)
    }
  }

  drainMissing(): readonly string[] {
    const next = [...this.missing]
    this.missing.clear()
    return next
  }

  registerGlyph(key: string, location: Omit<GlyphAtlasLocationV2, 'generation'>): GlyphAtlasLocationV2 {
    const next = { ...location, generation: this.generation + 1 }
    this.generation = next.generation
    this.glyphs.set(key, next)
    this.missing.delete(key)
    return next
  }
}
