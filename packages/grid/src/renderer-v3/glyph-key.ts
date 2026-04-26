export type GlyphIdV3 = number

export interface GlyphKeyV3 {
  readonly fontInternId: number
  readonly dprBucket: number
  readonly glyph: string
}

export interface GlyphKeyRegistryStatsV3 {
  readonly glyphCount: number
}

export class GlyphKeyRegistryV3 {
  private readonly idsByKey = new Map<string, GlyphIdV3>()
  private readonly keysById = new Map<GlyphIdV3, GlyphKeyV3>()
  private nextGlyphId = 1

  stats(): GlyphKeyRegistryStatsV3 {
    return {
      glyphCount: this.idsByKey.size,
    }
  }

  intern(key: GlyphKeyV3): GlyphIdV3 {
    const encoded = encodeGlyphKeyV3(key)
    const existing = this.idsByKey.get(encoded)
    if (existing !== undefined) {
      return existing
    }

    const glyphId = this.nextGlyphId++
    this.idsByKey.set(encoded, glyphId)
    this.keysById.set(glyphId, { ...key })
    return glyphId
  }

  resolve(glyphId: GlyphIdV3): GlyphKeyV3 | null {
    return this.keysById.get(glyphId) ?? null
  }
}

export function encodeGlyphKeyV3(key: GlyphKeyV3): string {
  return `${key.fontInternId}:${key.dprBucket}:${key.glyph}`
}
