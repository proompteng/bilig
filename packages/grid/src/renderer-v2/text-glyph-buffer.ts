import type { ResolvedCellTextLayout } from '../text/gridTextPacket.js'
import type { GlyphAtlasV2 } from './glyphAtlasV2.js'

export const TEXT_GLYPH_INSTANCE_FLOAT_COUNT = 12

export function buildTextGlyphInstanceBuffer(input: {
  readonly layouts: readonly ResolvedCellTextLayout[]
  readonly atlas: GlyphAtlasV2
}): { readonly floats: Float32Array; readonly glyphCount: number; readonly missingGlyphCount: number } {
  const placements = input.layouts.flatMap((layout) => layout.lines.flatMap((line) => line.glyphs))
  const floats = new Float32Array(Math.max(1, placements.length) * TEXT_GLYPH_INSTANCE_FLOAT_COUNT)
  let missingGlyphCount = 0
  placements.forEach((glyph, index) => {
    const location = input.atlas.resolve(glyph.atlasGlyphKey)
    if (!location) {
      input.atlas.queueMissing(glyph.atlasGlyphKey)
      missingGlyphCount += 1
    }
    const offset = index * TEXT_GLYPH_INSTANCE_FLOAT_COUNT
    floats[offset + 0] = glyph.worldX
    floats[offset + 1] = glyph.worldY
    floats[offset + 2] = glyph.width
    floats[offset + 3] = glyph.height
    floats[offset + 4] = location?.pageId ?? -1
    floats[offset + 5] = location?.generation ?? -1
    floats[offset + 6] = location?.x ?? 0
    floats[offset + 7] = location?.y ?? 0
    floats[offset + 8] = location?.width ?? 0
    floats[offset + 9] = location?.height ?? 0
    floats[offset + 10] = glyph.advance
    floats[offset + 11] = glyph.uvPadding
  })
  return {
    floats,
    glyphCount: placements.length,
    missingGlyphCount,
  }
}
