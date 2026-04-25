import { describe, expect, test } from 'vitest'
import { resolveCellTextLayoutV2 } from '../text/gridTextLayoutV2.js'
import { createFallbackTextMetricsProvider } from '../text/gridTextMetrics.js'
import type { FontKey } from '../text/gridTextPacket.js'
import { GlyphAtlasV2 } from '../renderer-v2/glyphAtlasV2.js'
import { buildTextGlyphInstanceBuffer } from '../renderer-v2/text-glyph-buffer.js'

const fontKey: FontKey = {
  dprBucket: 2,
  family: 'IBM Plex Sans',
  fontEpoch: 1,
  sizeCssPx: 11,
  style: 'normal',
  weight: 400,
}

describe('gridTextLayoutV2', () => {
  test('right-aligns numeric text with metric-based vertical centering', () => {
    const layout = resolveCellTextLayoutV2({
      cell: { col: 1, row: 2 },
      cellWorldRect: { x: 104, y: 44, width: 104, height: 22 },
      color: '#111111',
      displayText: '1234',
      fontKey,
      text: '1234',
    })

    expect(layout.horizontalAlign).toBe('right')
    expect(layout.verticalAlign).toBe('middle')
    expect(layout.lines).toHaveLength(1)
    expect(layout.lines[0]?.worldX).toBeGreaterThan(104)
    expect(layout.lines[0]?.baselineWorldY).toBeCloseTo(57.2)
  })

  test('wraps by words and emits decoration rects', () => {
    const layout = resolveCellTextLayoutV2({
      cell: { col: 0, row: 0 },
      cellWorldRect: { x: 0, y: 0, width: 60, height: 50 },
      fontKey,
      metrics: createFallbackTextMetricsProvider(),
      text: 'alpha beta gamma',
      underline: true,
      wrap: true,
    })

    expect(layout.wrap).toBe(true)
    expect(layout.verticalAlign).toBe('top')
    expect(layout.lines.length).toBeGreaterThan(1)
    expect(layout.decorations.length).toBe(layout.lines.length)
  })

  test('packs glyph instances and queues missing atlas glyphs', () => {
    const layout = resolveCellTextLayoutV2({
      cell: { col: 0, row: 0 },
      cellWorldRect: { x: 0, y: 0, width: 104, height: 22 },
      fontKey,
      text: 'A',
    })
    const atlas = new GlyphAtlasV2()
    const first = buildTextGlyphInstanceBuffer({ atlas, layouts: [layout] })

    expect(first.glyphCount).toBe(1)
    expect(first.missingGlyphCount).toBe(1)
    expect(atlas.drainMissing()).toHaveLength(1)

    const firstGlyph = layout.lines[0]?.glyphs[0]
    if (!firstGlyph) {
      throw new Error('expected first glyph')
    }
    atlas.registerGlyph(firstGlyph.atlasGlyphKey, { height: 14, pageId: 0, width: 8, x: 10, y: 12 })
    const second = buildTextGlyphInstanceBuffer({ atlas, layouts: [layout] })

    expect(second.missingGlyphCount).toBe(0)
    expect(Array.from(second.floats.slice(4, 10))).toEqual([0, 1, 10, 12, 8, 14])
  })
})
