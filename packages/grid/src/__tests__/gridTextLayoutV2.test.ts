import { describe, expect, test } from 'vitest'
import { GlyphKeyRegistryV3 } from '../renderer-v3/glyph-key.js'
import { TextAtlasPagesV3 } from '../renderer-v3/text-atlas-pages.js'
import { resolveCellTextLayoutV2 } from '../text/gridTextLayoutV2.js'
import { createFallbackTextMetricsProvider } from '../text/gridTextMetrics.js'
import type { FontKey } from '../text/gridTextPacket.js'

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

  test('emits glyph placements that can be registered in V3 atlas pages', () => {
    const layout = resolveCellTextLayoutV2({
      cell: { col: 0, row: 0 },
      cellWorldRect: { x: 0, y: 0, width: 104, height: 22 },
      fontKey,
      text: 'A',
    })
    const firstGlyph = layout.lines[0]?.glyphs[0]
    if (!firstGlyph) {
      throw new Error('expected first glyph')
    }

    const registry = new GlyphKeyRegistryV3()
    const glyphId = registry.intern({
      dprBucket: fontKey.dprBucket,
      fontInternId: 1,
      glyph: firstGlyph.glyph,
    })
    const atlas = new TextAtlasPagesV3()
    atlas.upsertPage({ height: 64, pageId: 1, width: 64 })
    const record = atlas.registerGlyph({ glyphId, pageId: 1, u0: 0.1, u1: 0.2, v0: 0.3, v1: 0.4 })

    expect(firstGlyph.atlasGlyphKey).toContain(':A')
    expect(record.glyphId).toBe(glyphId)
    expect(atlas.resolveGlyph(glyphId)).toBe(record)
    expect(atlas.stats().dirtyPageCount).toBe(1)
  })
})
