import { describe, expect, test } from 'vitest'
import { resolveGridRenderLayerIndex } from '../renderer-v2/grid-render-layers.js'

describe('grid render layers', () => {
  test('keeps text above selection fill and below active cell border', () => {
    expect(resolveGridRenderLayerIndex('selection-fill')).toBeLessThan(resolveGridRenderLayerIndex('body-text'))
    expect(resolveGridRenderLayerIndex('body-text')).toBeLessThan(resolveGridRenderLayerIndex('active-cell-border'))
  })
})
