import { describe, expect, test } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { getGridMetrics } from '../gridMetrics.js'
import { applyHiddenAxisSizes, resolveGridScrollSpacerSize } from '../gridScrollSurface.js'

describe('gridScrollSurface', () => {
  test('sizes the native scroll spacer from scrollable body extent', () => {
    const gridMetrics = getGridMetrics()
    const columnAxis = createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 })
    const rowAxis = createGridAxisWorldIndex({ axisLength: 30, defaultSize: 20 })

    expect(
      resolveGridScrollSpacerSize({
        columnAxis,
        rowAxis,
        frozenColumnWidth: 200,
        frozenRowHeight: 20,
        hostWidth: 520,
        hostHeight: 220,
        gridMetrics,
      }),
    ).toEqual({
      width: 2046,
      height: 624,
    })
  })

  test('collapses hidden axis entries before legacy scene paths consume sizes', () => {
    expect(applyHiddenAxisSizes({ 1: 132, 3: 96 }, { 3: true, 5: true })).toEqual({
      1: 132,
      3: 0,
      5: 0,
    })
  })
})
