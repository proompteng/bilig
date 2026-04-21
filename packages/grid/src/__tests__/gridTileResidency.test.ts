import { describe, expect, test } from 'vitest'
import { resolveGridTileResidency } from '../gridTileResidency.js'

describe('gridTileResidency', () => {
  test('keeps the visible resident tile', () => {
    expect(
      resolveGridTileResidency({
        visibleViewport: { colStart: 7, colEnd: 18, rowStart: 10, rowEnd: 25 },
      }).visible,
    ).toEqual({
      colEnd: 127,
      colStart: 0,
      rowEnd: 31,
      rowStart: 0,
    })
  })

  test('prefetches directional warm neighbors', () => {
    const plan = resolveGridTileResidency({
      velocityX: 10,
      velocityY: 5,
      visibleViewport: { colStart: 128, colEnd: 150, rowStart: 32, rowEnd: 60 },
    })

    expect(plan.warm).toContainEqual({ colStart: 256, colEnd: 383, rowStart: 32, rowEnd: 63 })
    expect(plan.warm).toContainEqual({ colStart: 128, colEnd: 255, rowStart: 64, rowEnd: 95 })
    expect(plan.warm).toContainEqual({ colStart: 256, colEnd: 383, rowStart: 64, rowEnd: 95 })
  })
})
