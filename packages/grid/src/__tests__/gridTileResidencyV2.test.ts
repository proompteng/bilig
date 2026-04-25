import { describe, expect, test } from 'vitest'
import { resolveGridTileResidencyV2 } from '../gridTileResidencyV2.js'

describe('gridTileResidencyV2', () => {
  test('keeps the visible resident tile', () => {
    expect(
      resolveGridTileResidencyV2({
        visibleViewport: { colStart: 7, colEnd: 18, rowStart: 10, rowEnd: 25 },
      }).visible,
    ).toEqual({
      colEnd: 255,
      colStart: 0,
      rowEnd: 95,
      rowStart: 0,
    })
  })

  test('keeps all neighboring boundary tiles warm', () => {
    const plan = resolveGridTileResidencyV2({
      velocityX: 10,
      velocityY: 5,
      visibleViewport: { colStart: 256, colEnd: 280, rowStart: 96, rowEnd: 120 },
    })

    expect(plan.warm).toHaveLength(8)
    expect(plan.warm).toContainEqual({ colStart: 512, colEnd: 767, rowStart: 96, rowEnd: 191 })
    expect(plan.warm).toContainEqual({ colStart: 256, colEnd: 511, rowStart: 192, rowEnd: 287 })
    expect(plan.warm).toContainEqual({ colStart: 512, colEnd: 767, rowStart: 192, rowEnd: 287 })
    expect(plan.warm).toContainEqual({ colStart: 0, colEnd: 255, rowStart: 0, rowEnd: 95 })
  })
})
