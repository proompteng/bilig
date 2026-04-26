import { describe, expect, it } from 'vitest'
import { DirtyMaskV3, DirtyTileIndexV3, markWorkbookDeltaDirtyTilesV3 } from '../renderer-v3/tile-damage-index.js'
import { tileKeyFromCell } from '../renderer-v3/tile-key.js'

describe('DirtyTileIndexV3', () => {
  it('maps cell range damage to touched fixed tiles', () => {
    const index = new DirtyTileIndexV3()
    index.markCellRange({
      sheetOrdinal: 1,
      dprBucket: 1,
      rowStart: 31,
      rowEnd: 32,
      colStart: 127,
      colEnd: 128,
      mask: DirtyMaskV3.Value | DirtyMaskV3.Text,
    })

    const keys = [
      tileKeyFromCell({ sheetOrdinal: 1, dprBucket: 1, row: 31, col: 127 }),
      tileKeyFromCell({ sheetOrdinal: 1, dprBucket: 1, row: 31, col: 128 }),
      tileKeyFromCell({ sheetOrdinal: 1, dprBucket: 1, row: 32, col: 127 }),
      tileKeyFromCell({ sheetOrdinal: 1, dprBucket: 1, row: 32, col: 128 }),
    ]

    expect([...index.peekWarm(keys)]).toEqual(keys)
    expect(index.getMask(keys[0])).toBe(DirtyMaskV3.Value | DirtyMaskV3.Text)
    expect([...index.consumeVisible(keys)]).toEqual(keys)
    expect([...index.peekWarm(keys)]).toEqual([])
  })

  it('marks dirty tiles from workbook delta batches', () => {
    const index = new DirtyTileIndexV3()
    markWorkbookDeltaDirtyTilesV3(
      index,
      {
        sheetOrdinal: 2,
        dirty: {
          axisX: Uint32Array.from([128, 129, DirtyMaskV3.Rect]),
          axisY: Uint32Array.from([32, 33, DirtyMaskV3.Text]),
          cellRanges: Uint32Array.from([0, 0, 0, 0, DirtyMaskV3.Value]),
        },
      },
      { dprBucket: 1 },
    )

    const origin = tileKeyFromCell({ sheetOrdinal: 2, dprBucket: 1, row: 0, col: 0 })
    const axisX = tileKeyFromCell({ sheetOrdinal: 2, dprBucket: 1, row: 0, col: 128 })
    const axisY = tileKeyFromCell({ sheetOrdinal: 2, dprBucket: 1, row: 32, col: 0 })

    expect(index.getMask(origin) & DirtyMaskV3.Value).toBe(DirtyMaskV3.Value)
    expect(index.getMask(axisX) & DirtyMaskV3.AxisX).toBe(DirtyMaskV3.AxisX)
    expect(index.getMask(axisY) & DirtyMaskV3.AxisY).toBe(DirtyMaskV3.AxisY)
    expect(index.consumeVisible([axisX, axisY])).toEqual([axisX, axisY])
    expect(index.consumeVisible([axisX, axisY])).toEqual([])
  })
})
