import { describe, expect, it } from 'vitest'
import { DirtyMaskV3, DirtyTileIndexV3 } from '../renderer-v3/tile-damage-index.js'
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
})
