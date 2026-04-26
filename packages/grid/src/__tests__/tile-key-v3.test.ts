import { describe, expect, it } from 'vitest'
import { VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import { debugTileKey, packTileKey53, tileKeyFromCell, tileKeysForViewport, unpackTileKey53 } from '../renderer-v3/tile-key.js'

describe('renderer-v3 tile keys', () => {
  it('packs fixed content tile coordinates into a reversible safe integer', () => {
    const key = packTileKey53({ sheetOrdinal: 17, rowTile: 1024, colTile: 121, dprBucket: 2 })

    expect(Number.isSafeInteger(key)).toBe(true)
    expect(unpackTileKey53(key)).toEqual({ sheetOrdinal: 17, rowTile: 1024, colTile: 121, dprBucket: 2 })
    expect(debugTileKey(key)).toBe('s17:r1024:c121:d2')
  })

  it('derives tile keys from cells and viewports using protocol tile dimensions', () => {
    expect(
      unpackTileKey53(tileKeyFromCell({ sheetOrdinal: 2, row: VIEWPORT_TILE_ROW_COUNT, col: VIEWPORT_TILE_COLUMN_COUNT, dprBucket: 1 })),
    ).toEqual({
      sheetOrdinal: 2,
      rowTile: 1,
      colTile: 1,
      dprBucket: 1,
    })

    expect(
      [...tileKeysForViewport({ sheetOrdinal: 2, dprBucket: 1, viewport: { rowStart: 0, rowEnd: 33, colStart: 0, colEnd: 129 } })].map(
        unpackTileKey53,
      ),
    ).toEqual([
      { sheetOrdinal: 2, rowTile: 0, colTile: 0, dprBucket: 1 },
      { sheetOrdinal: 2, rowTile: 0, colTile: 1, dprBucket: 1 },
      { sheetOrdinal: 2, rowTile: 1, colTile: 0, dprBucket: 1 },
      { sheetOrdinal: 2, rowTile: 1, colTile: 1, dprBucket: 1 },
    ])
  })

  it('rejects coordinates that do not fit the 53-bit layout', () => {
    expect(() => packTileKey53({ sheetOrdinal: -1, rowTile: 0, colTile: 0, dprBucket: 1 })).toThrow(RangeError)
    expect(() => packTileKey53({ sheetOrdinal: 0, rowTile: 0, colTile: 0, dprBucket: 16 })).toThrow(RangeError)
  })
})
