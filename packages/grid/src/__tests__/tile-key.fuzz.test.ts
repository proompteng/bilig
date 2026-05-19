import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import {
  MAX_TILE_COLUMN_INDEX,
  MAX_TILE_ROW_INDEX,
  debugTileKey,
  packTileKey53,
  tileKeyFromCell,
  tileKeysForViewport,
  unpackTileKey53,
  type TileKeyFields,
} from '../renderer-v3/tile-key.js'

describe('renderer v3 tile key fuzz', () => {
  it('should roundtrip generated tile coordinates through safe packed keys', async () => {
    await runProperty({
      suite: 'grid/renderer-v3/tile-key/pack-unpack-roundtrip',
      arbitrary: tileKeyFieldsArbitrary,
      predicate: async (fields) => {
        const key = packTileKey53(fields)

        expect(Number.isSafeInteger(key)).toBe(true)
        expect(unpackTileKey53(key)).toEqual(fields)
        expect(debugTileKey(key)).toBe(`s${fields.sheetOrdinal}:r${fields.rowTile}:c${fields.colTile}:d${fields.dprBucket}`)
      },
      parameters: { numRuns: 160 },
    })
  })

  it('should enumerate the same tile keys as a simple viewport reference model', async () => {
    await runProperty({
      suite: 'grid/renderer-v3/tile-key/viewport-enumeration',
      arbitrary: viewportKeyInputArbitrary,
      predicate: async ({ sheetOrdinal, dprBucket, rowStart, rowEnd, colStart, colEnd }) => {
        const viewport = { rowStart, rowEnd, colStart, colEnd }
        const actual = tileKeysForViewport({ sheetOrdinal, viewport, dprBucket }).map(unpackTileKey53)
        const expected = expectedViewportTileFields({ sheetOrdinal, dprBucket, rowStart, rowEnd, colStart, colEnd })

        expect(actual).toEqual(expected)
        for (let row = rowStart; row <= rowEnd; row += Math.max(1, Math.floor((rowEnd - rowStart + 1) / 3))) {
          for (let col = colStart; col <= colEnd; col += Math.max(1, Math.floor((colEnd - colStart + 1) / 3))) {
            const cellKey = tileKeyFromCell({ sheetOrdinal, row, col, dprBucket })
            expect(tileKeysForViewport({ sheetOrdinal, viewport, dprBucket })).toContain(cellKey)
          }
        }
      },
      parameters: { numRuns: 120 },
    })
  })
})

const tileKeyFieldsArbitrary: fc.Arbitrary<TileKeyFields> = fc.record({
  sheetOrdinal: fc.integer({ min: 0, max: 4_096 }),
  rowTile: fc.integer({ min: 0, max: MAX_TILE_ROW_INDEX }),
  colTile: fc.integer({ min: 0, max: MAX_TILE_COLUMN_INDEX }),
  dprBucket: fc.integer({ min: 0, max: 15 }),
})

const viewportKeyInputArbitrary = fc
  .record({
    sheetOrdinal: fc.integer({ min: 0, max: 256 }),
    dprBucket: fc.integer({ min: 0, max: 8 }),
    rowStart: fc.integer({ min: 0, max: 256 }),
    rowSpan: fc.integer({ min: 0, max: VIEWPORT_TILE_ROW_COUNT * 3 }),
    colStart: fc.integer({ min: 0, max: 384 }),
    colSpan: fc.integer({ min: 0, max: VIEWPORT_TILE_COLUMN_COUNT * 3 }),
  })
  .map(({ sheetOrdinal, dprBucket, rowStart, rowSpan, colStart, colSpan }) => ({
    sheetOrdinal,
    dprBucket,
    rowStart,
    rowEnd: Math.min(MAX_ROWS - 1, rowStart + rowSpan),
    colStart,
    colEnd: Math.min(MAX_COLS - 1, colStart + colSpan),
  }))

function expectedViewportTileFields(input: {
  readonly sheetOrdinal: number
  readonly dprBucket: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}): TileKeyFields[] {
  const fields: TileKeyFields[] = []
  const rowTileStart = Math.floor(input.rowStart / VIEWPORT_TILE_ROW_COUNT)
  const rowTileEnd = Math.floor(input.rowEnd / VIEWPORT_TILE_ROW_COUNT)
  const colTileStart = Math.floor(input.colStart / VIEWPORT_TILE_COLUMN_COUNT)
  const colTileEnd = Math.floor(input.colEnd / VIEWPORT_TILE_COLUMN_COUNT)
  for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
    for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
      fields.push({ sheetOrdinal: input.sheetOrdinal, rowTile, colTile, dprBucket: input.dprBucket })
    }
  }
  return fields
}
