import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import { DirtyMaskV3, DirtyTileIndexV3 } from '../renderer-v3/tile-damage-index.js'
import { tileKeyFromCell, unpackTileKey53, type TileKey53 } from '../renderer-v3/tile-key.js'

describe('dirty tile index fuzz', () => {
  it('should map generated cell ranges to exactly their touched fixed tiles', async () => {
    await runProperty({
      suite: 'grid/renderer-v3/tile-damage-index/range-to-tiles',
      arbitrary: dirtyRangeArbitrary,
      predicate: async ({ rowStart, rowEnd, colStart, colEnd, mask, sheetOrdinal, dprBucket }) => {
        const index = new DirtyTileIndexV3()
        index.markCellRange({ sheetOrdinal, dprBucket, rowStart, rowEnd, colStart, colEnd, mask })

        const expectedKeys = expectedTouchedKeys({ sheetOrdinal, dprBucket, rowStart, rowEnd, colStart, colEnd })

        expect(index.peekWarm(expectedKeys)).toEqual(expectedKeys)
        for (const key of expectedKeys) {
          expect(index.getMask(key) & mask).toBe(mask)
          for (const span of index.getSpans(key)) {
            expect(span.rowStart).toBeGreaterThanOrEqual(0)
            expect(span.rowEnd).toBeGreaterThanOrEqual(span.rowStart)
            expect(span.rowEnd).toBeLessThan(VIEWPORT_TILE_ROW_COUNT)
            expect(span.colStart).toBeGreaterThanOrEqual(0)
            expect(span.colEnd).toBeGreaterThanOrEqual(span.colStart)
            expect(span.colEnd).toBeLessThan(VIEWPORT_TILE_COLUMN_COUNT)
          }
        }

        expect(index.consumeVisible(expectedKeys)).toEqual(expectedKeys)
        expect(index.peekWarm(expectedKeys)).toEqual([])
      },
      parameters: { numRuns: 120 },
    })
  })
})

const dirtyMaskArbitrary = fc.integer({
  min: DirtyMaskV3.Value,
  max: DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border,
})

const dirtyRangeArbitrary = fc
  .record({
    sheetOrdinal: fc.integer({ min: 0, max: 16 }),
    dprBucket: fc.integer({ min: 0, max: 4 }),
    rowStart: fc.integer({ min: 0, max: VIEWPORT_TILE_ROW_COUNT * 4 - 1 }),
    rowSpan: fc.integer({ min: 0, max: VIEWPORT_TILE_ROW_COUNT * 2 }),
    colStart: fc.integer({ min: 0, max: VIEWPORT_TILE_COLUMN_COUNT * 4 - 1 }),
    colSpan: fc.integer({ min: 0, max: VIEWPORT_TILE_COLUMN_COUNT * 2 }),
    mask: dirtyMaskArbitrary,
  })
  .map(({ sheetOrdinal, dprBucket, rowStart, rowSpan, colStart, colSpan, mask }) => ({
    sheetOrdinal,
    dprBucket,
    rowStart,
    rowEnd: rowStart + rowSpan,
    colStart,
    colEnd: colStart + colSpan,
    mask,
  }))

function expectedTouchedKeys(input: {
  readonly sheetOrdinal: number
  readonly dprBucket: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}): TileKey53[] {
  const keys = new Map<string, TileKey53>()
  for (const row of [input.rowStart, input.rowEnd]) {
    for (const col of [input.colStart, input.colEnd]) {
      const key = tileKeyFromCell({ sheetOrdinal: input.sheetOrdinal, dprBucket: input.dprBucket, row, col })
      const fields = unpackTileKey53(key)
      keys.set(`${fields.rowTile}:${fields.colTile}`, key)
    }
  }
  const rowTileStart = Math.floor(input.rowStart / VIEWPORT_TILE_ROW_COUNT)
  const rowTileEnd = Math.floor(input.rowEnd / VIEWPORT_TILE_ROW_COUNT)
  const colTileStart = Math.floor(input.colStart / VIEWPORT_TILE_COLUMN_COUNT)
  const colTileEnd = Math.floor(input.colEnd / VIEWPORT_TILE_COLUMN_COUNT)
  for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
    for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
      const row = rowTile * VIEWPORT_TILE_ROW_COUNT
      const col = colTile * VIEWPORT_TILE_COLUMN_COUNT
      const key = tileKeyFromCell({ sheetOrdinal: input.sheetOrdinal, dprBucket: input.dprBucket, row, col })
      keys.set(`${rowTile}:${colTile}`, key)
    }
  }
  return [...keys.values()]
}
