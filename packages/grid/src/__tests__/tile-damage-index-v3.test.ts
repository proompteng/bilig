import { describe, expect, it } from 'vitest'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
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
    expect(index.getSpans(keys[0])).toEqual([
      {
        colEnd: 127,
        colStart: 127,
        mask: DirtyMaskV3.Value | DirtyMaskV3.Text,
        rowEnd: 31,
        rowStart: 31,
      },
    ])
    expect([...index.consumeVisible(keys)]).toEqual(keys)
    expect(index.getSpans(keys[0])).toEqual([])
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
    expect(index.getSpans(axisX)).toContainEqual({
      colEnd: 1,
      colStart: 0,
      mask: DirtyMaskV3.Rect | DirtyMaskV3.AxisX,
      rowEnd: 31,
      rowStart: 0,
    })
    expect(index.getSpans(axisY)).toContainEqual({
      colEnd: 127,
      colStart: 0,
      mask: DirtyMaskV3.Text | DirtyMaskV3.AxisY,
      rowEnd: 1,
      rowStart: 0,
    })
    expect(index.consumeVisible([axisX, axisY])).toEqual([axisX, axisY])
    expect(index.consumeVisible([axisX, axisY])).toEqual([])
    expect(index.getSpans(axisX)).toEqual([])
    expect(index.getSpans(axisY)).toEqual([])
  })

  it('expands true axis geometry damage through the rest of the tile', () => {
    const index = new DirtyTileIndexV3()
    markWorkbookDeltaDirtyTilesV3(
      index,
      {
        sheetOrdinal: 2,
        dirty: {
          axisX: Uint32Array.from([128, 128, DirtyMaskV3.AxisX | DirtyMaskV3.Rect | DirtyMaskV3.Text]),
          axisY: Uint32Array.from([32, 32, DirtyMaskV3.AxisY | DirtyMaskV3.Rect | DirtyMaskV3.Text]),
          cellRanges: new Uint32Array(),
        },
      },
      { dprBucket: 1 },
    )

    const axisX = tileKeyFromCell({ sheetOrdinal: 2, dprBucket: 1, row: 0, col: 128 })
    const axisY = tileKeyFromCell({ sheetOrdinal: 2, dprBucket: 1, row: 32, col: 0 })

    expect(index.getSpans(axisX)).toContainEqual({
      colEnd: 127,
      colStart: 0,
      mask: DirtyMaskV3.AxisX | DirtyMaskV3.Rect | DirtyMaskV3.Text,
      rowEnd: 31,
      rowStart: 0,
    })
    expect(index.getSpans(axisY)).toContainEqual({
      colEnd: 127,
      colStart: 0,
      mask: DirtyMaskV3.AxisY | DirtyMaskV3.Rect | DirtyMaskV3.Text,
      rowEnd: 31,
      rowStart: 0,
    })
  })

  it('tracks full-column and full-row cell damage as axis spans instead of enumerating sheet tiles', () => {
    const index = new DirtyTileIndexV3()
    const styleMask = DirtyMaskV3.Style | DirtyMaskV3.Rect | DirtyMaskV3.Border
    index.markCellRange({
      sheetOrdinal: 3,
      dprBucket: 1,
      rowStart: 0,
      rowEnd: MAX_ROWS - 1,
      colStart: 1,
      colEnd: 4,
      mask: styleMask,
    })

    const firstVisibleColumnTile = tileKeyFromCell({ sheetOrdinal: 3, dprBucket: 1, row: 0, col: 1 })
    const deepColumnTile = tileKeyFromCell({ sheetOrdinal: 3, dprBucket: 1, row: 160, col: 1 })
    const columnMask = styleMask | DirtyMaskV3.AxisX

    expect(index.getMask(firstVisibleColumnTile)).toBe(columnMask)
    expect(index.getMask(deepColumnTile)).toBe(columnMask)
    expect(index.getSpans(firstVisibleColumnTile)).toEqual([
      {
        colEnd: 4,
        colStart: 1,
        mask: columnMask,
        rowEnd: 31,
        rowStart: 0,
      },
    ])
    expect(index.getSpans(deepColumnTile)).toEqual([
      {
        colEnd: 4,
        colStart: 1,
        mask: columnMask,
        rowEnd: 31,
        rowStart: 0,
      },
    ])

    expect(index.consumeVisible([firstVisibleColumnTile])).toEqual([firstVisibleColumnTile])
    expect(index.getUnconsumedMask(firstVisibleColumnTile)).toBe(0)
    expect(index.getUnconsumedMask(deepColumnTile)).toBe(columnMask)

    index.markCellRange({
      sheetOrdinal: 3,
      dprBucket: 1,
      rowStart: 2,
      rowEnd: 5,
      colStart: 0,
      colEnd: MAX_COLS - 1,
      mask: DirtyMaskV3.Text,
    })

    const rowTile = tileKeyFromCell({ sheetOrdinal: 3, dprBucket: 1, row: 2, col: 0 })
    expect(index.getSpans(rowTile)).toContainEqual({
      colEnd: 127,
      colStart: 0,
      mask: DirtyMaskV3.Text | DirtyMaskV3.AxisY,
      rowEnd: 5,
      rowStart: 2,
    })
  })
})
