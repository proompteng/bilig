import { describe, expect, it } from 'vitest'
import { parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellSnapshot, type CellValue } from '@bilig/protocol'
import { isSelectionAggregateSummary } from '@bilig/grid'
import { collectWorkbookSelectionSummary } from '../workbook-selection-summary.js'

function summary(values: Readonly<Record<string, CellValue>>, startAddress = 'A1', endAddress = 'A1048576') {
  return collectWorkbookSelectionSummary(
    {
      workbook: {
        getSheet: () => ({
          grid: {
            forEachCellEntry: (listener) =>
              Object.keys(values).forEach((address, index) => {
                const cell = parseCellAddress(address)
                listener(index, cell.row, cell.col)
              }),
          },
        }),
      },
      getCell: (sheetName, address) => ({ sheetName, address, value: values[address] ?? { tag: ValueTag.Empty }, flags: 0, version: 1 }),
    },
    { sheetName: 'Sheet1', startAddress, endAddress },
  )
}

describe('workbook selection summaries', () => {
  it('includes offscreen cells, reversed ranges and calculated numeric values', () => {
    expect(
      summary(
        {
          A1: { tag: ValueTag.Number, value: 0.0004 },
          A5000: { tag: ValueTag.Number, value: 0.0006 },
          A10000: { tag: ValueTag.Number, value: 0.001 },
          B1: { tag: ValueTag.Number, value: 99 },
        },
        'A1048576',
        'A1',
      ),
    ).toEqual({ nonEmptyCount: 3, numericCount: 3, sum: 0.002, min: 0.0004, max: 0.001 })
  })

  it('counts text, booleans and empty-string formula results while excluding blank cells', () => {
    expect(
      summary({
        A1: { tag: ValueTag.String, value: 'hello', stringId: 1 },
        A2: { tag: ValueTag.Boolean, value: true },
        A3: { tag: ValueTag.String, value: '', stringId: 2 },
        A4: { tag: ValueTag.Empty },
      }),
    ).toEqual({ nonEmptyCount: 3, numericCount: 0, sum: 0, min: null, max: null })
  })

  it('preserves low-order contributions when large positive and negative values cancel', () => {
    expect(
      summary({
        A1: { tag: ValueTag.Number, value: 1e16 },
        A2: { tag: ValueTag.Number, value: 1 },
        A3: { tag: ValueTag.Number, value: -1e16 },
      }).sum,
    ).toBe(1)
  })

  it('handles 200,000 stored cells without spreading an array onto the call stack', () => {
    let visited = 0
    const result = collectWorkbookSelectionSummary(
      {
        workbook: {
          getSheet: () => ({
            grid: {
              forEachCellEntry: (listener) => {
                for (let row = 0; row < 200_000; row += 1) listener(row, row, 0)
              },
            },
          }),
        },
        getCell: (sheetName, address): CellSnapshot => {
          visited += 1
          return { sheetName, address, value: { tag: ValueTag.Number, value: 1 }, flags: 0, version: 1 }
        },
      },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'XFD1048576' },
    )
    expect(visited).toBe(200_000)
    expect(result).toEqual({ nonEmptyCount: 200_000, numericCount: 200_000, sum: 200_000, min: 1, max: 1 })
  })

  it('rejects malformed IPC results and inconsistent extrema', () => {
    expect(isSelectionAggregateSummary(summary({}))).toBe(true)
    expect(isSelectionAggregateSummary({ nonEmptyCount: 1, numericCount: 2, sum: 3, min: 1, max: 2 })).toBe(false)
    expect(isSelectionAggregateSummary({ nonEmptyCount: 1, numericCount: 1, sum: NaN, min: 1, max: 1 })).toBe(false)
    expect(isSelectionAggregateSummary({ nonEmptyCount: 1, numericCount: 1, sum: 1, min: 2, max: 1 })).toBe(false)
  })
})
