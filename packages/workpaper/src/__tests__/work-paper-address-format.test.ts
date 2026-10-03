import { describe, expect, it } from 'vitest'
import {
  formatQualifiedCellAddress,
  formatTrackedA1,
  formatWorkPaperCellAddressText,
  formatWorkPaperCellRangeText,
  parseWorkPaperCellAddressText,
  parseWorkPaperCellRangeText,
  quoteSheetNameIfNeeded,
  resolveDefaultWorkPaperSheetName,
  sourceRangeRef,
} from '../work-paper-address-format.js'

const requireTestSheetId = (sheetName: string): number => (sheetName === 'Data' ? 7 : 9)
const testSheetName = (sheetId: number): string => (sheetId === 1 ? 'Data' : 'Q1 Plan')

describe('work paper address format helpers', () => {
  it('quotes sheet names only when required', () => {
    expect(quoteSheetNameIfNeeded('Data_1.$')).toBe('Data_1.$')
    expect(quoteSheetNameIfNeeded('Q1 Plan')).toBe("'Q1 Plan'")
    expect(quoteSheetNameIfNeeded("Bob's Plan")).toBe("'Bob''s Plan'")
  })

  it('formats qualified and unqualified cell addresses', () => {
    expect(formatQualifiedCellAddress(undefined, 0, 0)).toBe('A1')
    expect(formatQualifiedCellAddress('Sheet1', 2, 27)).toBe('Sheet1!AB3')
    expect(formatQualifiedCellAddress('Q1 Plan', 1, 1)).toBe("'Q1 Plan'!B2")
  })

  it('formats cached and uncached A1 addresses consistently', () => {
    expect(formatTrackedA1(0, 0)).toBe('A1')
    expect(formatTrackedA1(0, 0)).toBe('A1')
    expect(formatTrackedA1(8192, 70)).toBe('BS8193')
  })

  it('builds range refs for engine calls', () => {
    expect(
      sourceRangeRef('Data', {
        start: { sheet: 1, row: 0, col: 1 },
        end: { sheet: 1, row: 2, col: 3 },
      }),
    ).toEqual({
      sheetName: 'Data',
      startAddress: 'B1',
      endAddress: 'D3',
    })
  })

  it('resolves default sheet names for parser calls', () => {
    expect(
      resolveDefaultWorkPaperSheetName({
        defaultSheetId: 2,
        sheets: [{ name: 'Only' }],
        sheetName: (sheetId) => `Sheet${sheetId}`,
      }),
    ).toBe('Sheet2')
    expect(
      resolveDefaultWorkPaperSheetName({
        sheets: [{ name: 'Only' }],
        sheetName: (sheetId) => `Sheet${sheetId}`,
      }),
    ).toBe('Only')
    expect(
      resolveDefaultWorkPaperSheetName({
        sheets: [{ name: 'One' }, { name: 'Two' }],
        sheetName: (sheetId) => `Sheet${sheetId}`,
      }),
    ).toBeUndefined()
  })

  it('parses cell and range text with optional default sheets', () => {
    expect(parseWorkPaperCellAddressText({ value: 'B2', defaultSheetName: 'Data', requireSheetId: requireTestSheetId })).toEqual({
      sheet: 7,
      row: 1,
      col: 1,
    })
    expect(parseWorkPaperCellAddressText({ value: 'B2', requireSheetId: requireTestSheetId })).toBeUndefined()
    expect(parseWorkPaperCellRangeText({ value: 'Data!A1:B3', requireSheetId: requireTestSheetId })).toEqual({
      start: { sheet: 7, row: 0, col: 0 },
      end: { sheet: 7, row: 2, col: 1 },
    })
    expect(parseWorkPaperCellRangeText({ value: 'A:A', defaultSheetName: 'Data', requireSheetId: requireTestSheetId })).toBeUndefined()
  })

  it('formats public cell and range text with context sheet options', () => {
    const address = { sheet: 2, row: 4, col: 3 }
    const range = {
      start: { sheet: 2, row: 0, col: 0 },
      end: { sheet: 2, row: 1, col: 1 },
    }

    expect(formatWorkPaperCellAddressText({ address, optionsOrContextSheetId: 2, sheetName: testSheetName })).toBe('D5')
    expect(formatWorkPaperCellAddressText({ address, optionsOrContextSheetId: 1, sheetName: testSheetName })).toBe("'Q1 Plan'!D5")
    expect(formatWorkPaperCellRangeText({ range, optionsOrContextSheetId: { includeSheetName: true }, sheetName: testSheetName })).toBe(
      "'Q1 Plan'!A1:B2",
    )
  })
})
