import { describe, expect, it } from 'vitest'
import {
  columnToIndex,
  formatAddress,
  formatRangeAddress,
  indexToColumn,
  isCellReferenceText,
  isColumnReferenceText,
  isRowReferenceText,
  parseCellAddress,
  parseRangeAddress,
  toQualifiedAddress,
} from '../addressing.js'

describe('addressing helpers', () => {
  it('round-trips columns and detects reference text kinds', () => {
    expect(columnToIndex('A')).toBe(0)
    expect(columnToIndex('Z')).toBe(25)
    expect(columnToIndex('AA')).toBe(26)
    expect(indexToColumn(0)).toBe('A')
    expect(indexToColumn(25)).toBe('Z')
    expect(indexToColumn(26)).toBe('AA')
    expect(formatAddress(4, 27)).toBe('AB5')

    expect(isCellReferenceText('b12')).toBe(true)
    expect(isCellReferenceText('12')).toBe(false)
    expect(isColumnReferenceText('$ab')).toBe(true)
    expect(isColumnReferenceText('A1')).toBe(false)
    expect(isRowReferenceText('42')).toBe(true)
    expect(isRowReferenceText('A42')).toBe(false)
  })

  it('parses cells and ranges with quoted sheet names and normalized ordering', () => {
    expect(parseCellAddress("'My Sheet'!b12")).toEqual({
      sheetName: 'My Sheet',
      row: 11,
      col: 1,
      text: 'B12',
    })

    expect(parseRangeAddress('B3:A1', 'Sheet1')).toEqual({
      kind: 'cells',
      sheetName: 'Sheet1',
      start: { kind: 'cell', sheetName: 'Sheet1', row: 0, col: 0, text: 'A1' },
      end: { kind: 'cell', sheetName: 'Sheet1', row: 2, col: 1, text: 'B3' },
    })

    expect(parseRangeAddress('7:3', 'Sheet1')).toEqual({
      kind: 'rows',
      sheetName: 'Sheet1',
      start: { kind: 'row', sheetName: 'Sheet1', row: 2, text: '3' },
      end: { kind: 'row', sheetName: 'Sheet1', row: 6, text: '7' },
    })

    expect(parseRangeAddress('D:B', 'Sheet1')).toEqual({
      kind: 'cols',
      sheetName: 'Sheet1',
      start: { kind: 'col', sheetName: 'Sheet1', col: 1, text: 'B' },
      end: { kind: 'col', sheetName: 'Sheet1', col: 3, text: 'D' },
    })

    expect(formatRangeAddress(parseRangeAddress("'Ops Team'!D:B", 'Ignored'))).toBe("'Ops Team'!B:D")
    expect(toQualifiedAddress('Sheet1', 'b3')).toBe('Sheet1!B3')
  })

  it('rejects invalid cell and range addresses', () => {
    expect(() => parseCellAddress('A:A')).toThrow('Invalid reference: A:A')
    expect(() => parseRangeAddress('A1')).toThrow('Invalid range address: A1')
    expect(() => parseRangeAddress('A1:B')).toThrow('Range endpoints must use the same reference type')
    expect(() => parseRangeAddress('Sheet1!A1:Sheet2!B2')).toThrow('Range endpoints must target the same sheet')
    expect(() => parseRangeAddress('A0:stillbad')).toThrow('Invalid reference: A0')
  })
})
