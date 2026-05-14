import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellSnapshot } from '@bilig/protocol'
import { cellToCsvValue, parseCsv, parseCsvCellInput, serializeCsv } from '../csv.js'

function mockCell(overrides: Partial<CellSnapshot>): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address: 'A1',
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
    ...overrides,
  }
}

describe('csv helpers', () => {
  it('serializes cell snapshots into CSV scalar text', () => {
    expect(
      cellToCsvValue(
        mockCell({
          address: 'A1',
          formula: 'A2+1',
          value: { tag: ValueTag.Number, value: 0 },
        }),
      ),
    ).toBe('=A2+1')
    expect(cellToCsvValue(mockCell({ address: 'A1', value: { tag: ValueTag.Empty } }))).toBe('')
    expect(
      cellToCsvValue(
        mockCell({
          address: 'A1',
          value: { tag: ValueTag.Number, value: 42.5 },
        }),
      ),
    ).toBe('42.5')
    expect(
      cellToCsvValue(
        mockCell({
          address: 'A1',
          value: { tag: ValueTag.Boolean, value: true },
        }),
      ),
    ).toBe('TRUE')
    expect(
      cellToCsvValue(
        mockCell({
          address: 'A1',
          value: { tag: ValueTag.String, value: 'hello', stringId: 0 },
        }),
      ),
    ).toBe('hello')
    expect(
      cellToCsvValue(
        mockCell({
          address: 'A1',
          value: { tag: ValueTag.Error, code: ErrorCode.Ref },
        }),
      ),
    ).toBe('#Ref')
  })

  it('serializes and parses quoted CSV content, CRLF rows, and trailing empties', () => {
    expect(
      serializeCsv([
        ['plain', 'two,parts', 'quote"inside', 'semi;inside', 'tab\tinside'],
        ['line\nbreak', ''],
      ]),
    ).toBe('plain,"two,parts","quote""inside","semi;inside","tab\tinside"\n"line\nbreak",')

    expect(parseCsv('"two,parts","quote""inside"\r\nlast,')).toEqual([
      ['two,parts', 'quote"inside'],
      ['last', ''],
    ])
    expect(parseCsv(serializeCsv([['text:;']]))).toEqual([['text:;']])
    expect(parseCsv('')).toEqual([])
    expect(parseCsv('solo')).toEqual([['solo']])
  })

  it('auto-detects semicolon-delimited CSV rows', () => {
    expect(parseCsv('Account;Amount;Tax\n4000;125,50;20,08')).toEqual([
      ['Account', 'Amount', 'Tax'],
      ['4000', '125,50', '20,08'],
    ])
    expect(parseCsv('Name;Description;Amount\nFoo;"contains;semicolon";125,50')).toEqual([
      ['Name', 'Description', 'Amount'],
      ['Foo', 'contains;semicolon', '125,50'],
    ])
  })

  it('parses CSV cell inputs into formulas, booleans, numbers, raw strings, or empties', () => {
    expect(parseCsvCellInput('   ')).toBeUndefined()
    expect(parseCsvCellInput('=SUM(A1:A2)')).toEqual({ formula: 'SUM(A1:A2)' })
    expect(parseCsvCellInput('TRUE')).toEqual({ value: true })
    expect(parseCsvCellInput('FALSE')).toEqual({ value: false })
    expect(parseCsvCellInput(' -12.5 ')).toEqual({ value: -12.5 })
    expect(parseCsvCellInput('001')).toEqual({ value: '001' })
    expect(parseCsvCellInput('0')).toEqual({ value: 0 })
    expect(parseCsvCellInput('hello')).toEqual({ value: 'hello' })
  })

  it('parses decimal-comma CSV cell inputs when requested', () => {
    expect(parseCsvCellInput('125,50', { decimalSeparator: ',' })).toEqual({ value: 125.5 })
    expect(parseCsvCellInput('-12,25', { decimalSeparator: ',' })).toEqual({ value: -12.25 })
    expect(parseCsvCellInput('0,00', { decimalSeparator: ',' })).toEqual({ value: 0 })
    expect(parseCsvCellInput('1.234', { decimalSeparator: ',' })).toEqual({ value: 1234 })
    expect(parseCsvCellInput('1.234,56', { decimalSeparator: ',' })).toEqual({ value: 1234.56 })
  })

  it('parses common accounting number formats', () => {
    expect(parseCsvCellInput('$1,234.56')).toEqual({ value: 1234.56 })
    expect(parseCsvCellInput('12.5%')).toEqual({ value: 0.125 })
    expect(parseCsvCellInput('-3.25%')).toEqual({ value: -0.0325 })
    expect(parseCsvCellInput('($987.65)')).toEqual({ value: -987.65 })
    expect(parseCsvCellInput('(987.65)')).toEqual({ value: -987.65 })
  })
})
