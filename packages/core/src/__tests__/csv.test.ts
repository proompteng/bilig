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
        ['plain', 'two,parts', 'quote"inside'],
        ['line\nbreak', ''],
      ]),
    ).toBe('plain,"two,parts","quote""inside"\n"line\nbreak",')

    expect(parseCsv('"two,parts","quote""inside"\r\nlast,')).toEqual([
      ['two,parts', 'quote"inside'],
      ['last', ''],
    ])
    expect(parseCsv('')).toEqual([])
    expect(parseCsv('solo')).toEqual([['solo']])
  })

  it('parses CSV cell inputs into formulas, booleans, numbers, raw strings, or empties', () => {
    expect(parseCsvCellInput('   ')).toBeUndefined()
    expect(parseCsvCellInput('=SUM(A1:A2)')).toEqual({ formula: 'SUM(A1:A2)' })
    expect(parseCsvCellInput('TRUE')).toEqual({ value: true })
    expect(parseCsvCellInput('FALSE')).toEqual({ value: false })
    expect(parseCsvCellInput(' -12.5 ')).toEqual({ value: -12.5 })
    expect(parseCsvCellInput('001')).toEqual({ value: 1 })
    expect(parseCsvCellInput('hello')).toEqual({ value: 'hello' })
  })
})
