import { describe, expect, it } from 'vitest'
import { findWorkbookRangeRecord, overlayWorkbookRangeRecords } from '../workbook-range-records.js'

type TestRangeRecord = {
  range: {
    sheetName: string
    startAddress: string
    endAddress: string
  }
  id: string
}

describe('workbook range records', () => {
  it('splits overlapping records and keeps the external range shape plain', () => {
    const records = overlayWorkbookRangeRecords<TestRangeRecord>(
      [
        {
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C3' },
          id: 'style-a',
        },
      ],
      {
        range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' },
        id: 'style-b',
      },
      (range, record) => ({ range, id: record.id }),
      () => false,
    )

    expect(records).toEqual([
      { range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C1' }, id: 'style-a' },
      { range: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'C3' }, id: 'style-a' },
      { range: { sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'A2' }, id: 'style-a' },
      { range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C2' }, id: 'style-a' },
      { range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }, id: 'style-b' },
    ])
    expect(records[0]?.range).not.toHaveProperty('startRow')
    expect(records[0]?.range).not.toHaveProperty('endCol')
  })

  it('finds the most recently applied matching record', () => {
    const record = findWorkbookRangeRecord<TestRangeRecord>(
      [
        {
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C3' },
          id: 'style-a',
        },
        {
          range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
          id: 'style-b',
        },
      ],
      1,
      1,
    )

    expect(record).toEqual({
      range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
      id: 'style-b',
    })
  })
})
