import { describe, expect, it } from 'vitest'
import type { CellNumberFormatRecord, CellStyleRecord } from '@bilig/protocol'
import {
  axisMetadataKey,
  cellNumberFormatIdForCode,
  cellStyleIdForKey,
  cellStyleKey,
  deleteRecordsBySheet,
  normalizeCellNumberFormatRecord,
  normalizeCellStyleRecord,
} from '../workbook-store-records.js'

describe('workbook store records', () => {
  it('normalizes style records into stable workbook records', () => {
    const style: CellStyleRecord = {
      id: ' style-a ',
      fill: { backgroundColor: ' #abc ' },
      font: {
        family: ' Arial ',
        color: '#123456',
        size: 240,
        bold: true,
        underline: true,
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrap: true,
        indent: 24.8,
      },
      borders: {
        top: { color: '#f0f', style: 'solid', weight: 'thin' },
        right: { color: '#112233', style: 'solid' },
      },
    }

    expect(normalizeCellStyleRecord(style)).toEqual({
      id: 'style-a',
      fill: { backgroundColor: '#aabbcc' },
      font: {
        family: 'Arial',
        color: '#123456',
        size: 144,
        bold: true,
        underline: true,
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrap: true,
        indent: 16,
      },
      borders: {
        top: { color: '#ff00ff', style: 'solid', weight: 'thin' },
      },
    })
  })

  it('rejects invalid style and format identifiers', () => {
    expect(() => normalizeCellStyleRecord({ id: '   ' })).toThrow('Cell style id must be non-empty')
    expect(() =>
      normalizeCellStyleRecord({
        id: 'style-a',
        fill: { backgroundColor: 'red' },
      }),
    ).toThrow('Unsupported background color: red')

    const format: CellNumberFormatRecord = { id: ' ', code: '0.00' }
    expect(() => normalizeCellNumberFormatRecord(format)).toThrow('Cell number format id must be non-empty')
  })

  it('builds deterministic record keys and deletes sheet-scoped entries', () => {
    const styleKey = cellStyleKey({
      id: 'style-a',
      font: { bold: true },
    })
    expect(cellStyleIdForKey(styleKey)).toBe(cellStyleIdForKey(styleKey))
    expect(cellNumberFormatIdForCode('$0.00')).toBe(cellNumberFormatIdForCode('$0.00'))
    expect(axisMetadataKey('Sheet1', 2, 3)).toBe('Sheet1:2:3')

    const bucket = new Map<string, { sheetName: string; value: string }>([
      ['a', { sheetName: 'Sheet1', value: 'keep?' }],
      ['b', { sheetName: 'Sheet2', value: 'keep' }],
      ['c', { sheetName: 'Sheet1', value: 'delete' }],
    ])

    deleteRecordsBySheet(bucket, 'Sheet1', (record) => record.sheetName)

    expect([...bucket.entries()]).toEqual([['b', { sheetName: 'Sheet2', value: 'keep' }]])
  })
})
