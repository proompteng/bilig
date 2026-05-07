import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { importXlsx } from '../index.js'

describe('XLSX sparse ranges', () => {
  it('imports actual cells without scanning every coordinate in a broad sparse ref', () => {
    const imported = importXlsx(buildBroadSparseWorkbookBytes(), 'broad-sparse.xlsx')
    const sheet = imported.snapshot.sheets[0]

    expect(sheet?.cells).toEqual([{ address: 'XFD512', formula: '40+2' }])
    expect(imported.preview.sheets[0]).toMatchObject({
      rowCount: 512,
      columnCount: 16_384,
      nonEmptyCellCount: 1,
    })
  }, 15_000)
})

function buildBroadSparseWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet: XLSX.WorkSheet = {
    XFD512: { t: 'n', f: '40+2', v: 42 },
    '!ref': 'A1:XFD512',
  }
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sparse')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}
