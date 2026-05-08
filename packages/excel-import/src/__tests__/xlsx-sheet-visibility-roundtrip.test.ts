import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('sheet visibility roundtrip', () => {
  it('preserves hidden and very hidden worksheet state', () => {
    const imported = importXlsx(buildSheetVisibilityWorkbookBytes(), 'sheet-visibility.xlsx')

    expect(imported.snapshot.sheets.map((sheet) => ({ name: sheet.name, visibility: sheet.metadata?.visibility }))).toEqual([
      { name: 'Inputs', visibility: undefined },
      { name: 'Support', visibility: 'hidden' },
      { name: 'Audit', visibility: 'veryHidden' },
    ])

    const exported = XLSX.read(exportXlsx(imported.snapshot), { type: 'array' })
    expect(exported.Workbook?.Sheets?.map((sheet) => sheet.Hidden ?? 0)).toEqual([0, 1, 2])
  })
})

function buildSheetVisibilityWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['visible']]), 'Inputs')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['hidden']]), 'Support')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['very hidden']]), 'Audit')
  workbook.Workbook = {
    ...workbook.Workbook,
    Sheets: [
      { name: 'Inputs', Hidden: 0 },
      { name: 'Support', Hidden: 1 },
      { name: 'Audit', Hidden: 2 },
    ],
  }
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}
