import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { ValueTag } from '../../packages/protocol/src/enums.js'
import { extractFormulaOracles, inspectWorkbookFootprint } from '../public-workbook-corpus-workbook.ts'

describe('public workbook corpus workbook helpers', () => {
  it('extracts formula oracles from broad sparse worksheet refs', () => {
    const oracles = extractFormulaOracles(buildBroadSparseWorkbookBytes())

    expect(oracles).toEqual([
      {
        sheetName: 'Sparse',
        address: 'XFD512',
        expected: { tag: ValueTag.Number, value: 42 },
      },
    ])
  }, 15_000)

  it('records explicit used ranges from actual populated cells instead of broad worksheet refs', () => {
    const footprint = inspectWorkbookFootprint(buildBroadSparseWorkbookBytes(), 'sparse.xlsx')

    expect(footprint.workbookMetadata.dimensions).toEqual([
      {
        sheetName: 'Sparse',
        rowCount: 512,
        columnCount: 16_384,
        nonEmptyCellCount: 1,
        usedRange: { startRow: 511, startColumn: 16_383, endRow: 511, endColumn: 16_383 },
      },
    ])
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
