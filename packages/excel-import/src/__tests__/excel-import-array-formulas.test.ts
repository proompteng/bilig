import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { SpreadsheetEngine } from '@bilig/core'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { importXlsx } from '../index.js'

function buildArrayFormulaWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    [1, 0, null, 1, null, null, null, null],
    [0, 1, null, 2, null, null, null, null],
  ])

  sheet.F1 = { t: 'n', f: 'MMULT(MINVERSE(A1:B2),D1:D2)', F: 'F1:F2', v: 1 }
  sheet.F2 = { t: 'n', F: 'F1:F2', v: 2 }
  sheet.H1 = { t: 'n', f: 'INDEX(MMULT(MINVERSE(A1:B2),D1:D2),1,1)' }
  sheet.H2 = { t: 'n', f: 'INDEX(MMULT(MINVERSE(A1:B2),D1:D2),2,1)' }
  sheet['!ref'] = 'A1:H2'

  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

describe('excel import array formulas', () => {
  it('imports legacy array-formula ranges so cached follower values do not block the leader', async () => {
    const imported = importXlsx(buildArrayFormulaWorkbook(), 'array-formula-import-evaluation.xlsx')

    expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'F1', rows: 2, cols: 1 }])

    const engine = new SpreadsheetEngine({ workbookName: 'issue-108-array-formula-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F1')).not.toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
  })
})
