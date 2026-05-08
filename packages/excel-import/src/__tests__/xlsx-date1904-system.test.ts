import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'

import { exportXlsx, importXlsx } from '../index.js'

describe('1904 date system import', () => {
  it('preserves workbookPr date1904 and evaluates date formulas in the workbook date system', () => {
    const imported = importXlsx(buildDate1904WorkbookBytes(), 'date1904-finance-dates.xlsx')

    expect(imported.snapshot.workbook.metadata?.calculationSettings).toMatchObject({
      dateSystem: '1904',
    })

    const engine = new SpreadsheetEngine({ workbookName: 'date1904-import' })
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Date1904', 'B2')).toEqual({ tag: ValueTag.Number, value: 1904 })
    expect(engine.getCellValue('Date1904', 'C2')).toMatchObject({
      tag: ValueTag.String,
      value: '1904-01-02',
    })

    expect(workbookXml(exportXlsx(engine.exportSnapshot()))).toContain('date1904="1"')
  })
})

function buildDate1904WorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Serial', 'Year', 'Text'],
    [1, null, null],
  ])
  sheet.B2 = { t: 'n', f: 'YEAR(A2)', v: 1904 }
  sheet.C2 = { t: 's', f: 'TEXT(A2,"yyyy-mm-dd")', v: '1904-01-02' }
  sheet['!ref'] = 'A1:C2'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Date1904')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceWorkbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  zip['xl/workbook.xml'] = strToU8(
    /<workbookPr\b/u.test(sourceWorkbookXml)
      ? sourceWorkbookXml.replace(/<workbookPr\b([^>]*)\/>/u, '<workbookPr$1 date1904="1"/>')
      : sourceWorkbookXml.replace(/<sheets\b/u, '<workbookPr date1904="1"/><sheets'),
  )
  return zipSync(zip)
}

function workbookXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/workbook.xml'] ?? new Uint8Array())
}
