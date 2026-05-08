import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('worksheet sheetPr properties roundtrip', () => {
  it('preserves non-tabColor sheetPr codeName, outlinePr, and pageSetUpPr metadata', () => {
    const imported = importXlsx(buildWorksheetPropertiesWorkbookBytes(), 'worksheet-properties.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.sheetPr).toEqual({
      xml: '<sheetPr codeName="Sheet8"><outlinePr summaryBelow="0" summaryRight="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    })
    expect(imported.snapshot.sheets[0]?.metadata?.tabColor).toEqual({ rgb: 'FFFF0000' })
    expect(imported.snapshot.sheets[1]?.metadata?.sheetPr).toEqual({
      xml: '<sheetPr codeName="Sheet2"><pageSetUpPr/></sheetPr>',
    })

    const exportedZip = unzipSync(exportXlsx(imported.snapshot))
    const firstSheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const secondSheetXml = strFromU8(exportedZip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())

    expect(firstSheetXml).toContain(
      '<sheetPr codeName="Sheet8"><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0" summaryRight="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    )
    expect(secondSheetXml).toContain('<sheetPr codeName="Sheet2"><pageSetUpPr/></sheetPr>')
  })
})

function buildWorksheetPropertiesWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['summary']]), 'Summary')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['detail']]), 'Detail')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  replaceSheetPr(
    zip,
    1,
    '<sheetPr codeName="Sheet8"><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0" summaryRight="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
  )
  replaceSheetPr(zip, 2, '<sheetPr codeName="Sheet2"><pageSetUpPr/></sheetPr>')
  return zipSync(zip)
}

function replaceSheetPr(zip: Record<string, Uint8Array>, sheetIndex: number, sheetPrXml: string): void {
  const sheetPath = `xl/worksheets/sheet${String(sheetIndex)}.xml`
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1>${sheetPrXml}`))
}
