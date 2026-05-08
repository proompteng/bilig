import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { SpreadsheetEngine } from '@bilig/core'

import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX sheet protection attribute roundtrip', () => {
  it('preserves worksheet sheetProtection attributes without normalizing them to sheet=1', () => {
    const imported = importXlsx(buildSheetProtectionWorkbookBytes(), 'sheet-protection-options.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.sheetProtection).toEqual({
      sheetName: 'Protected',
      xmlAttributes: [
        { name: 'selectLockedCells', value: '1' },
        { name: 'algorithmName', value: 'workpaper&excel' },
      ],
    })

    const exportedSheetXml = sheetXml(exportXlsx(imported.snapshot))
    expect(exportedSheetXml).toContain('<sheetProtection selectLockedCells="1" algorithmName="workpaper&amp;excel"/>')
    expect(exportedSheetXml).not.toContain('<sheetProtection sheet="1"/>')

    const engine = new SpreadsheetEngine({ workbookName: 'sheet-protection-options-engine' })
    engine.importSnapshot(imported.snapshot)
    const exportedFromEngineSheetXml = sheetXml(exportXlsx(engine.exportSnapshot()))
    expect(exportedFromEngineSheetXml).toContain('<sheetProtection selectLockedCells="1" algorithmName="workpaper&amp;excel"/>')
    expect(exportedFromEngineSheetXml).not.toContain('<sheetProtection sheet="1"/>')
  })
})

function buildSheetProtectionWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([['protected']])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Protected')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceSheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sourceSheetXml.replace('<sheetData>', '<sheetProtection selectLockedCells="1" algorithmName="workpaper&amp;excel"/><sheetData>'),
  )
  return zipSync(zip)
}

function sheetXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
}
