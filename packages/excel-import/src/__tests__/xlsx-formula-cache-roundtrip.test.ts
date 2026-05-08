import { strFromU8, unzipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { describe, expect, it } from 'vitest'

import { exportXlsx, importXlsx } from '../index.js'

describe('formula cache roundtrip', () => {
  it('preserves cached formula result values through import and export', () => {
    const imported = importXlsx(buildFormulaCacheWorkbookBytes(), 'formula-cache-source.xlsx')
    const importedCells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(importedCells.get('B1')).toMatchObject({ formula: 'A1+A2', value: 1550 })
    expect(importedCells.get('B2')).toMatchObject({ formula: 'SUM(A1:A2)', value: 1550 })
    expect(importedCells.get('C2')).toMatchObject({ formula: 'B1/A1', value: 1.2916666666666667 })

    const exported = exportXlsx(imported.snapshot)
    const exportedSheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(cellXml(exportedSheetXml, 'B1')).toContain('<v>1550</v>')
    expect(cellXml(exportedSheetXml, 'B2')).toContain('<v>1550</v>')
    expect(cellXml(exportedSheetXml, 'C2')).toContain('<v>1.2916666666666667</v>')

    const reimported = importXlsx(exported, 'formula-cache-roundtrip.xlsx')
    const reimportedCells = new Map(reimported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(reimportedCells.get('B1')).toMatchObject({ formula: 'A1+A2', value: 1550 })
    expect(reimportedCells.get('B2')).toMatchObject({ formula: 'SUM(A1:A2)', value: 1550 })
    expect(reimportedCells.get('C2')).toMatchObject({ formula: 'B1/A1', value: 1.2916666666666667 })
  })
})

function buildFormulaCacheWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    [1200, null, null],
    [350, null, null],
  ])
  sheet.B1 = { t: 'n', f: 'A1+A2', v: 1550 }
  sheet.B2 = { t: 'n', f: 'SUM(A1:A2)', v: 1550 }
  sheet.C2 = { t: 'n', f: 'B1/A1', v: 1.2916666666666667 }
  sheet['!ref'] = 'A1:C2'

  XLSX.utils.book_append_sheet(workbook, sheet, 'FormulaCache')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function cellXml(sheetXml: string, address: string): string {
  return sheetXml.match(new RegExp(`<c[^>]* r="${address}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0] ?? ''
}
