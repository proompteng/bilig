import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { SpreadsheetEngine } from '@bilig/core'

import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX worksheet dimension roundtrip', () => {
  it('preserves sheet defaults plus column and row geometry metadata', () => {
    const imported = importXlsx(buildDimensionWorkbookBytes(), 'dimension-roundtrip.xlsx')

    const exported = exportXlsx(imported.snapshot)
    const exportedSheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(exportedSheetXml).toContain(
      '<sheetFormatPr baseColWidth="12" defaultColWidth="13.75" defaultRowHeight="14.6" customHeight="1" outlineLevelRow="2" outlineLevelCol="3" thickTop="1" thickBottom="1"/>',
    )
    expect(exportedSheetXml).toContain('<col min="1" max="1" width="23.3828125" customWidth="1" bestFit="1"/>')
    expect(exportedSheetXml).toContain(
      '<col min="3" max="5" width="3.3828125" customWidth="1" bestFit="1" outlineLevel="2" collapsed="1"/>',
    )
    expect(exportedSheetXml).toContain('<row r="1" ht="45.45" customHeight="1" thickBot="1">')
    expect(exportedSheetXml).toContain('<row r="4" ht="15.9"/>')
    expect(exportedSheetXml).toContain('<row r="24" ht="15" thickTop="1" thickBot="1"/>')

    const reimported = importXlsx(exported, 'dimension-roundtrip-exported.xlsx')
    const reexportedSheetXml = strFromU8(unzipSync(exportXlsx(reimported.snapshot))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    expect(reexportedSheetXml).toContain(
      '<col min="3" max="5" width="3.3828125" customWidth="1" bestFit="1" outlineLevel="2" collapsed="1"/>',
    )
    expect(reexportedSheetXml).toContain('<row r="24" ht="15" thickTop="1" thickBot="1"/>')

    const engine = new SpreadsheetEngine({ workbookName: 'dimension-roundtrip-engine' })
    engine.importSnapshot(imported.snapshot)
    const exportedFromEngineSheetXml = strFromU8(
      unzipSync(exportXlsx(engine.exportSnapshot()))['xl/worksheets/sheet1.xml'] ?? new Uint8Array(),
    )
    expect(exportedFromEngineSheetXml).toContain(
      '<sheetFormatPr baseColWidth="12" defaultColWidth="13.75" defaultRowHeight="14.6" customHeight="1" outlineLevelRow="2" outlineLevelCol="3" thickTop="1" thickBottom="1"/>',
    )
    expect(exportedFromEngineSheetXml).toContain(
      '<col min="3" max="5" width="3.3828125" customWidth="1" bestFit="1" outlineLevel="2" collapsed="1"/>',
    )
    expect(exportedFromEngineSheetXml).toContain('<row r="24" ht="15" thickTop="1" thickBot="1"/>')

    const editedSnapshot = structuredClone(imported.snapshot)
    const editedMetadata = editedSnapshot.sheets[0].metadata!
    editedMetadata.rows = [...(editedMetadata.rows ?? []), { id: 'row:9', index: 9, size: 20 }]
    editedMetadata.columns = [...(editedMetadata.columns ?? []), { id: 'col:7', index: 7, size: 72 }]
    const editedSheetXml = strFromU8(unzipSync(exportXlsx(editedSnapshot))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    expect(editedSheetXml).toContain('<row r="10" ht="20" customHeight="1"/>')
    expect(editedSheetXml).toMatch(/<col min="8" max="8" width="[^"]+" customWidth="1"\/>/u)
  })
})

function buildDimensionWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[123]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Dimensions')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(dimensionWorksheetXml)
  return zipSync(zip)
}

const dimensionWorksheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<dimension ref="A1:E24"/>',
  '<sheetFormatPr baseColWidth="12" defaultColWidth="13.75" defaultRowHeight="14.6" customHeight="1" outlineLevelRow="2" outlineLevelCol="3" thickTop="1" thickBottom="1"/>',
  '<cols>',
  '<col min="1" max="1" width="23.3828125" customWidth="1" bestFit="1"/>',
  '<col min="3" max="5" width="3.3828125" customWidth="1" bestFit="1" outlineLevel="2" collapsed="1"/>',
  '</cols>',
  '<sheetData>',
  '<row r="1" ht="45.45" customHeight="1" thickBot="1"><c r="A1"><v>123</v></c></row>',
  '<row r="4" ht="15.9"/>',
  '<row r="24" ht="15" thickTop="1" thickBot="1"/>',
  '</sheetData>',
  '</worksheet>',
].join('')
