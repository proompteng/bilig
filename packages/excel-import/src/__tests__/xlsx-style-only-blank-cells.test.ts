import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { stripNoOpEmptyRowsFromXlsx, stripStyleOnlyBlankCellsForSheetJs } from '../xlsx-style-only-blank-cells.js'

describe('stripStyleOnlyBlankCellsForSheetJs', () => {
  it('removes no-op empty rows while preserving meaningful row metadata', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetFormatPr defaultRowHeight="11.25"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2" ht="11.25" customHeight="1"/>
    <row r="3" spans="1:6"/>
    <row r="4" ht="12" customHeight="1"/>
    <row r="5" hidden="1"/>
    <row r="6" ht="11.25" customHeight="1"></row>
    <row r="7"><c r="A7" s="2"/></row>
    <row r="8" s="2" customFormat="1"><c r="A8" s="2"/></row>
    <row r="9" ht="11.25" customHeight="1" x14ac:dyDescent="0.25"/>
  </sheetData>
</worksheet>`),
    })
    const stripped = stripStyleOnlyBlankCellsForSheetJs(bytes, unzipSync(bytes))
    const sheetXml = strFromU8(unzipSync(stripped)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(sheetXml).toContain('<row r="1"><c r="A1"><v>1</v></c></row>')
    expect(sheetXml).not.toContain('r="2"')
    expect(sheetXml).not.toContain('r="3"')
    expect(sheetXml).toContain('<row r="4" ht="12" customHeight="1"/>')
    expect(sheetXml).toContain('<row r="5" hidden="1"/>')
    expect(sheetXml).not.toContain('r="6"')
    expect(sheetXml).not.toContain('r="7"')
    expect(sheetXml).toContain('<row r="8" s="2" customFormat="1"></row>')
    expect(sheetXml).not.toContain('r="9"')
  })

  it('can strip no-op rows while retaining blank style cells for artifact scans', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetFormatPr defaultRowHeight="11.25"/>
  <sheetData>
    <row r="1" ht="11.25" customHeight="1"/>
    <row r="2"><c r="A2" s="2"/></row>
  </sheetData>
</worksheet>`),
    })
    const stripped = stripNoOpEmptyRowsFromXlsx(bytes, unzipSync(bytes))
    const sheetXml = strFromU8(unzipSync(stripped)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(sheetXml).not.toContain('r="1"')
    expect(sheetXml).toContain('<row r="2"><c r="A2" s="2"/></row>')
  })

  it('does not mutate the source zip while producing SheetJS parser bytes', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="2"/></row>
  </sheetData>
</worksheet>`),
    })
    const zip = unzipSync(bytes)
    const stripped = stripStyleOnlyBlankCellsForSheetJs(bytes, zip)
    const originalSheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const strippedSheetXml = strFromU8(unzipSync(stripped)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(originalSheetXml).toContain('<c r="A1" s="2"/>')
    expect(strippedSheetXml).not.toContain('<c r="A1" s="2"/>')
  })
})
