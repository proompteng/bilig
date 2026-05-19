import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('large simple XLSX conditional format artifacts', () => {
  it('does not retain raw conditional-format XML for faithfully typed simple rules', () => {
    const bytes = buildConditionalFormatWorkbook(
      '<conditionalFormatting sqref="A1:A2"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>',
    )
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('simple conditional formatting should use streamed typed records instead of inflating worksheet XML')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'simple-conditional-format.xlsx', zip, { minByteLength: 0 })
    const metadata = imported?.snapshot.sheets[0]?.metadata

    expect(metadata?.conditionalFormats).toEqual([
      {
        id: 'xlsx-cf:Data:A1:A2:1',
        range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A2' },
        rule: { kind: 'cellIs', operator: 'greaterThan', values: [3] },
        style: {},
        priority: 1,
      },
    ])
    expect(metadata?.conditionalFormatArtifacts).toBeUndefined()
  })

  it('retains raw conditional-format XML when exact style artifacts are still needed', () => {
    const bytes = buildConditionalFormatWorkbook(
      '<conditionalFormatting sqref="A1:A2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>',
      {
        'xl/styles.xml': strToU8(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></dxf></dxfs>',
            '</styleSheet>',
          ].join(''),
        ),
      },
    )
    const imported = tryImportLargeSimpleXlsx(bytes, 'styled-conditional-format.xlsx', readXlsxZipEntriesLazy(bytes), {
      minByteLength: 0,
    })

    expect(imported?.snapshot.sheets[0]?.metadata?.conditionalFormats).toHaveLength(1)
    expect(imported?.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml).toContain('dxfId="0"')
  })
})

function buildConditionalFormatWorkbook(conditionalFormattingXml: string, extraEntries: Record<string, Uint8Array> = {}): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A2"/>
  <sheetData><row r="1"><c r="A1"><v>7</v></c></row><row r="2"><c r="A2"><v>1</v></c></row></sheetData>
  ${conditionalFormattingXml}
</worksheet>`),
    ...extraEntries,
  })
}
