import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('large simple XLSX import materialization lifetime', () => {
  it('materializes independent sheets before global resolution phases', () => {
    const bytes = buildIndependentWorkbook([
      {
        name: 'First',
        path: 'xl/worksheets/sheet1.xml',
        xml: worksheetXml('A', 7),
      },
      {
        name: 'Second',
        path: 'xl/worksheets/sheet2.xml',
        xml: worksheetXml('B', 11),
      },
    ])
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet1 XML should be streamed instead of inflated')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet2.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet2 XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'independent-sheets.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(imported?.snapshot.sheets.map((sheet) => sheet.cells)).toEqual([
      [
        { address: 'A1', value: 7 },
        { address: 'B1', value: 'A inline' },
      ],
      [
        { address: 'A1', value: 11 },
        { address: 'B1', value: 'B inline' },
      ],
    ])
    expect(imported?.stats.phaseTelemetry.map((entry) => entry.phase)).toEqual([
      'zip-setup',
      'worksheet-scan',
      'metadata-parsing',
      'public-snapshot-materialization',
      'shared-string-resolution',
      'style-parsing',
      'zip-source-release',
    ])
  })
})

function worksheetXml(label: string, value: number): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B1"/>',
    `<sheetData><row r="1"><c r="A1"><v>${String(value)}</v></c><c r="B1" t="inlineStr"><is><t>${label} inline</t></is></c></row></sheetData>`,
    '</worksheet>',
  ].join('')
}

function buildIndependentWorkbook(sheets: readonly { readonly name: string; readonly path: string; readonly xml: string }[]): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets
    .map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${String(index + 1)}" r:id="rId${String(index + 1)}"/>`)
    .join('')}</sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets
  .map(
    (sheet, index) =>
      `<Relationship Id="rId${String(index + 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${sheet.path.slice('xl/'.length)}"/>`,
  )
  .join('')}
</Relationships>`),
    ...Object.fromEntries(sheets.map((sheet) => [sheet.path, strToU8(sheet.xml)])),
  })
}
