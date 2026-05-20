import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { tryInspectLargeSimpleXlsxHeadless } from '../xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('large simple XLSX data validations', () => {
  it('streams supported data validations into sheet metadata without inflating worksheet XML', () => {
    const bytes = buildWorkbook(dataValidationWorksheetXml())
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'data-validations.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(imported?.stats.dataValidationCount).toBe(4)
    expect(imported?.snapshot.sheets[0]?.metadata?.validations).toEqual([
      {
        range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A1' },
        rule: { kind: 'list', values: ['Open', 'Closed'] },
        allowBlank: true,
      },
      {
        range: { sheetName: 'Data', startAddress: 'B1', endAddress: 'B1' },
        rule: { kind: 'decimal', operator: 'between', values: [0, 1] },
        promptTitle: 'Ratio',
        promptMessage: 'Enter 0 to 1',
        errorStyle: 'warning',
        errorTitle: 'Invalid ratio',
        errorMessage: 'Use a decimal between 0 and 1.',
      },
      {
        range: { sheetName: 'Data', startAddress: 'C1', endAddress: 'C1' },
        rule: { kind: 'whole', operator: 'between', values: [1, 10] },
      },
      {
        range: { sheetName: 'Data', startAddress: 'C2', endAddress: 'C2' },
        rule: { kind: 'whole', operator: 'between', values: [1, 10] },
      },
    ])
  })

  it('keeps headless verifier counts for supported data validations', () => {
    const bytes = buildWorkbook(dataValidationWorksheetXml())

    const inspected = tryInspectLargeSimpleXlsxHeadless(bytes, 'headless-data-validations.xlsx', readXlsxZipEntriesLazy(bytes), {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(inspected?.stats.dataValidationCount).toBe(4)
    expect(inspected?.stats.cellCount).toBe(3)
  })

  it('counts broad headless data-validation refs without requiring product metadata materialization', () => {
    const refCount = 5_000
    const sqref = Array.from({ length: refCount }, (_value, index) => `A${String(index + 1)}`).join(' ')
    const bytes = buildWorkbook(
      [
        worksheetPrefix(),
        `<dataValidations count="1"><dataValidation type="whole" operator="between" sqref="${sqref}"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>`,
        '</worksheet>',
      ].join(''),
    )

    const inspected = tryInspectLargeSimpleXlsxHeadless(bytes, 'headless-broad-data-validations.xlsx', readXlsxZipEntriesLazy(bytes), {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(inspected?.stats.dataValidationCount).toBe(refCount)
    expect(inspected?.stats.cellCount).toBe(3)
  })

  it('rejects unsupported data validations from the large-simple product path', () => {
    const bytes = buildWorkbook(
      [
        worksheetPrefix(),
        '<dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>A1&gt;0</formula1></dataValidation></dataValidations>',
        '</worksheet>',
      ].join(''),
    )

    expect(
      tryImportLargeSimpleXlsx(bytes, 'unsupported-data-validations.xlsx', readXlsxZipEntriesLazy(bytes), { minByteLength: 0 }),
    ).toBeNull()
    expect(
      tryInspectLargeSimpleXlsxHeadless(bytes, 'unsupported-data-validations.xlsx', readXlsxZipEntriesLazy(bytes), {
        minByteLength: 0,
      }),
    ).toBeNull()
  })
})

function dataValidationWorksheetXml(): string {
  return [
    worksheetPrefix(),
    '<dataValidations count="3">',
    '<dataValidation type="list" allowBlank="1" sqref="A1"><formula1>"Open,Closed"</formula1></dataValidation>',
    '<dataValidation type="decimal" operator="between" promptTitle="Ratio" prompt="Enter 0 to 1" ',
    'errorStyle="warning" errorTitle="Invalid ratio" error="Use a decimal between 0 and 1." sqref="B1">',
    '<formula1>0</formula1><formula2>1</formula2></dataValidation>',
    '<dataValidation type="whole" operator="between" sqref="C1 C2"><formula1>1</formula1><formula2>10</formula2></dataValidation>',
    '</dataValidations>',
    '</worksheet>',
  ].join('')
}

function worksheetPrefix(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:C1"/>',
    '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Open</t></is></c><c r="B1"><v>0.5</v></c><c r="C1"><v>5</v></c></row></sheetData>',
  ].join('')
}

function buildWorkbook(worksheetXml: string): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(worksheetXml),
  })
}
