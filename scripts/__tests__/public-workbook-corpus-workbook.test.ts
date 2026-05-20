import { deflateRawSync } from 'node:zlib'

import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx } from '../../packages/excel-import/src/index.js'
import { ValueTag } from '../../packages/protocol/src/enums.js'
import type { WorkbookSnapshot } from '../../packages/protocol/src/types.js'
import { extractFormulaOracles, inspectWorkbookFootprint, inspectWorkbookFootprintForWorker } from '../public-workbook-corpus-workbook.ts'

describe('public workbook corpus workbook helpers', () => {
  it('extracts formula oracles from broad sparse worksheet refs', () => {
    const oracles = extractFormulaOracles(buildBroadSparseWorkbookBytes())

    expect(oracles).toEqual([
      {
        sheetName: 'Sparse',
        address: 'XFD512',
        expected: { tag: ValueTag.Number, value: 42 },
      },
    ])
  }, 15_000)

  it('records explicit used ranges from actual populated cells instead of broad worksheet refs', () => {
    const footprint = inspectWorkbookFootprint(buildBroadSparseWorkbookBytes(), 'sparse.xlsx')

    expect(footprint.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 1,
      formulaCellCount: 1,
      valueCellCount: 1,
    })
    expect(footprint.workbookMetadata.dimensions).toEqual([
      {
        sheetName: 'Sparse',
        rowCount: 512,
        columnCount: 16_384,
        nonEmptyCellCount: 1,
        usedRange: { startRow: 511, startColumn: 16_383, endRow: 511, endColumn: 16_383 },
      },
    ])
  }, 15_000)

  it('counts raw XLSX pivot table parts even when semantic pivot import is unavailable', () => {
    const footprint = inspectWorkbookFootprint(exportXlsx(buildPivotWorkbookSnapshot()), 'raw-pivot.xlsx')

    expect(footprint.featureCounts.pivotCount).toBe(1)
  }, 15_000)

  it('does not skip valued cells after adjacent self-closing blank cells in XLSX XML', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithAdjacentBlankCellElements(), 'adjacent-blanks.xlsx')

    expect(footprint.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 2,
      formulaCellCount: 0,
      valueCellCount: 2,
    })
    expect(footprint.workbookMetadata.dimensions).toEqual([
      {
        sheetName: 'Sheet1',
        rowCount: 1,
        columnCount: 5,
        nonEmptyCellCount: 2,
        usedRange: { startRow: 0, startColumn: 2, endRow: 0, endColumn: 4 },
      },
    ])
  })

  it('uses the worker footprint path without materializing worksheet XML', async () => {
    const footprint = await inspectWorkbookFootprintForWorker(buildBroadSparseWorkbookBytes(), 'sparse.xlsx')

    expect(footprint.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 1,
      formulaCellCount: 1,
      valueCellCount: 1,
    })
    expect(footprint.workbookMetadata.dimensions[0]?.usedRange).toEqual({
      startRow: 511,
      startColumn: 16_383,
      endRow: 511,
      endColumn: 16_383,
    })
  }, 15_000)

  it('marks value-only XLSX files with hyperlinks and drawings eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithHyperlinksAndDrawing(), 'hyperlinks-drawing.xlsx')

    expect(footprint.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 1,
      formulaCellCount: 0,
      valueCellCount: 1,
    })
    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('marks value-only XLSX files with printer settings eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithPrinterSettings(), 'printer-settings.xlsx')

    expect(footprint.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 1,
      formulaCellCount: 0,
      valueCellCount: 1,
    })
    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('marks value-only XLSX files with rich shared strings eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithRichSharedString(), 'rich-shared-string.xlsx')

    expect(footprint.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 1,
      formulaCellCount: 0,
      valueCellCount: 1,
    })
    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('preserves large simple import eligibility through the isolated worker footprint parser', async () => {
    const footprint = await inspectWorkbookFootprintForWorker(buildWorkbookWithHyperlinksAndDrawing(), 'hyperlinks-drawing.xlsx')

    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('marks XLSX files with simple formula cells eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithFormulaCell(), 'formula.xlsx')

    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('marks XLSX files with auto filters eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithAutoFilter(), 'auto-filter.xlsx')

    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('marks XLSX files with table parts eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithTablePart(), 'table.xlsx')

    expect(footprint.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 1,
      tableCount: 1,
      formulaCellCount: 0,
      valueCellCount: 1,
    })
    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('marks conditional formatting eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithConditionalFormatting(), 'conditional-format.xlsx')

    expect(footprint.featureCounts.conditionalFormatCount).toBe(1)
    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('marks supported data validations eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithSupportedDataValidations(), 'data-validations.xlsx')

    expect(footprint.featureCounts.dataValidationCount).toBe(3)
    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('keeps unsupported data validations out of the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithUnsupportedDataValidation(), 'unsupported-data-validation.xlsx')

    expect(footprint.featureCounts.dataValidationCount).toBe(0)
    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: false, blockers: ['unsupported-data-validations=1'] })
  })

  it('marks shared formula worksheets eligible for the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithSharedFormulaCell(), 'shared-formula.xlsx')

    expect(footprint.featureCounts.formulaCellCount).toBe(2)
    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
  })

  it('keeps structured-reference formula worksheets out of the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithStructuredReferenceFormula(), 'structured-reference-formula.xlsx')

    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: false, blockers: ['unsupported-formula-cells=1'] })
  })

  it('keeps chart packages out of the large simple import budget', () => {
    const footprint = inspectWorkbookFootprint(buildWorkbookWithChartPackage(), 'chart.xlsx')

    expect(footprint.largeSimpleXlsxImport).toEqual({ eligible: false, blockers: ['unsupported-package-parts=1'] })
  })
})

function buildBroadSparseWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet: XLSX.WorkSheet = {
    XFD512: { t: 'n', f: '40+2', v: 42 },
    '!ref': 'A1:XFD512',
  }
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sparse')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildWorkbookWithAdjacentBlankCellElements(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1">',
        '<c r="A1" s="1"/><c r="B1" s="1"/><c r="C1"><v>7</v></c><c r="D1" s="1"/><c r="E1" t="str"><v>ok</v></c>',
        '</row></sheetData></worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithHyperlinksAndDrawing(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<hyperlinks><hyperlink ref="A1" r:id="rIdHyperlink1"/></hyperlinks>',
        '<drawing r:id="rIdDrawing1"/></worksheet>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/_rels/sheet1.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdHyperlink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ',
        'Target="https://example.com" TargetMode="External"/>',
        '<Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ',
        'Target="../drawings/drawing1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/drawings/drawing1.xml',
      text: '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:absoluteAnchor/></xdr:wsDr>',
    },
  ])
}

function buildWorkbookWithRichSharedString(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" ',
        'Target="sharedStrings.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/sharedStrings.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">',
        '<si><r><rPr><b/></rPr><t>Rich</t></r><r><t> Text</t></r></si></sst>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithPrinterSettings(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<pageSetup orientation="landscape" r:id="rIdPrinterSettings1"/></worksheet>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/_rels/sheet1.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdPrinterSettings1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" ',
        'Target="../printerSettings/printerSettings1.bin"/></Relationships>',
      ].join(''),
    },
    { path: 'xl/printerSettings/printerSettings1.bin', text: 'raw-printer-settings' },
  ])
}

function buildWorkbookWithFormulaCell(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithAutoFilter(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<autoFilter ref="A1:A3"/></worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithTablePart(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts></worksheet>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/_rels/sheet1.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" ',
        'Target="../tables/table1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/tables/table1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:A2">',
        '<tableColumns count="1"><tableColumn id="1" name="Amount"/></tableColumns></table>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithConditionalFormatting(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
        '</worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithSupportedDataValidations(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<dataValidations count="2">',
        '<dataValidation type="list" sqref="A1"><formula1>"Open,Closed"</formula1></dataValidation>',
        '<dataValidation type="whole" operator="between" sqref="B1 B2"><formula1>1</formula1><formula2>10</formula2></dataValidation>',
        '</dataValidations>',
        '</worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithUnsupportedDataValidation(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>A1&gt;0</formula1></dataValidation></dataValidations>',
        '</worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithSharedFormulaCell(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row>',
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row></sheetData></worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithStructuredReferenceFormula(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1"><f>SUM(Table1[Amount])</f><v>7</v></c></row></sheetData></worksheet>',
      ].join(''),
    },
  ])
}

function buildWorkbookWithChartPackage(): Uint8Array {
  return buildZip([
    {
      path: 'xl/workbook.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ].join(''),
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ',
        'Target="worksheets/sheet1.xml"/></Relationships>',
      ].join(''),
    },
    {
      path: 'xl/worksheets/sheet1.xml',
      text: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>',
      ].join(''),
    },
    { path: 'xl/charts/chart1.xml', text: '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>' },
  ])
}

function buildZip(entries: readonly { readonly path: string; readonly text: string }[]): Uint8Array {
  const localParts: Buffer[] = []
  const centralParts: Buffer[] = []
  let localOffset = 0
  for (const entry of entries) {
    const name = Buffer.from(entry.path)
    const uncompressed = Buffer.from(entry.text)
    const compressed = deflateRawSync(uncompressed)
    const localHeader = Buffer.alloc(30 + name.length)
    localHeader.writeUInt32LE(0x04034b50, 0)
    localHeader.writeUInt16LE(20, 4)
    localHeader.writeUInt16LE(8, 8)
    localHeader.writeUInt32LE(compressed.length, 18)
    localHeader.writeUInt32LE(uncompressed.length, 22)
    localHeader.writeUInt16LE(name.length, 26)
    name.copy(localHeader, 30)
    localParts.push(localHeader, compressed)

    const centralHeader = Buffer.alloc(46 + name.length)
    centralHeader.writeUInt32LE(0x02014b50, 0)
    centralHeader.writeUInt16LE(20, 4)
    centralHeader.writeUInt16LE(20, 6)
    centralHeader.writeUInt16LE(8, 10)
    centralHeader.writeUInt32LE(compressed.length, 20)
    centralHeader.writeUInt32LE(uncompressed.length, 24)
    centralHeader.writeUInt16LE(name.length, 28)
    centralHeader.writeUInt32LE(localOffset, 42)
    name.copy(centralHeader, 46)
    centralParts.push(centralHeader)
    localOffset += localHeader.length + compressed.length
  }
  const centralDirectory = Buffer.concat(centralParts)
  const endOfCentralDirectory = Buffer.alloc(22)
  endOfCentralDirectory.writeUInt32LE(0x06054b50, 0)
  endOfCentralDirectory.writeUInt16LE(entries.length, 8)
  endOfCentralDirectory.writeUInt16LE(entries.length, 10)
  endOfCentralDirectory.writeUInt32LE(centralDirectory.length, 12)
  endOfCentralDirectory.writeUInt32LE(localOffset, 16)
  return Buffer.concat([...localParts, centralDirectory, endOfCentralDirectory])
}

function buildPivotWorkbookSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'raw-pivot',
      metadata: {
        pivots: [
          {
            name: 'RevenuePivot',
            sheetName: 'Pivot',
            address: 'A1',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Revenue', summarizeBy: 'sum' }],
            rows: 3,
            cols: 2,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Revenue' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 12 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 8 },
        ],
      },
      {
        id: 2,
        name: 'Pivot',
        order: 1,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Sum of Revenue' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 12 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 8 },
        ],
      },
    ],
  }
}
