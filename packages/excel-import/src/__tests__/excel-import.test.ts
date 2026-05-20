import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { projectWorkbookSemanticSnapshot as projectSupportedSnapshotSemantics, readRuntimeImage, SpreadsheetEngine } from '@bilig/core'
import {
  CSV_CONTENT_TYPE,
  EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES,
  InvalidXlsxZipContainerError,
  LEGACY_XLS_CONTENT_TYPE,
  WORKBOOK_IMPORT_CONTENT_TYPES,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  externalWorkbookReferencesWarning,
  exportXlsx,
  importCsv,
  importWorkbookFile,
  importXlsx,
  readImportedXlsxCellStyle,
} from '../index.js'

function buildWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()

  const sheet1 = XLSX.utils.aoa_to_sheet([
    [1, 2],
    [3, null],
  ])
  sheet1['C1'] = { t: 'n', f: 'A1+B1', z: '0.00' }
  sheet1['!ref'] = 'A1:C2'
  sheet1['!cols'] = [{ wpx: 120 }, { wch: 10 }, { wpx: 80 }]
  sheet1['!rows'] = [{ hpx: 30 }, { hpt: 18 }]
  sheet1['!merges'] = [{ s: { r: 3, c: 0 }, e: { r: 3, c: 1 } }]

  const sheet2 = XLSX.utils.aoa_to_sheet([['hello'], [true]])
  sheet2['A1'] = {
    ...sheet2['A1'],
    c: [{ a: 'Greg', t: 'comment' }],
  }

  XLSX.utils.book_append_sheet(workbook, sheet1, 'Sheet1')
  XLSX.utils.book_append_sheet(workbook, sheet2, 'Sheet2')
  workbook.Workbook = {
    Names: [
      { Name: 'InputValue', Ref: 'Sheet1!$A$1' },
      { Name: 'InputBlock', Ref: 'Sheet1!$A$1:$B$2' },
    ],
  }

  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildExternalLinkCacheWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[]])
  sheet.A1 = { t: 'n', f: "'[1]External Data'!A1+'[1]External Data'!A2", v: 5 }
  sheet['!ref'] = 'A1:A1'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Report')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'])
      .replace(/<workbook\b([^>]*)>/u, (match) =>
        match.includes('xmlns:r=')
          ? match
          : match.replace('>', ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'),
      )
      .replace('</workbook>', '<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="External Data"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="1"><cell r="A1"><v>2</v></cell></row>',
      '<row r="2"><cell r="A2"><v>3</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/source.xlsx" TargetMode="External"/>' +
      '</Relationships>',
  )
  return zipSync(zip)
}

function buildExternalGetPivotDataLinkCacheWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[]])
  sheet.A1 = {
    t: 'n',
    f: 'GETPIVOTDATA("Amount",\'[1]External Pivot\'!$G$3,"Region","East")',
    v: 15,
  }
  sheet['!ref'] = 'A1:A1'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Report')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'])
      .replace(/<workbook\b([^>]*)>/u, (match) =>
        match.includes('xmlns:r=')
          ? match
          : match.replace('>', ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'),
      )
      .replace('</workbook>', '<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="External Pivot"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="3"><cell r="G3" t="str"><v>Row Labels</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/pivot-source.xlsx" TargetMode="External"/>' +
      '</Relationships>',
  )
  return zipSync(zip)
}

function buildUnsupportedFunctionCacheWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[]])
  sheet.A1 = {
    t: 'n',
    f: '_xldudf_WISEPRICE(B1,"Shares Outstanding")',
    v: 14935800000,
  }
  sheet.B1 = { t: 's', v: 'AAPL' }
  sheet.C1 = { t: 's', f: '_FV(B1,"Ticker symbol",TRUE)', v: 'AAPL' }
  sheet['!ref'] = 'A1:C1'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildBinaryWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Name', 'Value'],
    ['alpha', 12],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['notes']]), 'Sheet2')
  return XLSX.write(workbook, { bookType: 'xlsb', type: 'buffer' })
}

function buildLegacyWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Department', 'Amount'],
    ['Operations', 1250],
    ['Finance', 1800],
  ])
  sheet.C2 = { t: 'n', f: 'B2+B3', v: 3050 }
  sheet['!ref'] = 'A1:C3'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Salary')
  return XLSX.write(workbook, { bookType: 'xls', type: 'buffer' })
}

function buildNamespacedFormulaWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[1], [2]])
  sheet.A3 = { t: 'n', f: 'msoxl:=SUM(A1:A2)', v: 3 }
  sheet.B3 = { t: 'n', f: 'of:=SUM(A1:A2)', v: 3 }
  sheet['!ref'] = 'A1:B3'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Expenses')
  return XLSX.write(workbook, { bookType: 'ods', type: 'buffer' })
}

function buildMacroEnabledWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['safe value']]), 'Sheet1')
  workbook.Workbook = {
    WBProps: { CodeName: 'ThisWorkbook' },
    Sheets: [{ name: 'Sheet1', CodeName: 'Sheet1' }],
  }
  workbook.vbaraw = new Uint8Array([1, 2, 3, 4])
  return XLSX.write(workbook, { bookType: 'xlsm', type: 'buffer', bookVBA: true })
}

function readZipUint16(bytes: Uint8Array, offset: number): number {
  const low = bytes[offset]
  const high = bytes[offset + 1]
  if (low === undefined || high === undefined) {
    throw new Error('Invalid ZIP fixture')
  }
  return low | (high << 8)
}

function buildCorruptZipBackedWorkbook(): Uint8Array {
  const bytes = zipSync({ 'xl/workbook.xml': strToU8('a'.repeat(1000)) })
  const nameLength = readZipUint16(bytes, 26)
  const extraLength = readZipUint16(bytes, 28)
  const compressedDataStart = 30 + nameLength + extraLength
  const originalByte = bytes[compressedDataStart]
  if (originalByte === undefined) {
    throw new Error('Invalid ZIP fixture')
  }
  const corrupted = new Uint8Array(bytes)
  corrupted[compressedDataStart] = originalByte ^ 0xff
  return corrupted
}

function buildGenericWorkflowWorkbookFixture(shape: 'multi-sheet-operations' | 'single-sheet-planning'): Uint8Array {
  const workbook = XLSX.utils.book_new()
  if (shape === 'multi-sheet-operations') {
    const dashboard = XLSX.utils.aoa_to_sheet([
      ['OPERATIONS DASHBOARD', null, null, null],
      [],
      ['Metric', 'Value'],
      ['Total budget'],
      ['Open balance'],
      ['Completion rate'],
    ])
    dashboard.B4 = { t: 'n', f: 'SUM(Ledger!F:F)' }
    dashboard.B5 = { t: 'n', f: 'SUMIF(Ledger!H:H,"Open",Ledger!G:G)' }
    dashboard.B6 = { t: 'n', f: 'IF(B4>0,1-B5/B4,0)' }
    dashboard['!ref'] = 'A1:D6'
    dashboard['!cols'] = [{ wpx: 180 }, { wpx: 118 }, { wpx: 96 }, { wpx: 96 }]
    dashboard['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
    dashboard['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }]

    const ledger = XLSX.utils.aoa_to_sheet([
      ['OPERATIONS LEDGER', null, null, null, null, null, null, null],
      [],
      ['ID', 'Date', 'Owner', 'Workstream', 'Category', 'Budget', 'Open Balance', 'Status'],
      ['OP001', 45292, 'Facilities', 'Office refresh', 'Capital', 12000, null, 'Open'],
      ['OP002', 45323, 'Engineering', 'Data migration', 'Platform', 18000, null, 'Open'],
    ])
    ledger.G4 = { t: 'n', f: 'F4-SUMIF(Rollforward!$B:$B,A4,Rollforward!$E:$E)' }
    ledger.G5 = { t: 'n', f: 'F5-SUMIF(Rollforward!$B:$B,A5,Rollforward!$E:$E)' }
    ledger['!ref'] = 'A1:H5'
    ledger['!cols'] = [{ wpx: 132 }, { wpx: 96 }, { wpx: 142 }, { wpx: 210 }, { wpx: 138 }, { wpx: 118 }, { wpx: 138 }, { wpx: 92 }]
    ledger['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
    ledger['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }]

    const rollforward = XLSX.utils.aoa_to_sheet([
      ['ROLLFORWARD', null, null, null, null],
      [],
      ['Period', 'Item ID', 'Description', 'Monthly Change', 'Cumulative Change'],
      ['Jan 2024', 'OP001', 'Office refresh'],
      ['Feb 2024', 'OP001', 'Office refresh'],
      ['Mar 2024', 'OP002', 'Data migration'],
    ])
    rollforward.D4 = { t: 'n', f: 'VLOOKUP(B4,Ledger!A:F,6,FALSE())/12' }
    rollforward.E4 = { t: 'n', f: 'D4' }
    rollforward.D5 = { t: 'n', f: 'VLOOKUP(B5,Ledger!A:F,6,FALSE())/12' }
    rollforward.E5 = { t: 'n', f: 'IF(B5=B4,E4+D5,D5)' }
    rollforward.D6 = { t: 'n', f: 'VLOOKUP(B6,Ledger!A:F,6,FALSE())/12' }
    rollforward.E6 = { t: 'n', f: 'IF(B6=B5,E5+D6,D6)' }
    rollforward['!ref'] = 'A1:E6'
    rollforward['!cols'] = [{ wpx: 112 }, { wpx: 96 }, { wpx: 210 }, { wpx: 126 }, { wpx: 148 }]
    rollforward['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
    rollforward['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }]

    XLSX.utils.book_append_sheet(workbook, dashboard, 'Dashboard')
    XLSX.utils.book_append_sheet(workbook, ledger, 'Ledger')
    XLSX.utils.book_append_sheet(workbook, rollforward, 'Rollforward')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Category'], ['Capital'], ['Platform']]), 'Lookups')
    return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
  }

  const planning = XLSX.utils.aoa_to_sheet([
    ['Monthly Planning Schedule', null, null, null, null, null, null, null, null],
    ['Owner', 'Workstream', 'Start Date', 'End Date', 'Budget', 'Jan 2026', 'Feb 2026', 'Planned', 'Remaining'],
    ['TenantWorks', 'Facilities platform', 46054, 46234, 6600],
    ['Blue Harbor', 'Insurance binder', 46023, 46388, 12000],
  ])
  planning.F3 = { t: 'n', f: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)' }
  planning.G3 = { t: 'n', f: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,2,1),0))-MAX($C3,DATE(2026,2,1))+1)/($D3-$C3+1),0),2)' }
  planning.H3 = { t: 'n', f: 'ROUND(SUM(F3:G3),2)' }
  planning.I3 = { t: 'n', f: 'ROUND(E3-H3,2)' }
  planning.F4 = { t: 'n', f: 'ROUND(IFERROR($E4*MAX(0,MIN($D4,EOMONTH(DATE(2026,1,1),0))-MAX($C4,DATE(2026,1,1))+1)/($D4-$C4+1),0),2)' }
  planning.G4 = { t: 'n', f: 'ROUND(IFERROR($E4*MAX(0,MIN($D4,EOMONTH(DATE(2026,2,1),0))-MAX($C4,DATE(2026,2,1))+1)/($D4-$C4+1),0),2)' }
  planning.H4 = { t: 'n', f: 'ROUND(SUM(F4:G4),2)' }
  planning.I4 = { t: 'n', f: 'ROUND(E4-H4,2)' }
  planning['!ref'] = 'A1:I4'
  planning['!cols'] = [
    { wpx: 168 },
    { wpx: 190 },
    { wpx: 104 },
    { wpx: 104 },
    { wpx: 118 },
    { wpx: 96 },
    { wpx: 96 },
    { wpx: 134 },
    { wpx: 138 },
  ]
  planning['!rows'] = [{ hpx: 30 }, { hpx: 24 }]
  planning['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }]
  XLSX.utils.book_append_sheet(workbook, planning, 'Monthly Plan')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

describe('excel import', () => {
  it('imports sheets, formulas, dimensions, and warnings from xlsx bytes', () => {
    const imported = importXlsx(buildWorkbook(), 'Quarterly Report.xlsx')

    expect(imported.workbookName).toBe('Quarterly Report')
    expect(imported.sheetNames).toEqual(['Sheet1', 'Sheet2'])
    expect(imported.snapshot.workbook.name).toBe('Quarterly Report')
    expect(imported.snapshot.sheets).toHaveLength(2)

    expect(imported.snapshot.sheets[0]).toMatchObject({
      name: 'Sheet1',
      metadata: {
        columns: [
          { index: 0, size: 120 },
          { index: 1, size: 65 },
          { index: 2, size: 80 },
        ],
        rows: [
          { index: 0, size: 30 },
          { index: 1, size: 18 },
        ],
        merges: [{ sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'B4' }],
      },
    })
    expect(imported.snapshot.sheets[0]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A1', value: 1 })]))
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: 'C1', formula: 'A1+B1', format: '0.00' })]),
    )
    expect(imported.snapshot.sheets[1]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A1', value: 'hello' })]))
    expect(imported.snapshot.sheets[1]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A2', value: true })]))

    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'InputBlock', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } },
      { name: 'InputValue', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' } },
    ])
    expect(imported.snapshot.sheets[1]?.metadata?.commentThreads).toEqual([
      {
        threadId: 'xlsx-comment:Sheet2:A1',
        sheetName: 'Sheet2',
        address: 'A1',
        comments: [{ id: 'xlsx-comment:Sheet2:A1:1', body: 'comment', authorDisplayName: 'Greg' }],
      },
    ])
    expect(imported.warnings).toEqual([])
    expect(imported.preview.workbookName).toBe('Quarterly Report')
    expect(imported.preview.sheetCount).toBe(2)
    expect(imported.preview.sheets).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          name: 'Sheet1',
          rowCount: 2,
          columnCount: 3,
          nonEmptyCellCount: 4,
          previewRows: [
            ['1', '2', '=A1+B1'],
            ['3', '', ''],
          ],
        }),
      ]),
    )
  })

  it('resolves external workbook cell references from saved XLSX external-link caches', async () => {
    const imported = importXlsx(buildExternalLinkCacheWorkbook(), 'external-link-cache.xlsx')
    const formulaCell = imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')

    expect(formulaCell?.formula).toBe('2+3')
    expect(imported.warnings).toEqual([externalWorkbookReferencesWarning])
    expect(imported.snapshot.workbook.metadata?.externalWorkbookReferences).toEqual([
      {
        bookIndex: 1,
        packagePath: 'xl/externalLinks/externalLink5.xml',
        target: 'file:///tmp/source.xlsx',
        targetMode: 'External',
        workbookName: 'source.xlsx',
        sheetNames: ['External Data'],
      },
    ])
    expect(imported.snapshot.workbook.metadata?.unsupportedFormulaDependencies).toEqual([
      {
        kind: 'external-workbook-reference',
        sheetName: 'Report',
        address: 'A1',
        formula: "'[1]External Data'!A1+'[1]External Data'!A2",
        importedFormula: '2+3',
        linkedWorkbooks: [
          {
            bookIndex: 1,
            packagePath: 'xl/externalLinks/externalLink5.xml',
            target: 'file:///tmp/source.xlsx',
            targetMode: 'External',
            workbookName: 'source.xlsx',
            sheetNames: ['External Data'],
          },
        ],
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: true,
        cachedExternalReferenceValuesUsed: true,
        resolvedExternalReferenceCount: 2,
        unresolvedExternalReferenceCount: 0,
        reason:
          'Formula depends on an external workbook reference; cached linked values are preserved but linked workbooks are not recalculated during import.',
      },
    ])

    const engine = new SpreadsheetEngine({ workbookName: 'external-link-cache-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Report', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
  })

  it('preserves external GETPIVOTDATA anchors instead of replacing them with cached labels', () => {
    const imported = importXlsx(buildExternalGetPivotDataLinkCacheWorkbook(), 'external-pivot-link-cache.xlsx')
    const formulaCell = imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')

    expect(formulaCell?.formula).toBe('GETPIVOTDATA("Amount",\'[1]External Pivot\'!$G$3,"Region","East")')
    expect(formulaCell?.value).toBe(15)
    expect(imported.warnings).toContain(externalWorkbookReferencesWarning)
    expect(imported.snapshot.workbook.metadata?.externalWorkbookReferences).toEqual([
      {
        bookIndex: 1,
        packagePath: 'xl/externalLinks/externalLink5.xml',
        target: 'file:///tmp/pivot-source.xlsx',
        targetMode: 'External',
        workbookName: 'pivot-source.xlsx',
        sheetNames: ['External Pivot'],
      },
    ])
    expect(imported.snapshot.workbook.metadata?.unsupportedFormulaDependencies).toEqual([
      expect.objectContaining({
        kind: 'external-workbook-reference',
        sheetName: 'Report',
        address: 'A1',
        formula: 'GETPIVOTDATA("Amount",\'[1]External Pivot\'!$G$3,"Region","East")',
        importedFormula: 'GETPIVOTDATA("Amount",\'[1]External Pivot\'!$G$3,"Region","East")',
        linkedWorkbooks: [
          {
            bookIndex: 1,
            packagePath: 'xl/externalLinks/externalLink5.xml',
            target: 'file:///tmp/pivot-source.xlsx',
            targetMode: 'External',
            workbookName: 'pivot-source.xlsx',
            sheetNames: ['External Pivot'],
          },
        ],
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: true,
        cachedExternalReferenceValuesUsed: false,
        resolvedExternalReferenceCount: 0,
        unresolvedExternalReferenceCount: 0,
      }),
    ])
  })

  it('retains cached values for imported formula cells that use unavailable add-in functions', () => {
    const imported = importXlsx(buildUnsupportedFunctionCacheWorkbook(), 'udf-cache.xlsx')
    const sheet = imported.snapshot.sheets[0]

    expect(sheet?.cells.find((cell) => cell.address === 'A1')).toMatchObject({
      formula: '_xldudf_WISEPRICE(B1,"Shares Outstanding")',
      value: 14935800000,
    })
    expect(sheet?.cells.find((cell) => cell.address === 'C1')).toMatchObject({
      formula: '_FV(B1,"Ticker symbol",TRUE)',
      value: 'AAPL',
    })
  })

  it('drops degenerate single-cell merge records during import', async () => {
    const workbook = XLSX.utils.book_new()
    const sheet = XLSX.utils.aoa_to_sheet([['A', 'B']])
    sheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } },
      { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
    ]
    XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'single-cell-merge.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.merges).toEqual([{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' }])
    const engine = new SpreadsheetEngine({ workbookName: 'single-cell-merge-import' })
    await engine.ready()
    expect(() => engine.importSnapshot(imported.snapshot)).not.toThrow()
  })

  it('ignores zero-size row and column metadata during import', () => {
    const workbook = XLSX.utils.book_new()
    const sheet = XLSX.utils.aoa_to_sheet([['Value']])
    sheet['!rows'] = [{ hpx: 0 }]
    sheet['!cols'] = [{ wpx: 0 }]
    XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'zero-size.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.rows).toBeUndefined()
    expect(imported.snapshot.sheets[0]?.metadata?.columns).toBeUndefined()
  })

  it('preserves hidden row metadata even when the row has no custom height', () => {
    const workbook = XLSX.utils.book_new()
    const sheet = XLSX.utils.aoa_to_sheet([
      ['Header', 'Value'],
      ['Visible', 10],
      ['Hidden', 20],
    ])
    sheet['!rows'] = [undefined, undefined, { hidden: true }]
    XLSX.utils.book_append_sheet(workbook, sheet, 'Table')

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'hidden-row.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.rows).toEqual([{ id: 'row:2', index: 2, hidden: true }])
  })

  it('canonicalizes imported multiline text to LF line breaks', () => {
    const workbook = XLSX.utils.book_new()
    const sheet = XLSX.utils.aoa_to_sheet([['Line 1\r\nLine 2\rLine 3']])
    XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'multiline.xlsx')

    expect(imported.snapshot.sheets[0]?.cells).toContainEqual({
      address: 'A1',
      row: 0,
      col: 0,
      value: 'Line 1\nLine 2\nLine 3',
    })
  })

  it('preserves external workbook defined names as formulas across export round trips', () => {
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['local']]), 'Sheet1')
    workbook.Workbook = {
      Names: [
        { Name: 'ExternalRange', Ref: '[1]Sheet1!$A$1:$A$2' },
        { Name: 'ExternalBrokenRef', Ref: '[2]Sheet1!#REF!' },
      ],
    }

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'external-defined-names.xlsx')
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'ExternalBrokenRef', value: { kind: 'formula', formula: '=[2]Sheet1!#REF!' } },
      { name: 'ExternalRange', value: { kind: 'formula', formula: '=[1]Sheet1!$A$1:$A$2' } },
    ])

    const roundTripped = importXlsx(exportXlsx(imported.snapshot), 'external-defined-names.xlsx')
    expect(roundTripped.snapshot.workbook.metadata?.definedNames).toEqual(imported.snapshot.workbook.metadata?.definedNames)
  })

  it('preserves sheet-scoped defined names across import and export round trips', () => {
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[100]]), 'Global')
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[7, 10, { f: 'LocalBonus*LocalRevenue', v: 70 }]]), 'Local')
    workbook.Workbook = {
      Names: [
        { Name: 'LocalBonus', Ref: 'Global!$A$1' },
        { Name: 'LocalBonus', Sheet: 1, Ref: 'Local!$A$1' },
        { Name: 'LocalRevenue', Sheet: 1, Ref: 'Local!$B$1' },
      ],
    }

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'scoped-defined-names.xlsx')

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'LocalBonus', value: { kind: 'cell-ref', sheetName: 'Global', address: 'A1' } },
      { name: 'LocalBonus', scopeSheetName: 'Local', value: { kind: 'cell-ref', sheetName: 'Local', address: 'A1' } },
      { name: 'LocalRevenue', scopeSheetName: 'Local', value: { kind: 'cell-ref', sheetName: 'Local', address: 'B1' } },
    ])

    const roundTripped = importXlsx(exportXlsx(imported.snapshot), 'scoped-defined-names-roundtrip.xlsx')
    expect(roundTripped.warnings).toEqual([])
    expect(roundTripped.snapshot.workbook.metadata?.definedNames).toEqual(imported.snapshot.workbook.metadata?.definedNames)
  })

  it('preserves formula-only cells across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'formula-only-export' },
      sheets: [
        {
          id: 1,
          name: 'Summary',
          order: 0,
          cells: [
            { address: 'A1', value: 12 },
            { address: 'A2', value: 5 },
            { address: 'A3', formula: 'A1-A2' },
          ],
        },
      ],
    }

    const roundTripped = importXlsx(exportXlsx(snapshot), 'formula-only-export.xlsx')

    expect(roundTripped.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: 'A3', formula: 'A1-A2' })]),
    )
  })

  it('preserves sheet names with trailing spaces across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'Trailing Space Workbook' },
      sheets: [
        {
          id: 'sheet-trailing-space',
          name: 'Table 2.1.2  ',
          order: 0,
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Header' },
            { address: 'B1', row: 0, col: 1, value: 'Value' },
          ],
          metadata: {
            merges: [{ sheetName: 'Table 2.1.2  ', startAddress: 'A1', endAddress: 'B1' }],
          },
        },
      ],
    }

    const roundTripped = importXlsx(exportXlsx(snapshot), 'trailing-space-sheet.xlsx')
    expect(roundTripped.snapshot.sheets[0]?.name).toBe('Table 2.1.2  ')
    expect(roundTripped.snapshot.sheets[0]?.metadata?.merges).toEqual([
      { sheetName: 'Table 2.1.2  ', startAddress: 'A1', endAddress: 'B1' },
    ])
  })

  it('preserves leading-zero number formats across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      workbook: { name: 'leading-zero-number-format' },
      sheets: [
        {
          id: 1,
          name: 'Codes',
          order: 0,
          cells: [{ address: 'A1', value: 7, format: '00' }],
        },
      ],
    }

    const bytes = exportXlsx(snapshot)
    const zip = unzipSync(bytes)
    const stylesXml = strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())
    const roundTripped = importXlsx(bytes, 'leading-zero-number-format.xlsx')

    expect(stylesXml).toContain('formatCode="00"')
    expect(stylesXml).not.toContain('numFmtId="00"')
    expect(roundTripped.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', row: 0, col: 0, value: 7, format: '00' }])
  })

  it('preserves macro payloads without executing them across macro-enabled workbook import and export', () => {
    const imported = importXlsx(buildMacroEnabledWorkbook(), 'Macro Workbook.xlsm')

    expect(imported.workbookName).toBe('Macro Workbook')
    expect(imported.warnings).toContain('Macros were preserved but not executed during XLSX import.')
    expect(imported.snapshot.workbook.metadata?.macroPayloads).toEqual([
      {
        kind: 'vbaProject',
        storage: 'base64',
        dataBase64: 'AQIDBA==',
        byteLength: 4,
        preservedWithoutExecution: true,
        workbookCodeName: 'ThisWorkbook',
        sheetCodeNames: [{ sheetName: 'Sheet1', codeName: 'Sheet1' }],
      },
    ])
    expect(imported.snapshot.sheets[0]?.cells).toEqual([expect.objectContaining({ address: 'A1', value: 'safe value' })])

    const exported = exportXlsx(imported.snapshot)
    const roundTripped = XLSX.read(exported, { type: 'array', bookVBA: true })
    expect(Array.from(roundTripped.vbaraw ?? [])).toEqual([1, 2, 3, 4])
    expect(roundTripped.SheetNames).toEqual(['Sheet1'])
    expect(roundTripped.Workbook?.WBProps?.CodeName).toBe('ThisWorkbook')
    expect(roundTripped.Workbook?.Sheets?.[0]?.CodeName).toBe('Sheet1')
    expect(roundTripped.Sheets['Sheet1']?.['A1']?.v).toBe('safe value')
  })

  it('maps imported xlsx styles into Bilig style records', () => {
    expect(
      readImportedXlsxCellStyle({
        patternType: 'solid',
        fgColor: { rgb: '1D3989' },
        font: {
          name: 'Aptos',
          sz: 12,
          bold: true,
          italic: true,
          underline: true,
          color: { rgb: 'FFFFFFFF' },
        },
        alignment: {
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
          indent: 1,
        },
        border: {
          bottom: {
            style: 'thin',
            color: { rgb: 'FF000000' },
          },
        },
        protection: {
          locked: false,
          hidden: true,
        },
      }),
    ).toEqual({
      fill: { backgroundColor: '#1d3989' },
      font: {
        family: 'Aptos',
        size: 12,
        bold: true,
        italic: true,
        underline: true,
        color: '#ffffff',
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrap: true,
        indent: 1,
      },
      borders: {
        bottom: {
          style: 'solid',
          weight: 'thin',
          color: '#000000',
        },
      },
      protection: {
        locked: false,
        hidden: true,
      },
    })
  })

  it('coalesces repeated imported xlsx styles into rectangular ranges', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Styled import',
        metadata: {
          styles: [
            {
              id: 'header-fill',
              fill: { backgroundColor: '#1d3989' },
              font: { bold: true, color: '#ffffff' },
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Styled',
          order: 0,
          cells: [
            { address: 'A1', value: 'Header' },
            { address: 'B1', value: 'Header' },
            { address: 'C1', value: 'Header' },
            { address: 'A2', value: 'Header' },
            { address: 'B2', value: 'Header' },
            { address: 'C2', value: 'Header' },
            { address: 'A3', value: 'Header' },
            { address: 'B3', value: 'Header' },
            { address: 'C3', value: 'Header' },
          ],
          metadata: {
            styleRanges: [
              {
                range: { sheetName: 'Styled', startAddress: 'A1', endAddress: 'C3' },
                styleId: 'header-fill',
              },
            ],
          },
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'styled-block.xlsx')
    const styleRanges = imported.snapshot.sheets[0]?.metadata?.styleRanges ?? []
    const styleRange = styleRanges[0]

    expect(styleRanges).toHaveLength(1)
    expect(styleRange?.range).toEqual({ sheetName: 'Styled', startAddress: 'A1', endAddress: 'C3' })
    const style = imported.snapshot.workbook.metadata?.styles?.find((entry) => entry.id === styleRange?.styleId)
    expect(style).toMatchObject({
      fill: { backgroundColor: '#1d3989' },
      font: { bold: true, color: '#ffffff' },
    })
  })

  it('preserves cell-level protection style metadata across XLSX export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Cell protection styles',
        metadata: {
          styles: [
            {
              id: 'unlocked-input',
              fill: { backgroundColor: '#fff2cc' },
              protection: { locked: false },
            },
            {
              id: 'hidden-formula',
              font: { color: '#000000' },
              protection: { locked: true, hidden: true },
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Protected',
          order: 0,
          metadata: {
            sheetProtection: { sheetName: 'Protected' },
            styleRanges: [
              { range: { sheetName: 'Protected', startAddress: 'B2', endAddress: 'B3' }, styleId: 'unlocked-input' },
              { range: { sheetName: 'Protected', startAddress: 'C2', endAddress: 'C3' }, styleId: 'hidden-formula' },
            ],
          },
          cells: [
            { address: 'B2', value: 10 },
            { address: 'B3', value: 25 },
            { address: 'C2', formula: 'B2*2' },
            { address: 'C3', formula: 'B3*2' },
          ],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const zip = unzipSync(exported)
    const stylesXml = strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())
    const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const imported = importXlsx(exported, 'cell-protection-style-roundtrip.xlsx')
    const stylesById = new Map((imported.snapshot.workbook.metadata?.styles ?? []).map((style) => [style.id, style]))
    const styleForRange = (startAddress: string, endAddress: string) => {
      const styleRange = imported.snapshot.sheets[0]?.metadata?.styleRanges?.find(
        (entry) => entry.range.startAddress === startAddress && entry.range.endAddress === endAddress,
      )
      return styleRange ? stylesById.get(styleRange.styleId) : undefined
    }

    expect(stylesXml).toContain('applyProtection="1"')
    expect(stylesXml).toContain('<protection locked="0"/>')
    expect(stylesXml).toContain('<protection locked="1" hidden="1"/>')
    expect(sheetXml).toContain('<sheetProtection sheet="1"/>')
    expect(styleForRange('B2', 'B3')?.protection).toEqual({ locked: false })
    expect(styleForRange('C2', 'C3')?.protection).toEqual({ locked: true, hidden: true })
  })

  it('imports multiple generic workbook shapes without file-specific dispatch', () => {
    const operations = importXlsx(buildGenericWorkflowWorkbookFixture('multi-sheet-operations'), 'operations-workflow.xlsx')
    expect(operations.sheetNames).toEqual(['Dashboard', 'Ledger', 'Rollforward', 'Lookups'])
    const ledger = operations.snapshot.sheets.find((sheet) => sheet.name === 'Ledger')
    expect(ledger).toMatchObject({
      name: 'Ledger',
      metadata: {
        columns: expect.arrayContaining([{ id: 'col:0', index: 0, size: 132 }]),
        rows: expect.arrayContaining([{ id: 'row:0', index: 0, size: 30 }]),
        merges: [{ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'H1' }],
      },
    })
    expect(ledger?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A4', value: 'OP001' }),
        expect.objectContaining({ address: 'B4', value: 45292 }),
        expect.objectContaining({
          address: 'G4',
          formula: 'F4-SUMIF(Rollforward!$B:$B,A4,Rollforward!$E:$E)',
        }),
      ]),
    )
    const rollforward = operations.snapshot.sheets.find((sheet) => sheet.name === 'Rollforward')
    expect(rollforward?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'E5', formula: 'IF(B5=B4,E4+D5,D5)' })]))

    const planning = importXlsx(buildGenericWorkflowWorkbookFixture('single-sheet-planning'), 'monthly-plan.xlsx')
    expect(planning.sheetNames).toEqual(['Monthly Plan'])
    expect(planning.snapshot.sheets[0]).toMatchObject({
      name: 'Monthly Plan',
      metadata: {
        columns: expect.arrayContaining([{ id: 'col:0', index: 0, size: 168 }]),
        rows: expect.arrayContaining([{ id: 'row:0', index: 0, size: 30 }]),
        merges: [{ sheetName: 'Monthly Plan', startAddress: 'A1', endAddress: 'I1' }],
      },
    })
    expect(planning.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A3', value: 'TenantWorks' }),
        expect.objectContaining({
          address: 'F3',
          formula: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)',
        }),
      ]),
    )
  })

  it('imports csv files into a single-sheet workbook preview', async () => {
    const imported = importCsv('Name,Value\nalpha,12\nbeta,=A2', 'metrics.csv')

    expect(imported.workbookName).toBe('metrics')
    expect(imported.sheetNames).toEqual(['metrics'])
    expect(imported.snapshot.sheets[0]).toMatchObject({
      name: 'metrics',
      cells: [
        { address: 'A1', value: 'Name' },
        { address: 'B1', value: 'Value' },
        { address: 'A2', value: 'alpha' },
        { address: 'B2', value: 12 },
        { address: 'A3', value: 'beta' },
        { address: 'B3', formula: 'A2' },
      ],
    })
    expect(readRuntimeImage(imported.snapshot)?.sheetCells).toEqual([
      {
        sheetName: 'metrics',
        coords: [
          { row: 0, col: 0 },
          { row: 0, col: 1 },
          { row: 1, col: 0 },
          { row: 1, col: 1 },
          { row: 2, col: 0 },
          { row: 2, col: 1 },
        ],
        coordinateOrder: 'dense-row-major',
        dimensions: { width: 2, height: 3 },
        cellCount: 6,
      },
    ])
    const engine = new SpreadsheetEngine({ workbookName: imported.workbookName, replicaId: 'csv-formula-runtime-image-restore' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    expect(engine.getCellValue('metrics', 'B3')).toEqual({ tag: ValueTag.String, value: 'alpha', stringId: expect.any(Number) })
    expect(imported.preview).toMatchObject({
      workbookName: 'metrics',
      sheetCount: 1,
      sheets: [
        {
          name: 'metrics',
          rowCount: 3,
          columnCount: 2,
          nonEmptyCellCount: 6,
          previewRows: [
            ['Name', 'Value'],
            ['alpha', '12'],
            ['beta', '=A2'],
          ],
        },
      ],
    })
  })

  it('attaches runtime coordinates for literal-only dense csv imports', () => {
    const imported = importCsv('Name,Value\nalpha,12\nbeta,24', 'literal-metrics.csv')
    const runtimeImage = readRuntimeImage(imported.snapshot)

    expect(runtimeImage?.sheetCells).toEqual([
      {
        sheetName: 'literal-metrics',
        coords: [
          { row: 0, col: 0 },
          { row: 0, col: 1 },
          { row: 1, col: 0 },
          { row: 1, col: 1 },
          { row: 2, col: 0 },
          { row: 2, col: 1 },
        ],
        coordinateOrder: 'dense-row-major',
        dimensions: { width: 2, height: 3 },
        cellCount: 6,
      },
    ])
  })

  it('parses common accounting number formats from csv imports', () => {
    const imported = importCsv(
      'Account,Amount,Margin,Variance\nRevenue,"$1,234.56",12.5%,"$1,234.56"\nCOGS,"($987.65)",-3.25%,"(987.65)"',
      'accounting.csv',
    )

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'B2', row: 1, col: 1, value: 1234.56 }),
        expect.objectContaining({ address: 'C2', row: 1, col: 2, value: 0.125 }),
        expect.objectContaining({ address: 'D2', row: 1, col: 3, value: 1234.56 }),
        expect.objectContaining({ address: 'B3', row: 2, col: 1, value: -987.65 }),
        expect.objectContaining({ address: 'C3', row: 2, col: 2, value: -0.0325 }),
        expect.objectContaining({ address: 'D3', row: 2, col: 3, value: -987.65 }),
      ]),
    )
  })

  it('imports semicolon-delimited accounting csv files with decimal commas', () => {
    const imported = importCsv('Account;Amount;Tax\n4000;125,50;20,08\n5000;-12,25;0,00', 'locale-accounting.csv')

    expect(imported.warnings).toEqual([])
    expect(imported.preview.sheets[0]?.previewRows).toEqual([
      ['Account', 'Amount', 'Tax'],
      ['4000', '125,50', '20,08'],
      ['5000', '-12,25', '0,00'],
    ])
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A1', row: 0, col: 0, value: 'Account' }),
        expect.objectContaining({ address: 'B1', row: 0, col: 1, value: 'Amount' }),
        expect.objectContaining({ address: 'C1', row: 0, col: 2, value: 'Tax' }),
        expect.objectContaining({ address: 'A2', row: 1, col: 0, value: '4000' }),
        expect.objectContaining({ address: 'B2', row: 1, col: 1, value: 125.5 }),
        expect.objectContaining({ address: 'C2', row: 1, col: 2, value: 20.08 }),
        expect.objectContaining({ address: 'A3', row: 2, col: 0, value: '5000' }),
        expect.objectContaining({ address: 'B3', row: 2, col: 1, value: -12.25 }),
        expect.objectContaining({ address: 'C3', row: 2, col: 2, value: 0 }),
      ]),
    )
  })

  it('dispatches workbook import by content type', () => {
    const imported = importWorkbookFile(new TextEncoder().encode('A,B\n1,2'), 'dispatch.csv', CSV_CONTENT_TYPE)

    expect(imported.workbookName).toBe('dispatch')
    expect(imported.sheetNames).toEqual(['dispatch'])
  })

  it('normalizes workbook import content type parameters and case before dispatching', () => {
    const csvBytes = new TextEncoder().encode('A,B\n1,2')
    const csvVariants = [' text/csv; charset=utf-8 ', 'TEXT/CSV']
    for (const contentType of csvVariants) {
      const imported = importWorkbookFile(csvBytes, 'dispatch.csv', contentType)
      expect(imported.preview.contentType).toBe(CSV_CONTENT_TYPE)
      expect(imported.sheetNames).toEqual(['dispatch'])
    }

    const xlsxBytes = buildWorkbook()
    const xlsxVariants = [`${XLSX_CONTENT_TYPE}; charset=binary`, XLSX_CONTENT_TYPE.toUpperCase()]
    for (const contentType of xlsxVariants) {
      const imported = importWorkbookFile(xlsxBytes, 'dispatch.xlsx', contentType)
      expect(imported.preview.contentType).toBe(XLSX_CONTENT_TYPE)
      expect(imported.sheetNames).toEqual(['Sheet1', 'Sheet2'])
    }
  })

  it('dispatches binary Excel workbooks by XLSB content type', () => {
    const imported = importWorkbookFile(buildBinaryWorkbook(), 'dispatch.xlsb', 'application/vnd.ms-excel.sheet.binary.macroEnabled.12')

    expect(imported.preview.contentType).toBe(XLSB_CONTENT_TYPE)
    expect(imported.workbookName).toBe('dispatch')
    expect(imported.sheetNames).toEqual(['Sheet1', 'Sheet2'])
  })

  it('dispatches legacy Excel workbooks by XLS content type', () => {
    expect(EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES).toContain(LEGACY_XLS_CONTENT_TYPE)
    expect(WORKBOOK_IMPORT_CONTENT_TYPES).toContain(LEGACY_XLS_CONTENT_TYPE)

    const imported = importWorkbookFile(buildLegacyWorkbook(), 'legacy-salary.xls', 'application/vnd.ms-excel; charset=binary')

    expect(imported.preview.contentType).toBe(LEGACY_XLS_CONTENT_TYPE)
    expect(imported.workbookName).toBe('legacy-salary')
    expect(imported.sheetNames).toEqual(['Salary'])
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A2', value: 'Operations' }),
        expect.objectContaining({ address: 'B3', value: 1800 }),
        expect.objectContaining({ address: 'C2', value: 3050 }),
      ]),
    )
  })

  it('imports namespaced spreadsheet formulas as executable formulas', async () => {
    const imported = importWorkbookFile(buildNamespacedFormulaWorkbook(), 'legacy-expenses.xls', 'application/vnd.ms-excel')
    const importedCells = imported.snapshot.sheets[0]?.cells

    expect(imported.preview.contentType).toBe(LEGACY_XLS_CONTENT_TYPE)
    expect(imported.preview.sheets[0]?.previewRows[2]).toEqual(['=SUM(A1:A2)', '=SUM(A1:A2)'])
    expect(importedCells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A3', formula: 'SUM(A1:A2)', value: 3 }),
        expect.objectContaining({ address: 'B3', formula: 'SUM(A1:A2)', value: 3 }),
      ]),
    )

    const engine = new SpreadsheetEngine({ workbookName: 'namespaced-formula-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Expenses', 'A3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Expenses', 'B3')).toEqual({ tag: ValueTag.Number, value: 3 })
  })

  it('dispatches macro-enabled Excel workbooks by standard XLSM content type', () => {
    expect(EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES).toContain(XLSM_CONTENT_TYPE)
    expect(WORKBOOK_IMPORT_CONTENT_TYPES).toContain(XLSM_CONTENT_TYPE)

    const imported = importWorkbookFile(
      buildMacroEnabledWorkbook(),
      'Macro Workbook.xlsm',
      'application/vnd.ms-excel.sheet.macroEnabled.12; charset=binary',
    )

    expect(imported.preview.contentType).toBe(XLSM_CONTENT_TYPE)
    expect(imported.workbookName).toBe('Macro Workbook')
    expect(imported.warnings).toContain('Macros were preserved but not executed during XLSX import.')
    expect(imported.snapshot.workbook.metadata?.macroPayloads).toEqual([
      expect.objectContaining({
        kind: 'vbaProject',
        dataBase64: 'AQIDBA==',
        byteLength: 4,
        preservedWithoutExecution: true,
      }),
    ])
  })

  it('rejects corrupt zip-backed xlsx packages before parsing', () => {
    const bytes = buildCorruptZipBackedWorkbook()

    expect(() => importXlsx(bytes, 'corrupt.xlsx')).toThrow(InvalidXlsxZipContainerError)
    expect(() => importXlsx(bytes, 'corrupt.xlsx')).toThrow('Invalid or corrupt XLSX zip container')
    expect(() => importWorkbookFile(bytes, 'corrupt.xlsx', XLSX_CONTENT_TYPE)).toThrow(InvalidXlsxZipContainerError)
  })

  it('bounds whole-column defined names to the imported sheet extent', () => {
    const workbook = XLSX.utils.book_new()
    const sheet = XLSX.utils.aoa_to_sheet([
      ['Symbol', 'Year', 'Revenue'],
      ['AAA', 2020, 100],
      ['BBB', 2021, 200],
    ])
    XLSX.utils.book_append_sheet(workbook, sheet, 'Projectdata_NYSE')
    workbook.Workbook = {
      Names: [
        { Name: 'Symbol', Ref: 'Projectdata_NYSE!$A:$A' },
        { Name: 'Year_num', Ref: 'Projectdata_NYSE!$B:$B' },
        { Name: 'Total_Revenue', Ref: 'Projectdata_NYSE!$C:$C' },
      ],
    }

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'nyse.xlsx')

    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'Symbol', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'A1', endAddress: 'A3' } },
      { name: 'Total_Revenue', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'C1', endAddress: 'C3' } },
      { name: 'Year_num', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'B1', endAddress: 'B3' } },
    ])
  })

  it('exports workbook snapshots to XLSX bytes that import back with supported workbook semantics', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Roundtrip Workbook',
        metadata: {
          calculationSettings: { mode: 'manual', compatibilityMode: 'excel-modern' },
          properties: [
            { key: 'locale', value: 'en-US' },
            { key: 'reviewed', value: true },
            { key: 'threshold', value: 0.085 },
          ],
          definedNames: [
            { name: 'SummaryTotal', value: { kind: 'cell-ref', sheetName: 'Summary', address: 'B1' } },
            { name: 'InputRegion', value: { kind: 'range-ref', sheetName: 'Inputs', startAddress: 'A1', endAddress: 'B1' } },
            { name: 'TaxRate', value: { kind: 'scalar', value: 0.085 } },
          ],
          styles: [
            {
              id: 'accent-total',
              fill: { backgroundColor: '#1d3989' },
              font: { family: 'Aptos', size: 12, bold: true, color: '#ffffff' },
              alignment: { horizontal: 'center', vertical: 'middle', wrap: true },
              borders: { bottom: { style: 'solid', weight: 'thin', color: '#000000' } },
            },
          ],
          tables: [
            {
              name: 'Input.Table',
              sheetName: 'Inputs',
              startAddress: 'A1',
              endAddress: 'D4',
              columnNames: ['Region', 'Product', 'Sales', 'Notes'],
              headerRow: true,
              totalsRow: false,
            },
          ],
          charts: [
            {
              id: 'summary-trend',
              sheetName: 'Summary',
              address: 'E1',
              source: { sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' },
              chartType: 'line',
              seriesOrientation: 'columns',
              firstRowAsHeaders: true,
              firstColumnAsLabels: true,
              title: 'Summary Trend',
              legendPosition: 'right',
              rows: 12,
              cols: 6,
            },
          ],
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Summary',
              address: 'E15',
              source: { sheetName: 'Inputs', startAddress: 'A1', endAddress: 'D4' },
              groupBy: ['Region'],
              values: [
                { sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Total Sales' },
                { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
              ],
              rows: 4,
              cols: 3,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Summary',
          order: 0,
          metadata: {
            styleRanges: [{ range: { sheetName: 'Summary', startAddress: 'B1', endAddress: 'B1' }, styleId: 'accent-total' }],
            commentThreads: [
              {
                threadId: 'summary-total-note',
                sheetName: 'Summary',
                address: 'B1',
                comments: [{ id: 'summary-total-note-1', body: 'Reviewed total', authorDisplayName: 'Finance' }],
              },
            ],
            columns: [
              { id: 'summary-col-0', index: 0, size: 132 },
              { id: 'summary-col-1', index: 1, size: 96 },
            ],
            rows: [
              { id: 'summary-row-0', index: 0, size: 30 },
              { id: 'summary-row-2', index: 2, size: 24 },
            ],
            freezePane: { rows: 1, cols: 2 },
            merges: [{ sheetName: 'Summary', startAddress: 'A5', endAddress: 'B5' }],
            sheetProtection: { sheetName: 'Summary' },
            protectedRanges: [
              {
                id: 'protect-summary-inputs',
                range: { sheetName: 'Summary', startAddress: 'A2', endAddress: 'B3' },
              },
            ],
            filters: [{ sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' }],
            sorts: [
              {
                range: { sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' },
                keys: [{ keyAddress: 'B1', direction: 'desc' }],
              },
            ],
            validations: [
              {
                range: { sheetName: 'Summary', startAddress: 'C2', endAddress: 'C4' },
                rule: { kind: 'whole', operator: 'between', values: [0, 100] },
                allowBlank: false,
                errorStyle: 'stop',
                errorTitle: 'Percent required',
                errorMessage: 'Enter a whole number from 0 to 100.',
              },
            ],
            conditionalFormats: [
              {
                id: 'summary-high-total',
                range: { sheetName: 'Summary', startAddress: 'B2', endAddress: 'B3' },
                rule: { kind: 'cellIs', operator: 'greaterThan', values: [1000] },
                style: { fill: { backgroundColor: '#f4cccc' }, font: { bold: true, color: '#990000' } },
                stopIfTrue: true,
                priority: 1,
              },
            ],
          },
          cells: [
            { address: 'A1', value: 'Metric' },
            { address: 'B1', formula: 'SUM(B2:B3)', format: '0.00' },
            { address: 'C1', value: true },
            { address: 'A2', value: 'Revenue' },
            { address: 'B2', value: 1250.5, format: '$#,##0.00' },
            { address: 'A3', value: 'Costs' },
            { address: 'B3', value: 450.25, format: '$#,##0.00' },
          ],
        },
        {
          id: 2,
          name: 'Inputs',
          order: 1,
          metadata: {
            validations: [
              {
                range: { sheetName: 'Inputs', startAddress: 'D2', endAddress: 'D4' },
                rule: { kind: 'list', values: ['Priority', 'Standard'] },
                allowBlank: true,
                showDropdown: true,
                promptTitle: 'Status',
                promptMessage: 'Pick a known priority.',
                errorStyle: 'warning',
                errorTitle: 'Unknown priority',
                errorMessage: 'Use Priority or Standard.',
              },
            ],
          },
          cells: [
            { address: 'A1', value: 'Region' },
            { address: 'B1', value: 'Product' },
            { address: 'C1', value: 'Sales' },
            { address: 'D1', value: 'Notes' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 'Widget' },
            { address: 'C2', value: 10 },
            { address: 'D2', value: 'Priority' },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 'Widget' },
            { address: 'C3', value: 7 },
            { address: 'D3', value: 'Priority' },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 'Gizmo' },
            { address: 'C4', value: 5 },
            { address: 'D4', value: 'Standard' },
          ],
        },
      ],
    }

    const bytes = exportXlsx(snapshot)
    const imported = importXlsx(bytes, 'roundtrip.xlsx')
    const zip = unzipSync(bytes)

    expect(bytes.byteLength).toBeGreaterThan(0)
    expect(Object.keys(zip)).toEqual(expect.arrayContaining(['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml']))
    expect(Object.keys(zip)).toEqual(
      expect.arrayContaining([
        'xl/pivotTables/pivotTable1.xml',
        'xl/pivotCache/pivotCacheDefinition1.xml',
        'xl/pivotCache/pivotCacheRecords1.xml',
        'xl/tables/table1.xml',
      ]),
    )
    expect(strFromU8(zip['xl/charts/chart1.xml'] ?? new Uint8Array())).toContain('<c:lineChart>')
    expect(strFromU8(zip['xl/drawings/_rels/drawing1.xml.rels'] ?? new Uint8Array())).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
    )
    expect(strFromU8(zip['xl/pivotTables/pivotTable1.xml'] ?? new Uint8Array())).toContain('<pivotTableDefinition')
    expect(strFromU8(zip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array())).toContain(
      '<worksheetSource ref="A1:D4" sheet="Inputs"/>',
    )
    expect(strFromU8(zip['_rels/.rels'] ?? new Uint8Array())).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
    )
    expect(strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())).toContain(
      'application/vnd.openxmlformats-officedocument.custom-properties+xml',
    )
    expect(strFromU8(zip['docProps/custom.xml'] ?? new Uint8Array())).toContain('<property ')
    expect(strFromU8(zip['docProps/custom.xml'] ?? new Uint8Array())).toContain('name="locale"><vt:lpwstr>en-US</vt:lpwstr>')
    expect(strFromU8(zip['docProps/custom.xml'] ?? new Uint8Array())).toContain('name="reviewed"><vt:bool>true</vt:bool>')
    expect(strFromU8(zip['docProps/custom.xml'] ?? new Uint8Array())).toContain('name="threshold"><vt:r8>0.085</vt:r8>')
    expect(strFromU8(zip['xl/tables/table1.xml'] ?? new Uint8Array())).toContain('<table ')
    expect(strFromU8(zip['xl/tables/table1.xml'] ?? new Uint8Array())).toContain('displayName="Input.Table"')
    expect(strFromU8(zip['xl/tables/table1.xml'] ?? new Uint8Array())).toContain('<tableColumn id="3" name="Sales"/>')
    expect(strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())).toContain('<calcPr calcMode="manual"/>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<dataValidations count="1">')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain(
      '<dataValidation type="whole" operator="between" allowBlank="0" errorStyle="stop"',
    )
    expect(strFromU8(zip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())).toContain(
      '<dataValidation type="list" allowBlank="1" showDropDown="0"',
    )
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<conditionalFormatting sqref="B2:B3">')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain(
      '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan" stopIfTrue="1"><formula>1000</formula></cfRule>',
    )
    expect(strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())).toContain('<dxfs count="1">')
    expect(strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())).toContain('<fgColor rgb="FFF4CCCC"/>')
    expect(strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())).toContain('<color rgb="FF990000"/>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain(
      '<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>',
    )
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<sheetProtection sheet="1"/>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<protectedRanges>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain(
      '<protectedRange name="protect-summary-inputs" sqref="A2:B3"/>',
    )
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<autoFilter ref="A1:B3"/>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<sortState ref="A1:B3">')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<sortCondition descending="1" ref="B1:B3"/>')
    expect(projectSupportedSnapshotSemantics(imported.snapshot)).toEqual(projectSupportedSnapshotSemantics(snapshot))
  })

  it('preserves imported frozen pane scroll targets on export', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'frozen-pane-scroll-target' },
      sheets: [
        {
          id: 1,
          name: 'View',
          order: 0,
          metadata: {
            freezePane: { rows: 3, cols: 2, topLeftCell: 'I32', activePane: 'bottomRight' },
          },
          cells: [
            { address: 'A1', value: 'Account' },
            { address: 'B1', value: 'Amount' },
            { address: 'I32', value: 'scroll target' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'frozen-pane-scroll-target.xlsx')
    const freezePane = imported.snapshot.sheets[0]?.metadata?.freezePane
    const exportedZip = unzipSync(exportXlsx(imported.snapshot))
    const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(freezePane).toEqual({ rows: 3, cols: 2, topLeftCell: 'I32', activePane: 'bottomRight' })
    expect(sheetXml).toContain('<pane xSplit="2" ySplit="3" topLeftCell="I32" activePane="bottomRight" state="frozen"/>')
  })

  it('preserves unicode table names across XLSX export', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'unicode-table-name',
        metadata: {
          tables: [
            {
              name: 'Données_FR',
              sheetName: 'Données - FR',
              startAddress: 'A1',
              endAddress: 'B2',
              columnNames: ['MESURE', 'CATÉGORIE'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Données - FR',
          order: 0,
          cells: [
            { address: 'A1', value: 'MESURE' },
            { address: 'B1', value: 'CATÉGORIE' },
            { address: 'A2', value: 'Taxes' },
            { address: 'B2', value: 'Fédéral' },
          ],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const exportedZip = unzipSync(exported)
    const imported = importXlsx(exported, 'unicode-table-name.xlsx')

    expect(strFromU8(exportedZip['xl/tables/table1.xml'] ?? new Uint8Array())).toContain('displayName="Données_FR"')
    expect(imported.snapshot.workbook.metadata?.tables?.[0]?.name).toBe('Données_FR')
  })

  it('preserves worksheet tab colors on import and export', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'tab-color-roundtrip' },
      sheets: [
        {
          id: 1,
          name: 'Rgb',
          order: 0,
          metadata: { tabColor: { rgb: 'FF0070C0' } },
          cells: [{ address: 'A1', value: 'rgb tab' }],
        },
        {
          id: 2,
          name: 'Theme',
          order: 1,
          metadata: { tabColor: { theme: '8' } },
          cells: [{ address: 'A1', value: 'theme tab' }],
        },
        {
          id: 3,
          name: 'Tint',
          order: 2,
          metadata: { tabColor: { theme: '0', tint: '-0.14999847407452621' } },
          cells: [{ address: 'A1', value: 'tint tab' }],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const imported = importXlsx(exported, 'tab-color-roundtrip.xlsx')
    const reexportedZip = unzipSync(exportXlsx(imported.snapshot))

    expect(imported.snapshot.sheets.map((sheet) => sheet.metadata?.tabColor)).toEqual([
      { rgb: 'FF0070C0' },
      { theme: '8' },
      { theme: '0', tint: '-0.14999847407452621' },
    ])
    expect(strFromU8(reexportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<tabColor rgb="FF0070C0"/>')
    expect(strFromU8(reexportedZip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())).toContain('<tabColor theme="8"/>')
    expect(strFromU8(reexportedZip['xl/worksheets/sheet3.xml'] ?? new Uint8Array())).toContain(
      '<tabColor theme="0" tint="-0.14999847407452621"/>',
    )
  })

  it('exports custom number formats on populated and blank cells', () => {
    const snapshot: WorkbookSnapshot = {
      workbook: { id: 'custom-number-format-workbook', name: 'custom-number-format-workbook' },
      sheets: [
        {
          id: 1,
          name: 'Formats',
          order: 0,
          cells: [
            { address: 'A1', value: 0, format: '00' },
            { address: 'A2', format: '00' },
            { address: 'B1', value: 12.34, format: '"$"#,##0.00' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'custom-number-format-workbook.xlsx')

    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A1', row: 0, col: 0, value: 0, format: '00' }),
        expect.objectContaining({ address: 'A2', row: 1, col: 0, format: '00' }),
        expect.objectContaining({ address: 'B1', row: 0, col: 1, value: 12.34, format: '"$"#,##0.00' }),
      ]),
    )
  })

  it('roundtrips worksheet AutoFilter criteria and custom filters', () => {
    const sourceSnapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'autofilter-criteria-source' },
      sheets: [
        {
          name: 'Ledger',
          order: 0,
          metadata: {
            filters: [{ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'D6' }],
          },
          cells: [
            { address: 'A1', value: 'Date' },
            { address: 'B1', value: 'Department' },
            { address: 'C1', value: 'Amount' },
            { address: 'D1', value: 'Status' },
            { address: 'A2', value: '2026-01-01' },
            { address: 'B2', value: 'Finance' },
            { address: 'C2', value: -100 },
            { address: 'D2', value: 'Approved' },
            { address: 'A3', value: '2026-01-02' },
            { address: 'B3', value: 'Operations' },
            { address: 'C3', value: 75 },
            { address: 'D3', value: 'Pending' },
            { address: 'A4', value: '2026-01-03' },
            { address: 'B4', value: 'Finance' },
            { address: 'C4', value: -25 },
            { address: 'D4', value: 'Approved' },
            { address: 'A5', value: '2026-01-04' },
            { address: 'B5', value: 'Sales' },
            { address: 'C5', value: 50 },
            { address: 'D5', value: 'Rejected' },
            { address: 'A6', value: '2026-01-05' },
            { address: 'B6', value: 'Finance' },
            { address: 'C6', value: -1 },
            { address: 'D6', value: 'Approved' },
          ],
        },
      ],
    }
    const sourceZip = unzipSync(exportXlsx(sourceSnapshot))
    const sheetPath = 'xl/worksheets/sheet1.xml'
    const sourceSheetXml = strFromU8(sourceZip[sheetPath] ?? new Uint8Array())
    sourceZip[sheetPath] = strToU8(
      sourceSheetXml.replace(
        '<autoFilter ref="A1:D6"/>',
        [
          '<autoFilter ref="A1:D6">',
          '<filterColumn colId="1"><filters blank="0"><filter val="Finance"/></filters></filterColumn>',
          '<filterColumn colId="2"><customFilters><customFilter operator="lessThan" val="0"/></customFilters></filterColumn>',
          '<filterColumn colId="3"><filters blank="0"><filter val="Approved"/></filters></filterColumn>',
          '</autoFilter>',
        ].join(''),
      ),
    )

    const imported = importXlsx(zipSync(sourceZip), 'autofilter-criteria-source.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.filters).toEqual([
      {
        sheetName: 'Ledger',
        startAddress: 'A1',
        endAddress: 'D6',
        criteria: [
          { colId: 1, filters: { blank: false, values: ['Finance'] } },
          { colId: 2, customFilters: { filters: [{ operator: 'lessThan', value: '0' }] } },
          { colId: 3, filters: { blank: false, values: ['Approved'] } },
        ],
      },
    ])

    const exportedZip = unzipSync(exportXlsx(imported.snapshot))
    const exportedSheetXml = strFromU8(exportedZip[sheetPath] ?? new Uint8Array())

    expect(exportedSheetXml).toContain('<autoFilter ref="A1:D6">')
    expect(exportedSheetXml).toContain('<filterColumn colId="1"><filters blank="0"><filter val="Finance"/></filters></filterColumn>')
    expect(exportedSheetXml).toContain(
      '<filterColumn colId="2"><customFilters><customFilter operator="lessThan" val="0"/></customFilters></filterColumn>',
    )
    expect(exportedSheetXml).toContain('<filterColumn colId="3"><filters blank="0"><filter val="Approved"/></filters></filterColumn>')
  })

  it('exports range-only number formats on populated and blank cells', () => {
    const snapshot: WorkbookSnapshot = {
      workbook: {
        id: 'range-number-format-workbook',
        name: 'range-number-format-workbook',
        metadata: {
          formats: [{ id: 'format-zip-code', code: '00000', kind: 'number' }],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Formats',
          order: 0,
          metadata: {
            formatRanges: [
              {
                range: { sheetName: 'Formats', startAddress: 'B2', endAddress: 'C3' },
                formatId: 'format-zip-code',
              },
            ],
          },
          cells: [{ address: 'B2', value: 7 }],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'range-number-format-workbook.xlsx')

    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'B2', row: 1, col: 1, value: 7, format: '00000' }),
        expect.objectContaining({ address: 'C3', row: 2, col: 2, format: '00000' }),
      ]),
    )
  })
})
