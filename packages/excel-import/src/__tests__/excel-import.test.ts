import { describe, expect, it } from 'vitest'
import { strFromU8, unzipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type { WorkbookChartSnapshot, WorkbookPivotSnapshot, WorkbookPivotValueSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importCsv, importWorkbookFile, importXlsx, readImportedXlsxCellStyle } from '../index.js'
import { CSV_CONTENT_TYPE } from '@bilig/agent-api'

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

function buildMacroEnabledWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['safe value']]), 'Sheet1')
  workbook.vbaraw = new Uint8Array([1, 2, 3, 4])
  return XLSX.write(workbook, { bookType: 'xlsm', type: 'buffer', bookVBA: true })
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

  it('warns and ignores macro payloads from macro-enabled workbook bytes', () => {
    const imported = importXlsx(buildMacroEnabledWorkbook(), 'Macro Workbook.xlsm')

    expect(imported.workbookName).toBe('Macro Workbook')
    expect(imported.warnings).toContain('Macros were ignored during XLSX import.')
    expect(imported.snapshot.sheets[0]?.cells).toEqual([expect.objectContaining({ address: 'A1', value: 'safe value' })])
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
    })
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

  it('imports csv files into a single-sheet workbook preview', () => {
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

  it('dispatches workbook import by content type', () => {
    const imported = importWorkbookFile(new TextEncoder().encode('A,B\n1,2'), 'dispatch.csv', CSV_CONTENT_TYPE)

    expect(imported.workbookName).toBe('dispatch')
    expect(imported.sheetNames).toEqual(['dispatch'])
  })

  it('exports workbook snapshots to XLSX bytes that import back with supported workbook semantics', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Roundtrip Workbook',
        metadata: {
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
            merges: [{ sheetName: 'Summary', startAddress: 'A5', endAddress: 'B5' }],
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
    expect(projectSupportedSnapshotSemantics(imported.snapshot)).toEqual(projectSupportedSnapshotSemantics(snapshot))
  })
})

interface ProjectedChartSemantics {
  id: string
  sheetName: string
  address: string
  source: WorkbookChartSnapshot['source']
  chartType: WorkbookChartSnapshot['chartType']
  seriesOrientation?: WorkbookChartSnapshot['seriesOrientation']
  firstRowAsHeaders?: WorkbookChartSnapshot['firstRowAsHeaders']
  firstColumnAsLabels?: WorkbookChartSnapshot['firstColumnAsLabels']
  title?: WorkbookChartSnapshot['title']
  legendPosition?: WorkbookChartSnapshot['legendPosition']
  rows: number
  cols: number
}

function projectChartSemantics(chart: WorkbookChartSnapshot): ProjectedChartSemantics {
  const projected: ProjectedChartSemantics = {
    id: chart.id,
    sheetName: chart.sheetName,
    address: chart.address,
    source: chart.source,
    chartType: chart.chartType,
    rows: chart.rows,
    cols: chart.cols,
  }
  if (chart.seriesOrientation !== undefined) {
    projected.seriesOrientation = chart.seriesOrientation
  }
  if (chart.firstRowAsHeaders !== undefined) {
    projected.firstRowAsHeaders = chart.firstRowAsHeaders
  }
  if (chart.firstColumnAsLabels !== undefined) {
    projected.firstColumnAsLabels = chart.firstColumnAsLabels
  }
  if (chart.title !== undefined) {
    projected.title = chart.title
  }
  if (chart.legendPosition !== undefined) {
    projected.legendPosition = chart.legendPosition
  }
  return projected
}

function projectPivotValue(value: WorkbookPivotValueSnapshot): WorkbookPivotValueSnapshot {
  const projected: WorkbookPivotValueSnapshot = {
    sourceColumn: value.sourceColumn,
    summarizeBy: value.summarizeBy,
  }
  if (value.outputLabel !== undefined) {
    projected.outputLabel = value.outputLabel
  }
  return projected
}

function projectPivotSemantics(pivot: WorkbookPivotSnapshot): WorkbookPivotSnapshot {
  return {
    name: pivot.name,
    sheetName: pivot.sheetName,
    address: pivot.address,
    source: pivot.source,
    groupBy: [...pivot.groupBy],
    values: pivot.values.map(projectPivotValue),
    rows: pivot.rows,
    cols: pivot.cols,
  }
}

function projectSupportedSnapshotSemantics(snapshot: WorkbookSnapshot) {
  const stylesById = new Map((snapshot.workbook.metadata?.styles ?? []).map((style) => [style.id, style]))
  const portableStyle = (styleId: string) => {
    const style = stylesById.get(styleId)
    if (!style) {
      return undefined
    }
    return {
      ...(style.fill ? { fill: style.fill } : {}),
      ...(style.font ? { font: style.font } : {}),
      ...(style.alignment ? { alignment: style.alignment } : {}),
      ...(style.borders ? { borders: style.borders } : {}),
    }
  }
  return {
    definedNames: (snapshot.workbook.metadata?.definedNames ?? [])
      .map((definedName) => ({ name: definedName.name, value: definedName.value }))
      .toSorted((left, right) => left.name.localeCompare(right.name)),
    charts: (snapshot.workbook.metadata?.charts ?? [])
      .map(projectChartSemantics)
      .toSorted((left, right) => left.id.localeCompare(right.id)),
    pivots: (snapshot.workbook.metadata?.pivots ?? [])
      .map(projectPivotSemantics)
      .toSorted((left, right) =>
        `${left.sheetName}:${left.address}:${left.name}`.localeCompare(`${right.sheetName}:${right.address}:${right.name}`),
      ),
    sheets: snapshot.sheets
      .toSorted((left, right) => left.order - right.order)
      .map((sheet) => ({
        name: sheet.name,
        order: sheet.order,
        cells: sheet.cells
          .map((cell) => ({
            address: cell.address,
            ...(cell.value !== undefined ? { value: cell.value } : {}),
            ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
            ...(cell.format !== undefined ? { format: cell.format } : {}),
          }))
          .toSorted((left, right) => left.address.localeCompare(right.address)),
        metadata: {
          columns: (sheet.metadata?.columns ?? [])
            .map(({ index, size }) => ({ index, size }))
            .toSorted((left, right) => left.index - right.index),
          rows: (sheet.metadata?.rows ?? [])
            .map(({ index, size }) => ({ index, size }))
            .toSorted((left, right) => left.index - right.index),
          merges: (sheet.metadata?.merges ?? [])
            .map(({ sheetName, startAddress, endAddress }) => ({ sheetName, startAddress, endAddress }))
            .toSorted((left, right) =>
              `${left.sheetName}:${left.startAddress}:${left.endAddress}`.localeCompare(
                `${right.sheetName}:${right.startAddress}:${right.endAddress}`,
              ),
            ),
          commentThreads: (sheet.metadata?.commentThreads ?? [])
            .map((thread) => ({
              sheetName: thread.sheetName,
              address: thread.address,
              comments: thread.comments.map((comment) => ({
                body: comment.body,
                ...(comment.authorDisplayName !== undefined ? { authorDisplayName: comment.authorDisplayName } : {}),
              })),
            }))
            .toSorted((left, right) => `${left.sheetName}:${left.address}`.localeCompare(`${right.sheetName}:${right.address}`)),
          styleRanges: (sheet.metadata?.styleRanges ?? [])
            .map((styleRange) => ({
              range: styleRange.range,
              style: portableStyle(styleRange.styleId),
            }))
            .toSorted((left, right) =>
              `${left.range.sheetName}:${left.range.startAddress}:${left.range.endAddress}`.localeCompare(
                `${right.range.sheetName}:${right.range.startAddress}:${right.range.endAddress}`,
              ),
            ),
        },
      })),
  }
}
