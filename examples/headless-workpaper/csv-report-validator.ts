import { WorkPaper } from '@bilig/headless'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type ReportCsvRow = {
  orderId: string
  channel: string
  gross: string
  refund: string
  reportedNet: string
}

const csvInput = `
order_id,channel,gross,refund,reported_net
ORD-1001,Marketplace,1200,50,1150
ORD-1002,Direct,800,0,800
ORD-1003,Partner,450,25,425
ORD-1004,Direct,1500,200,1300
`.trim()

const output = buildCsvReportValidation(csvInput)
assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function buildCsvReportValidation(input: string) {
  const sourceRows = parseReportCsv(input)
  const reportRows = sourceRows.map((row, index) => {
    const spreadsheetRow = index + 2
    return [
      row.orderId,
      row.channel,
      readInputNumber(row.gross, `gross row ${spreadsheetRow}`),
      readInputNumber(row.refund, `refund row ${spreadsheetRow}`),
      readInputNumber(row.reportedNet, `reported net row ${spreadsheetRow}`),
      `=C${spreadsheetRow}-D${spreadsheetRow}`,
      `=IF(ABS(E${spreadsheetRow}-F${spreadsheetRow})<0.01,"OK","Mismatch")`,
    ]
  })

  const lastDataRow = sourceRows.length + 1
  const workbook = WorkPaper.buildFromSheets({
    Report: [['Order ID', 'Channel', 'Gross', 'Refund', 'Reported net', 'Expected net', 'Validation'], ...reportRows],
    Summary: [
      ['Metric', 'Value'],
      ['Rows', `=COUNTA(Report!A2:A${lastDataRow})`],
      ['Gross total', `=SUM(Report!C2:C${lastDataRow})`],
      ['Refund total', `=SUM(Report!D2:D${lastDataRow})`],
      ['Reported net total', `=SUM(Report!E2:E${lastDataRow})`],
      ['Expected net total', `=SUM(Report!F2:F${lastDataRow})`],
      ['Mismatch count', `=COUNTIF(Report!G2:G${lastDataRow},"Mismatch")`],
      ['Status', '=IF(B7=0,"valid","invalid")'],
    ],
  })

  const reportSheet = requireSheet(workbook, 'Report')
  const summarySheet = requireSheet(workbook, 'Summary')
  const status = readString(workbook, summarySheet, 7, 1, 'status')

  return {
    sourceRows: sourceRows.length,
    valid: status === 'valid',
    totals: {
      gross: readNumber(workbook, summarySheet, 2, 1, 'gross total'),
      refund: readNumber(workbook, summarySheet, 3, 1, 'refund total'),
      reportedNet: readNumber(workbook, summarySheet, 4, 1, 'reported net total'),
      expectedNet: readNumber(workbook, summarySheet, 5, 1, 'expected net total'),
    },
    mismatchCount: readNumber(workbook, summarySheet, 6, 1, 'mismatch count'),
    firstReportRow: workbook.getRangeSerialized({
      start: { sheet: reportSheet, row: 1, col: 0 },
      end: { sheet: reportSheet, row: 1, col: 6 },
    })[0],
    verified: true,
  }
}

function parseReportCsv(input: string): ReportCsvRow[] {
  const lines = input
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
  const headerLine = lines[0]
  const dataLines = lines.slice(1)
  if (headerLine === undefined) {
    throw new Error('expected CSV header row')
  }

  const headers = headerLine.split(',').map((header) => header.trim())
  const expectedHeaders = ['order_id', 'channel', 'gross', 'refund', 'reported_net']
  if (JSON.stringify(headers) !== JSON.stringify(expectedHeaders)) {
    throw new Error(`expected CSV headers ${expectedHeaders.join(',')}, received ${headers.join(',')}`)
  }

  return dataLines.map((line, index) => {
    const values = line.split(',').map((value) => value.trim())
    if (values.length !== expectedHeaders.length) {
      throw new Error(`expected ${expectedHeaders.length} CSV fields on data row ${index + 2}, received ${values.length}`)
    }

    const [orderId, channel, gross, refund, reportedNet] = values
    if (!orderId || !channel || gross === undefined || refund === undefined || reportedNet === undefined) {
      throw new Error(`missing CSV field on data row ${index + 2}`)
    }

    return { orderId, channel, gross, refund, reportedNet }
  })
}

function readInputNumber(value: string, label: string): number {
  const parsed = Number(value)
  if (!Number.isFinite(parsed)) {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(value)}`)
  }
  return parsed
}

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readNumber(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 10000) / 10000
}

function readString(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): string {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'string') {
    throw new Error(`Expected ${label} to be text, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function assertOutput(actual: ReturnType<typeof buildCsvReportValidation>): void {
  const expected = {
    sourceRows: 4,
    valid: true,
    totals: {
      gross: 3950,
      refund: 275,
      reportedNet: 3675,
      expectedNet: 3675,
    },
    mismatchCount: 0,
    firstReportRow: [
      'ORD-1001',
      'Marketplace',
      1200,
      50,
      1150,
      '=C2-D2',
      '=IF(ABS(E2-F2)<0.01,"OK","Mismatch")',
    ],
    verified: true,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected CSV report validation result: ${JSON.stringify(actual)}`)
  }
}
