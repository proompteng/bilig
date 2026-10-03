import { WorkPaper } from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

const quoteItems = [
  {
    sku: 'PRO-ANNUAL',
    quantity: 12,
    listPrice: 240,
    discountPercent: 0.1,
  },
  {
    sku: 'SETUP',
    quantity: 1,
    listPrice: 1800,
    discountPercent: 0.25,
  },
  {
    sku: 'SUPPORT',
    quantity: 6,
    listPrice: 150,
    discountPercent: 0.05,
  },
  {
    sku: 'DATA-MIGRATION',
    quantity: 2,
    listPrice: 700,
    discountPercent: 0.15,
  },
]

const quoteRows = quoteItems.map((item, index) => {
  const spreadsheetRow = index + 2
  return [
    item.sku,
    item.quantity,
    item.listPrice,
    item.discountPercent,
    `=B${spreadsheetRow}*C${spreadsheetRow}`,
    `=E${spreadsheetRow}*D${spreadsheetRow}`,
    `=E${spreadsheetRow}-F${spreadsheetRow}`,
    `=IF(D${spreadsheetRow}>0.2,"Review","OK")`,
  ]
})

const workbook = WorkPaper.buildFromSheets({
  Quote: [['SKU', 'Quantity', 'List price', 'Discount %', 'List amount', 'Discount amount', 'Net amount', 'Approval flag'], ...quoteRows],
  Summary: [
    ['Metric', 'Value'],
    ['List total', '=SUM(Quote!E2:E5)'],
    ['Discount amount', '=SUM(Quote!F2:F5)'],
    ['Quote total', '=SUM(Quote!G2:G5)'],
    ['Discount percent', '=B3/B2'],
    ['Max line discount', '=MAX(Quote!D2:D5)'],
    ['Approval required', '=IF(COUNTIF(Quote!H2:H5,"Review")>0,"Review","Auto-approve")'],
  ],
})

const quoteSheet = requireSheet(workbook, 'Quote')
const summarySheet = requireSheet(workbook, 'Summary')
const approvalRequired = readString(workbook, summarySheet, 6, 1, 'Approval required')
const reviewedSku =
  readString(workbook, quoteSheet, 2, 7, 'Setup approval flag') === 'Review'
    ? readString(workbook, quoteSheet, 2, 0, 'Reviewed SKU')
    : undefined

const output = {
  quoteId: 'Q-2026-041',
  lineItems: quoteItems.length,
  listTotal: readNumber(workbook, summarySheet, 1, 1, 'List total'),
  discountAmount: readNumber(workbook, summarySheet, 2, 1, 'Discount amount'),
  quoteTotal: readNumber(workbook, summarySheet, 3, 1, 'Quote total'),
  discountPercent: readNumber(workbook, summarySheet, 4, 1, 'Discount percent'),
  maxLineDiscount: readNumber(workbook, summarySheet, 5, 1, 'Max line discount'),
  approvalRequired,
  reviewedSku,
  firstQuoteRow: workbook.getRangeSerialized({
    start: { sheet: quoteSheet, row: 1, col: 0 },
    end: { sheet: quoteSheet, row: 1, col: 7 },
  })[0],
  verified: true,
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

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

function assertOutput(actual: typeof output): void {
  const expected = {
    quoteId: 'Q-2026-041',
    lineItems: 4,
    listTotal: 6980,
    discountAmount: 993,
    quoteTotal: 5987,
    discountPercent: 0.1423,
    maxLineDiscount: 0.25,
    approvalRequired: 'Review',
    reviewedSku: 'SETUP',
    firstQuoteRow: ['PRO-ANNUAL', 12, 240, 0.1, '=B2*C2', '=E2*D2', '=E2-F2', '=IF(D2>0.2,"Review","OK")'],
    verified: true,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected quote approval result: ${JSON.stringify(actual)}`)
  }
}
