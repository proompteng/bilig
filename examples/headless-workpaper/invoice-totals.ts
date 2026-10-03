import { WorkPaper } from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

const lineItems = [
  {
    description: 'Implementation workshop',
    quantity: 5,
    unitPrice: 120,
  },
  {
    description: 'Data migration package',
    quantity: 2,
    unitPrice: 450,
  },
  {
    description: 'Admin training',
    quantity: 1,
    unitPrice: 150,
  },
  {
    description: 'Support handoff',
    quantity: 3,
    unitPrice: 80,
  },
]

const lineItemRows = lineItems.map((item, index) => {
  const row = index + 2
  return [item.description, item.quantity, item.unitPrice, `=B${row}*C${row}`]
})

const workbook = WorkPaper.buildFromSheets({
  Invoice: [['Description', 'Quantity', 'Unit price', 'Line total'], ...lineItemRows],
  Summary: [
    ['Metric', 'Value'],
    ['Subtotal', '=SUM(Invoice!D2:D5)'],
    ['Tax rate', 0.08],
    ['Tax', '=B2*B3'],
    ['Total', '=B2+B4'],
  ],
})

const invoiceSheet = requireSheet(workbook, 'Invoice')
const summarySheet = requireSheet(workbook, 'Summary')

const output = {
  invoiceNumber: 'INV-2026-001',
  lineItems: lineItems.length,
  subtotal: readNumber(workbook, summarySheet, 1, 1, 'Subtotal'),
  taxRate: readNumber(workbook, summarySheet, 2, 1, 'Tax rate'),
  tax: readNumber(workbook, summarySheet, 3, 1, 'Tax'),
  total: readNumber(workbook, summarySheet, 4, 1, 'Total'),
  formulas: workbook.getRangeSerialized({
    start: { sheet: summarySheet, row: 1, col: 1 },
    end: { sheet: summarySheet, row: 4, col: 1 },
  }),
  firstLineItem: workbook.getRangeSerialized({
    start: { sheet: invoiceSheet, row: 1, col: 0 },
    end: { sheet: invoiceSheet, row: 1, col: 3 },
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
  return Math.round(cell.value * 100) / 100
}

function assertOutput(actual: typeof output): void {
  const expected = {
    invoiceNumber: 'INV-2026-001',
    lineItems: 4,
    subtotal: 1890,
    taxRate: 0.08,
    tax: 151.2,
    total: 2041.2,
    formulas: [['=SUM(Invoice!D2:D5)'], [0.08], ['=B2*B3'], ['=B2+B4']],
    firstLineItem: ['Implementation workshop', 5, 120, '=B2*C2'],
    verified: true,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected invoice totals result: ${JSON.stringify(actual)}`)
  }
}
