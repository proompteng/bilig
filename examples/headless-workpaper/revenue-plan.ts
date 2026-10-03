import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

const workbook = WorkPaper.buildFromSheets({
  Deals: [
    ['Region', 'Segment', 'Customers', 'ARPA', 'Revenue'],
    ['West', 'Enterprise', 12, 1200, '=C2*D2'],
    ['East', 'SMB', 30, 250, '=C3*D3'],
    ['West', 'SMB', 18, 300, '=C4*D4'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Total revenue', '=SUM(Deals!E2:E4)'],
    ['West customers', '=SUMIF(Deals!A2:A4,"West",Deals!C2:C4)'],
    ['Enterprise ARPA', '=XLOOKUP("Enterprise",Deals!B2:B4,Deals!D2:D4)'],
    ['Target revenue', null],
    ['Qualified customer counts', '=FILTER(Deals!C2:C4,Deals!C2:C4>=18)'],
  ],
})

const dealsSheet = requireSheet(workbook, 'Deals')
const summarySheet = requireSheet(workbook, 'Summary')

workbook.addNamedExpression('GrowthRatePercent', 12)
workbook.setCellContents({ sheet: summarySheet, row: 4, col: 1 }, '=SUM(Deals!E2:E4)*(100+GrowthRatePercent)/100')

const initial = {
  totalRevenue: readNumber(workbook, summarySheet, 1, 1, 'initial total revenue'),
  westCustomers: readNumber(workbook, summarySheet, 2, 1, 'initial west customers'),
  targetRevenue: readNumber(workbook, summarySheet, 4, 1, 'initial target revenue'),
}

workbook.batch(() => {
  workbook.setCellContents({ sheet: dealsSheet, row: 1, col: 2 }, 20)
})

const persisted = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(persisted))
const restoredSummarySheet = requireSheet(restored, 'Summary')

const afterAgentEdit = {
  totalRevenue: readNumber(restored, restoredSummarySheet, 1, 1, 'edited total revenue'),
  westCustomers: readNumber(restored, restoredSummarySheet, 2, 1, 'edited west customers'),
  enterpriseArpa: readNumber(restored, restoredSummarySheet, 3, 1, 'enterprise ARPA'),
  targetRevenue: readNumber(restored, restoredSummarySheet, 4, 1, 'edited target revenue'),
  qualifiedCustomerCounts: readNumberColumn(restored, restoredSummarySheet, 5, 7, 1, 'qualified customer count'),
}

const output = {
  initial,
  afterAgentEdit,
  persistedSheets: restored.getSheetNames(),
  persistedNamedExpressions: restored.listNamedExpressions(),
  restoredGrowthRatePercent: readNamedNumber(restored, 'GrowthRatePercent'),
}

assertSummary(output)
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
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function readNumberColumn(
  workpaper: WorkPaperInstance,
  sheet: number,
  startRow: number,
  endRow: number,
  col: number,
  label: string,
): number[] {
  return workpaper
    .getRangeValues({
      start: { sheet, row: startRow, col },
      end: { sheet, row: endRow, col },
    })
    .flat()
    .map((cell: unknown, index: number) => {
      if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
        throw new Error(`Expected ${label} ${index + 1} to be a number, received ${JSON.stringify(cell)}`)
      }
      return cell.value
    })
}

function readNamedNumber(workpaper: WorkPaperInstance, name: string): number {
  const cell = workpaper.getNamedExpressionValue(name)
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected named expression "${name}" to be a number, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function assertSummary(summary: typeof output): void {
  const expected = {
    initial: {
      totalRevenue: 27300,
      westCustomers: 30,
      targetRevenue: 30576,
    },
    afterAgentEdit: {
      totalRevenue: 36900,
      westCustomers: 38,
      enterpriseArpa: 1200,
      targetRevenue: 41328,
      qualifiedCustomerCounts: [20, 30, 18],
    },
    persistedSheets: ['Deals', 'Summary'],
    persistedNamedExpressions: ['GrowthRatePercent'],
    restoredGrowthRatePercent: 12,
  }

  if (JSON.stringify(summary) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected WorkPaper result: ${JSON.stringify(summary)}`)
  }
}
