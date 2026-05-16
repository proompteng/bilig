import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '../packages/headless/dist/index.js'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 20],
    ['Average revenue', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Node 22 smoke workbook is missing required sheets')
}

const revenueAddress = { sheet: summary, row: 1, col: 1 }
const before = workbook.getCellDisplayValue(revenueAddress)
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)
const after = workbook.getCellDisplayValue(revenueAddress)

const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('Node 22 smoke restored workbook is missing Summary sheet')
}

const afterRestore = restored.getCellDisplayValue({ sheet: restoredSummary, row: 1, col: 1 })
const verified = before === '24000' && after === '38400' && afterRestore === '38400'
console.log(
  JSON.stringify(
    {
      node: process.version,
      before,
      after,
      afterRestore,
      verified,
    },
    null,
    2,
  ),
)

if (!verified) {
  throw new Error('Node 22 headless smoke failed formula readback or restore verification')
}
