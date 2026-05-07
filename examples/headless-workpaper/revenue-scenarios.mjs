import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Pipeline: [
    ['Segment', 'Leads', 'Conversion Rate', 'Customers', 'ARPA', 'Gross MRR', 'Churn Rate', 'Net MRR'],
    ['Enterprise', 80, 0.18, '=B2*C2', 4200, '=D2*E2', 0.05, '=F2*(1-G2)'],
    ['Mid-Market', 220, 0.22, '=B3*C3', 1100, '=D3*E3', 0.08, '=F3*(1-G3)'],
    ['SMB', 900, 0.09, '=B4*C4', 180, '=D4*E4', 0.12, '=F4*(1-G4)'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Total net MRR', '=SUM(Pipeline!H2:H4)'],
    ['Annual run rate', '=B2*12'],
    ['Enterprise net MRR', '=SUMIF(Pipeline!A2:A4,"Enterprise",Pipeline!H2:H4)'],
    ['Expansion target', '=B3*1.18'],
  ],
  Scenarios: [
    ['Scenario', 'Multiplier', 'Projected Net MRR', 'Projected ARR'],
    ['Conservative', 0.9, '=Summary!B2*B2', '=C2*12'],
    ['Expansion', 1.15, '=Summary!B2*B3', '=C3*12'],
    ['Stretch', 1.35, '=Summary!B2*B4', '=C4*12'],
  ],
})

const pipelineSheet = requireSheet(workbook, 'Pipeline')
const beforeEdit = readModel(workbook)

workbook.batch(() => {
  workbook.setCellContents({ sheet: pipelineSheet, row: 1, col: 1 }, 92)
  workbook.setCellContents({ sheet: pipelineSheet, row: 2, col: 2 }, 0.26)
})

const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
const afterEdit = readModel(restored)

const output = {
  beforeEdit,
  afterEdit,
  persistedSheets: restored.getSheetNames(),
  serializedBytes: Buffer.byteLength(serialized, 'utf8'),
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function readModel(workpaper) {
  const summarySheet = requireSheet(workpaper, 'Summary')
  const scenariosSheet = requireSheet(workpaper, 'Scenarios')

  return {
    totalNetMrr: readCurrency(workpaper, summarySheet, 1, 1, 'total net MRR'),
    annualRunRate: readCurrency(workpaper, summarySheet, 2, 1, 'annual run rate'),
    enterpriseNetMrr: readCurrency(workpaper, summarySheet, 3, 1, 'enterprise net MRR'),
    expansionTarget: readCurrency(workpaper, summarySheet, 4, 1, 'expansion target'),
    scenarios: {
      conservativeNetMrr: readCurrency(workpaper, scenariosSheet, 1, 2, 'conservative net MRR'),
      expansionNetMrr: readCurrency(workpaper, scenariosSheet, 2, 2, 'expansion net MRR'),
      stretchNetMrr: readCurrency(workpaper, scenariosSheet, 3, 2, 'stretch net MRR'),
    },
  }
}

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readCurrency(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function assertOutput(actual) {
  const expected = {
    beforeEdit: {
      totalNetMrr: 119267.2,
      annualRunRate: 1431206.4,
      enterpriseNetMrr: 57456,
      expansionTarget: 1688823.55,
      scenarios: {
        conservativeNetMrr: 107340.48,
        expansionNetMrr: 137157.28,
        stretchNetMrr: 161010.72,
      },
    },
    afterEdit: {
      totalNetMrr: 136791.2,
      annualRunRate: 1641494.4,
      enterpriseNetMrr: 66074.4,
      expansionTarget: 1936963.39,
      scenarios: {
        conservativeNetMrr: 123112.08,
        expansionNetMrr: 157309.88,
        stretchNetMrr: 184668.12,
      },
    },
    persistedSheets: ['Pipeline', 'Summary', 'Scenarios'],
  }

  const comparable = {
    beforeEdit: actual.beforeEdit,
    afterEdit: actual.afterEdit,
    persistedSheets: actual.persistedSheets,
  }

  if (JSON.stringify(comparable) !== JSON.stringify(expected) || actual.serializedBytes <= 0) {
    throw new Error(`Unexpected revenue scenario result: ${JSON.stringify(actual)}`)
  }
}
