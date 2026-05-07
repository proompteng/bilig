import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Assumptions: [
    ['Metric', 'Value'],
    ['Leads', 500],
    ['Conversion Rate', 0.08],
    ['ARPA', 240],
    ['Expansion Factor', 1.1],
  ],
  Plan: [
    ['Metric', 'Value'],
    ['Customers', '=Assumptions!B2*Assumptions!B3'],
    ['Gross MRR', '=B2*Assumptions!B4'],
    ['Expansion MRR', '=B3*Assumptions!B5'],
    ['Annualized ARR', '=B4*12'],
  ],
  Review: [
    ['Check', 'Value'],
    ['ARR target delta', '=Plan!B5-150000'],
  ],
})

const assumptionsSheet = requireSheet(workbook, 'Assumptions')
const planSheet = requireSheet(workbook, 'Plan')
const reviewSheet = requireSheet(workbook, 'Review')

const plannedEdits = [
  { cell: 'Assumptions!B2', address: { sheet: assumptionsSheet, row: 1, col: 1 }, value: 650 },
  { cell: 'Assumptions!B3', address: { sheet: assumptionsSheet, row: 2, col: 1 }, value: 0.1 },
  { cell: 'Assumptions!B5', address: { sheet: assumptionsSheet, row: 4, col: 1 }, value: 1.2 },
]

const before = readModel(workbook, planSheet, reviewSheet)
const beforeFormulaContracts = readFormulaContracts(workbook, planSheet, reviewSheet)
const edits = plannedEdits.map((edit) => ({
  cell: edit.cell,
  before: workbook.getCellSerialized(edit.address),
  after: edit.value,
}))

for (const edit of plannedEdits) {
  workbook.setCellContents(edit.address, edit.value)
}

const after = readModel(workbook, planSheet, reviewSheet)
const afterFormulaContracts = readFormulaContracts(workbook, planSheet, reviewSheet)
const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
const restoredPlanSheet = requireSheet(restored, 'Plan')
const restoredReviewSheet = requireSheet(restored, 'Review')
const restoredModel = readModel(restored, restoredPlanSheet, restoredReviewSheet)
const restoredFormulaContracts = readFormulaContracts(restored, restoredPlanSheet, restoredReviewSheet)

const output = {
  edits,
  before,
  after,
  restored: restoredModel,
  formulaContracts: afterFormulaContracts,
  verified: {
    formulasUnchanged: sameJson(beforeFormulaContracts, afterFormulaContracts),
    formulasPersisted: sameJson(afterFormulaContracts, restoredFormulaContracts),
    restoredMatchesAfter: sameJson(after, restoredModel),
    serializedBytes: Buffer.byteLength(serialized, 'utf8'),
  },
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function readModel(workpaper, plan, review) {
  return {
    customers: readNumber(workpaper, plan, 1, 1, 'customers'),
    grossMrr: readNumber(workpaper, plan, 2, 1, 'gross MRR'),
    expansionMrr: readNumber(workpaper, plan, 3, 1, 'expansion MRR'),
    annualizedArr: readNumber(workpaper, plan, 4, 1, 'annualized ARR'),
    arrTargetDelta: readNumber(workpaper, review, 1, 1, 'ARR target delta'),
  }
}

function readFormulaContracts(workpaper, plan, review) {
  return {
    customers: readFormula(workpaper, plan, 1, 1, 'customers'),
    grossMrr: readFormula(workpaper, plan, 2, 1, 'gross MRR'),
    expansionMrr: readFormula(workpaper, plan, 3, 1, 'expansion MRR'),
    annualizedArr: readFormula(workpaper, plan, 4, 1, 'annualized ARR'),
    arrTargetDelta: readFormula(workpaper, review, 1, 1, 'ARR target delta'),
  }
}

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readFormula(workpaper, sheet, row, col, label) {
  const formula = workpaper.getCellFormula({ sheet, row, col })
  if (formula === undefined) {
    throw new Error(`Expected ${label} to be a formula`)
  }
  return formula
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertOutput(actual) {
  const expected = {
    edits: [
      { cell: 'Assumptions!B2', before: 500, after: 650 },
      { cell: 'Assumptions!B3', before: 0.08, after: 0.1 },
      { cell: 'Assumptions!B5', before: 1.1, after: 1.2 },
    ],
    before: {
      customers: 40,
      grossMrr: 9600,
      expansionMrr: 10560,
      annualizedArr: 126720,
      arrTargetDelta: -23280,
    },
    after: {
      customers: 65,
      grossMrr: 15600,
      expansionMrr: 18720,
      annualizedArr: 224640,
      arrTargetDelta: 74640,
    },
    formulaContracts: {
      customers: '=Assumptions!B2*Assumptions!B3',
      grossMrr: '=B2*Assumptions!B4',
      expansionMrr: '=B3*Assumptions!B5',
      annualizedArr: '=B4*12',
      arrTargetDelta: '=Plan!B5-150000',
    },
  }

  const comparable = {
    edits: actual.edits,
    before: actual.before,
    after: actual.after,
    formulaContracts: actual.formulaContracts,
  }

  if (
    !sameJson(comparable, expected) ||
    !sameJson(actual.restored, expected.after) ||
    actual.verified.formulasUnchanged !== true ||
    actual.verified.formulasPersisted !== true ||
    actual.verified.restoredMatchesAfter !== true ||
    actual.verified.serializedBytes <= 0
  ) {
    throw new Error(`Unexpected agent writeback verification result: ${JSON.stringify(actual)}`)
  }
}
