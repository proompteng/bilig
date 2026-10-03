import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>
type SetInputCellArgs = {
  sheetName: string
  address: string
  value: string | number | boolean | null
  reason?: string
}

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Qualified opportunities', 20],
    ['Win rate', 0.25],
    ['Average ARR', 12000],
    ['Expansion multiplier', 1.1],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Expected customers', '=Inputs!B2*Inputs!B3'],
    ['Expected ARR', '=B2*Inputs!B4'],
    ['Expansion ARR', '=B3*Inputs!B5'],
    ['Target gap', '=B4-100000'],
  ],
})

const summarySheet = requireSheet(workbook, 'Summary')
const agentToolCall = {
  toolName: 'setInputCell',
  arguments: {
    sheetName: 'Inputs',
    address: 'B3',
    value: 0.4,
    reason: 'Use the latest qualified pipeline conversion estimate.',
  },
}

const tools = {
  readRange(range = 'Summary!A1:B5') {
    const parsedRange = workbook.simpleCellRangeFromString(range, summarySheet)
    if (parsedRange === undefined) {
      throw new Error(`Invalid readable range: ${range}`)
    }

    return {
      range,
      values: workbook.getRangeValues(parsedRange),
      serialized: workbook.getRangeSerialized(parsedRange),
    }
  },

  setInputCell({ sheetName, address, value, reason }: SetInputCellArgs) {
    if (sheetName !== 'Inputs') {
      throw new Error(`This tool can only edit Inputs cells, received ${sheetName}`)
    }

    const target = requireCellAddress(workbook, sheetName, address)
    const before = readSummary(workbook, summarySheet)
    const formulaContracts = readFormulaContracts(workbook, summarySheet)
    const previousValue = workbook.getCellSerialized(target)

    workbook.setCellContents(target, value)

    const after = readSummary(workbook, summarySheet)
    const serialized = serializeWorkbook(workbook)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
    const restoredSummarySheet = requireSheet(restored, 'Summary')
    const restoredSummary = readSummary(restored, restoredSummarySheet)
    const restoredFormulaContracts = readFormulaContracts(restored, restoredSummarySheet)

    return {
      editedCell: workbook.simpleCellAddressToString(target, {
        includeSheetName: true,
      }),
      reason,
      before,
      after,
      restored: restoredSummary,
      formulaContracts,
      verified: {
        previousValue,
        newValue: workbook.getCellSerialized(target),
        formulasPersisted: sameJson(formulaContracts, restoredFormulaContracts),
        restoredMatchesAfter: sameJson(after, restoredSummary),
        expectedArrImproved: after.expectedArr > before.expectedArr,
        targetGapClosed: before.targetGap < 0 && after.targetGap > 0,
        serializedBytes: Buffer.byteLength(serialized, 'utf8'),
      },
    }
  },
}

const output = {
  toolCall: agentToolCall,
  initialRead: tools.readRange(),
  toolResult: tools.setInputCell(agentToolCall.arguments),
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

function requireCellAddress(workpaper: WorkPaperInstance, sheetName: string, a1Address: string): CellAddress {
  const sheetId = requireSheet(workpaper, sheetName)
  const parsed = workpaper.simpleCellAddressFromString(a1Address, sheetId)

  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid cell address: ${sheetName}!${a1Address}`)
  }

  return parsed
}

function readSummary(workpaper: WorkPaperInstance, summary: number) {
  return {
    expectedCustomers: readNumber(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readNumber(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readNumber(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readNumber(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readFormulaContracts(workpaper: WorkPaperInstance, summary: number) {
  return {
    expectedCustomers: readFormula(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readFormula(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readFormula(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readFormula(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readNumber(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readFormula(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): string {
  const formula = workpaper.getCellFormula({ sheet, row, col })
  if (formula === undefined) {
    throw new Error(`Expected ${label} to be a formula`)
  }
  return formula
}

function serializeWorkbook(workpaper: WorkPaperInstance): string {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workpaper, {
      includeConfig: true,
    }),
  )
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertOutput(actual: typeof output): void {
  const result = actual.toolResult
  const expectedBefore = {
    expectedCustomers: 5,
    expectedArr: 60000,
    expansionArr: 66000,
    targetGap: -34000,
  }
  const expectedAfter = {
    expectedCustomers: 8,
    expectedArr: 96000,
    expansionArr: 105600,
    targetGap: 5600,
  }
  const expectedFormulaContracts = {
    expectedCustomers: '=Inputs!B2*Inputs!B3',
    expectedArr: '=B2*Inputs!B4',
    expansionArr: '=B3*Inputs!B5',
    targetGap: '=B4-100000',
  }

  if (
    actual.toolCall.toolName !== 'setInputCell' ||
    result.editedCell !== 'Inputs!B3' ||
    !sameJson(result.before, expectedBefore) ||
    !sameJson(result.after, expectedAfter) ||
    !sameJson(result.restored, expectedAfter) ||
    !sameJson(result.formulaContracts, expectedFormulaContracts) ||
    result.verified.previousValue !== 0.25 ||
    result.verified.newValue !== 0.4 ||
    !result.verified.formulasPersisted ||
    !result.verified.restoredMatchesAfter ||
    !result.verified.expectedArrImproved ||
    !result.verified.targetGapClosed ||
    result.verified.serializedBytes <= 0
  ) {
    throw new Error(`Unexpected agent tool-call result: ${JSON.stringify(actual)}`)
  }
}
