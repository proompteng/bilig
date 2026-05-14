import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'
import { tool } from 'ai'
import { z } from 'zod'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>
type WorkPaperToolSet = ReturnType<typeof createWorkPaperTools>

export type SetInputCellArgs = {
  sheetName: string
  address: string
  value: string | number | boolean | null
}

export type ReadSummaryArgs = {
  range?: string
}

export type SummaryReadback = {
  expectedCustomers: number
  expectedArr: number
  expansionArr: number
  targetGap: number
}

export type FormulaContracts = {
  expectedCustomers: string
  expectedArr: string
  expansionArr: string
  targetGap: string
}

export type WorkPaperReadResult = {
  range: string
  values: unknown[][]
  serialized: unknown[][]
}

export type WorkPaperWriteResult = {
  editedCell: string
  before: SummaryReadback
  after: SummaryReadback
  restored: SummaryReadback
  beforeContracts: FormulaContracts
  afterContracts: FormulaContracts
  checks: {
    previousValue: unknown
    newValue: unknown
    formulasPersisted: boolean
    restoredMatchesAfter: boolean
    expectedArrChanged: boolean
    serializedBytes: number
  }
}

export type AiSdkWorkPaperToolResult = {
  toolCallId: string
  toolName: string
  output: unknown
}

export type AiSdkWorkPaperSmokeProof = {
  apiShape: string
  toolNames: string[]
  toolCalls: Array<{ toolCallId: string; toolName: string; input: unknown }>
  readResult: WorkPaperReadResult
  writeResult: WorkPaperWriteResult
  text: string
}

const readSummaryInputSchema = z.object({
  range: z.string().default('Summary!A1:B5'),
})

const setInputCellInputSchema = z.object({
  sheetName: z.literal('Inputs'),
  address: z.string().regex(/^[A-Z]+[1-9][0-9]*$/),
  value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
})

export const modelUsage = {
  inputTokens: {
    total: 0,
    noCache: 0,
    cacheRead: 0,
    cacheWrite: 0,
  },
  outputTokens: {
    total: 0,
    text: 0,
    reasoning: 0,
  },
}

export function buildWorkbook() {
  return WorkPaper.buildFromSheets({
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
}

export function createAiSdkWorkPaperTools(workpaper = buildWorkbook()) {
  return createAiSdkTools(createWorkPaperTools(workpaper))
}

export function requireToolOutput(
  toolResults: ReadonlyArray<AiSdkWorkPaperToolResult>,
  toolCallId: string,
  toolName: 'readWorkPaperSummary' | 'setWorkPaperInputCell',
): unknown {
  const toolResult = toolResults.find((item) => item.toolCallId === toolCallId && item.toolName === toolName)

  if (toolResult === undefined) {
    throw new Error(`Missing ${toolName} result for ${toolCallId}; received ${JSON.stringify(toolResults)}`)
  }

  return toolResult.output
}

export function requireWorkPaperReadResult(value: unknown): WorkPaperReadResult {
  if (
    isRecord(value) &&
    'range' in value &&
    typeof value.range === 'string' &&
    'values' in value &&
    Array.isArray(value.values) &&
    'serialized' in value &&
    Array.isArray(value.serialized)
  ) {
    return {
      range: value.range,
      values: value.values,
      serialized: value.serialized,
    }
  }

  throw new Error(`Expected WorkPaper read result, received ${JSON.stringify(value)}`)
}

export function requireWorkPaperWriteResult(value: unknown): WorkPaperWriteResult {
  if (
    isRecord(value) &&
    'editedCell' in value &&
    typeof value.editedCell === 'string' &&
    isSummaryReadback(value.before) &&
    isSummaryReadback(value.after) &&
    isSummaryReadback(value.restored) &&
    isFormulaContracts(value.beforeContracts) &&
    isFormulaContracts(value.afterContracts) &&
    isWriteChecks(value.checks)
  ) {
    return {
      editedCell: value.editedCell,
      before: value.before,
      after: value.after,
      restored: value.restored,
      beforeContracts: value.beforeContracts,
      afterContracts: value.afterContracts,
      checks: value.checks,
    }
  }

  throw new Error(`Expected WorkPaper write result, received ${JSON.stringify(value)}`)
}

export function assertAiSdkWorkPaperSmokeProof(output: AiSdkWorkPaperSmokeProof, expectedApiShape: string): void {
  if (output.apiShape !== expectedApiShape) {
    throw new Error(`Unexpected API shape: ${output.apiShape}`)
  }

  if (!sameJson(output.toolNames, ['readWorkPaperSummary', 'setWorkPaperInputCell'])) {
    throw new Error(`Unexpected AI SDK tool names: ${JSON.stringify(output.toolNames)}`)
  }

  if (output.toolCalls.length !== 2) {
    throw new Error(`Expected two AI SDK tool calls, received ${output.toolCalls.length}`)
  }

  if (readGridNumber(output.readResult.values, 2, 1, 'read summary expected ARR') !== 60000) {
    throw new Error(`Unexpected read summary before edit: ${JSON.stringify(output.readResult.values)}`)
  }

  if (output.writeResult.editedCell !== 'Inputs!B3') {
    throw new Error(`Unexpected edited cell: ${output.writeResult.editedCell}`)
  }

  if (output.writeResult.before.expectedArr !== 60000 || output.writeResult.after.expectedArr !== 96000) {
    throw new Error(`Unexpected ARR readback: ${JSON.stringify(output.writeResult)}`)
  }

  if (
    output.writeResult.checks.previousValue !== 0.25 ||
    output.writeResult.checks.newValue !== 0.4 ||
    !output.writeResult.checks.formulasPersisted ||
    !output.writeResult.checks.restoredMatchesAfter ||
    !output.writeResult.checks.expectedArrChanged
  ) {
    throw new Error(`AI SDK WorkPaper checks failed: ${JSON.stringify(output.writeResult.checks)}`)
  }

  if (!output.text.includes('Edited Inputs!B3')) {
    throw new Error(`Unexpected final AI SDK text: ${output.text}`)
  }
}

function createAiSdkTools(localWorkPaperTools: WorkPaperToolSet) {
  return {
    readWorkPaperSummary: tool({
      description: 'Read computed WorkPaper summary values for a small range.',
      inputSchema: readSummaryInputSchema,
      execute: async ({ range = 'Summary!A1:B5' }: ReadSummaryArgs = {}) => localWorkPaperTools.readWorkPaperSummary(range),
    }),

    setWorkPaperInputCell: tool({
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      inputSchema: setInputCellInputSchema,
      execute: async (args: SetInputCellArgs) => localWorkPaperTools.setWorkPaperInputCell(args),
    }),
  }
}

function createWorkPaperTools(workpaper: WorkPaperInstance) {
  const summarySheet = requireSheet(workpaper, 'Summary')

  return {
    readWorkPaperSummary(range = 'Summary!A1:B5') {
      const parsedRange = workpaper.simpleCellRangeFromString(range, summarySheet)
      if (parsedRange === undefined) {
        throw new Error(`Invalid readable range: ${range}`)
      }

      return {
        range,
        values: workpaper.getRangeValues(parsedRange),
        serialized: workpaper.getRangeSerialized(parsedRange),
      }
    },

    setWorkPaperInputCell(args: SetInputCellArgs) {
      const parsedArgs = setInputCellInputSchema.parse(args)
      const address = requireCellAddress(workpaper, parsedArgs.sheetName, parsedArgs.address)
      const before = readSummary(workpaper, summarySheet)
      const beforeContracts = readFormulaContracts(workpaper, summarySheet)
      const previousValue = workpaper.getCellSerialized(address)

      workpaper.setCellContents(address, parsedArgs.value)

      const after = readSummary(workpaper, summarySheet)
      const afterContracts = readFormulaContracts(workpaper, summarySheet)
      const saved = serializeWorkPaperDocument(
        exportWorkPaperDocument(workpaper, {
          includeConfig: true,
        }),
      )
      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
      const restoredSummarySheet = requireSheet(restored, 'Summary')
      const restoredSummary = readSummary(restored, restoredSummarySheet)
      const restoredFormulaContracts = readFormulaContracts(restored, restoredSummarySheet)

      return {
        editedCell: workpaper.simpleCellAddressToString(address, {
          includeSheetName: true,
        }),
        before,
        after,
        restored: restoredSummary,
        beforeContracts,
        afterContracts,
        checks: {
          previousValue,
          newValue: workpaper.getCellSerialized(address),
          formulasPersisted: sameJson(afterContracts, restoredFormulaContracts),
          restoredMatchesAfter: sameJson(after, restoredSummary),
          expectedArrChanged: after.expectedArr > before.expectedArr,
          serializedBytes: Buffer.byteLength(saved, 'utf8'),
        },
      }
    },
  }
}

function requireCellAddress(workpaper: WorkPaperInstance, sheetName: string, address: string): CellAddress {
  const sheetId = requireSheet(workpaper, sheetName)
  const parsedAddress = workpaper.simpleCellAddressFromString(address, sheetId)

  if (parsedAddress === undefined || parsedAddress.sheet !== sheetId) {
    throw new Error(`Invalid WorkPaper address: ${sheetName}!${address}`)
  }

  return parsedAddress
}

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readSummary(workpaper: WorkPaperInstance, summarySheet: number) {
  return {
    expectedCustomers: readNumber(workpaper, summarySheet, 1, 1, 'expected customers'),
    expectedArr: readNumber(workpaper, summarySheet, 2, 1, 'expected ARR'),
    expansionArr: readNumber(workpaper, summarySheet, 3, 1, 'expansion ARR'),
    targetGap: readNumber(workpaper, summarySheet, 4, 1, 'target gap'),
  }
}

function readFormulaContracts(workpaper: WorkPaperInstance, summarySheet: number) {
  return {
    expectedCustomers: readFormula(workpaper, summarySheet, 1, 1, 'expected customers'),
    expectedArr: readFormula(workpaper, summarySheet, 2, 1, 'expected ARR'),
    expansionArr: readFormula(workpaper, summarySheet, 3, 1, 'expansion ARR'),
    targetGap: readFormula(workpaper, summarySheet, 4, 1, 'target gap'),
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

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function readGridNumber(values: WorkPaperReadResult['values'], row: number, col: number, label: string): number {
  const cell = values[row]?.[col]
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSummaryReadback(value: unknown): value is SummaryReadback {
  return (
    isRecord(value) &&
    typeof value.expectedCustomers === 'number' &&
    typeof value.expectedArr === 'number' &&
    typeof value.expansionArr === 'number' &&
    typeof value.targetGap === 'number'
  )
}

function isFormulaContracts(value: unknown): value is FormulaContracts {
  return (
    isRecord(value) &&
    typeof value.expectedCustomers === 'string' &&
    typeof value.expectedArr === 'string' &&
    typeof value.expansionArr === 'string' &&
    typeof value.targetGap === 'string'
  )
}

function isWriteChecks(value: unknown): value is WorkPaperWriteResult['checks'] {
  return (
    isRecord(value) &&
    'previousValue' in value &&
    'newValue' in value &&
    typeof value.formulasPersisted === 'boolean' &&
    typeof value.restoredMatchesAfter === 'boolean' &&
    typeof value.expectedArrChanged === 'boolean' &&
    typeof value.serializedBytes === 'number'
  )
}
