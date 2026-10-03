import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'
import { z } from 'zod'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>
type CellInputValue = string | number | boolean | null
type SetInputCellArgs = {
  sheetName: string
  address: string
  value: CellInputValue
}
type WorkPaperToolResult = WorkPaperReadResult | WorkPaperWriteResult
type WorkPaperReadResult = ReturnType<WorkPaperTools['readWorkPaperSummary']>
type WorkPaperWriteResult = ReturnType<WorkPaperTools['setWorkPaperInputCell']>
type WorkPaperTools = ReturnType<typeof createWorkPaperTools>
type OpenAiResponsesFunctionTool = {
  type: 'function'
  name: 'read_workpaper_summary' | 'set_workpaper_input_cell'
  description: string
  parameters: Record<string, unknown>
  strict: true
}
type OpenAiResponsesFunctionCall =
  | {
      type: 'function_call'
      call_id: string
      name: 'read_workpaper_summary'
      arguments: string
    }
  | {
      type: 'function_call'
      call_id: string
      name: 'set_workpaper_input_cell'
      arguments: string
    }
type OpenAiResponsesFunctionCallOutput = {
  type: 'function_call_output'
  call_id: string
  output: string
}
type OpenAiResponsesUserMessage = {
  role: 'user'
  content: string
}
type OpenAiResponsesInputItem = OpenAiResponsesUserMessage | OpenAiResponsesFunctionCall | OpenAiResponsesFunctionCallOutput

const readSummaryInputSchema = z.object({
  range: z.string().default('Summary!A1:B5'),
})

const setInputCellInputSchema = z.object({
  sheetName: z.literal('Inputs'),
  address: z.string().regex(/^[A-Z]+[1-9][0-9]*$/),
  value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
})

const workbook = buildWorkbook()
const workPaperTools = createWorkPaperTools(workbook)
const openAiTools = createOpenAiResponsesTools()
const firstInput: OpenAiResponsesInputItem[] = [
  {
    role: 'user',
    content: 'Read the revenue summary, then set the win rate input to 0.4 and report the calculated ARR.',
  },
]

const modelOutput: OpenAiResponsesFunctionCall[] = [
  {
    type: 'function_call',
    call_id: 'call_read_summary',
    name: 'read_workpaper_summary',
    arguments: JSON.stringify({ range: 'Summary!A1:B5' }),
  },
  {
    type: 'function_call',
    call_id: 'call_set_win_rate',
    name: 'set_workpaper_input_cell',
    arguments: JSON.stringify({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    }),
  },
]

const toolResults = modelOutput.map((call) => dispatchOpenAiResponsesCall(workPaperTools, call))
const toolOutputs = modelOutput.map(
  (call, index): OpenAiResponsesFunctionCallOutput => ({
    type: 'function_call_output',
    call_id: call.call_id,
    output: JSON.stringify(toolResults[index]),
  }),
)
const followupInput: OpenAiResponsesInputItem[] = [...firstInput, ...modelOutput, ...toolOutputs]
const writeResult = requireWriteResult(toolResults[1])
const finalAssistantMessage = [
  `Edited ${writeResult.editedCell}.`,
  `Expected ARR moved from ${String(writeResult.before.expectedArr)} to ${String(writeResult.after.expectedArr)}.`,
  `The restored workbook still reads ${String(writeResult.restored.expectedArr)}.`,
].join(' ')

const output = {
  apiShape: 'OpenAI Responses function_call -> function_call_output',
  toolNames: openAiTools.map((tool) => tool.name),
  firstInput,
  modelOutput,
  toolOutputs,
  followupInputTypes: followupInput.map((item) => ('type' in item ? item.type : item.role)),
  finalAssistantMessage,
  writeResult,
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function buildWorkbook() {
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

function createOpenAiResponsesTools(): OpenAiResponsesFunctionTool[] {
  return [
    {
      type: 'function',
      name: 'read_workpaper_summary',
      description: 'Read computed WorkPaper summary values for a small A1 range.',
      parameters: {
        type: 'object',
        required: ['range'],
        properties: {
          range: {
            type: 'string',
            description: 'Small A1 range including the sheet name.',
            default: 'Summary!A1:B5',
          },
        },
        additionalProperties: false,
      },
      strict: true,
    },
    {
      type: 'function',
      name: 'set_workpaper_input_cell',
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      parameters: {
        type: 'object',
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: {
            type: 'string',
            description: 'Editable sheet name. This example allows Inputs only.',
          },
          address: {
            type: 'string',
            description: 'A1 address inside the target sheet.',
          },
          value: {
            type: ['string', 'number', 'boolean', 'null'],
            description: 'Literal input value. Use a separate tool for formulas.',
          },
        },
        additionalProperties: false,
      },
      strict: true,
    },
  ]
}

function dispatchOpenAiResponsesCall(tools: WorkPaperTools, call: OpenAiResponsesFunctionCall): WorkPaperToolResult {
  if (call.name === 'read_workpaper_summary') {
    const args = readSummaryInputSchema.parse(JSON.parse(call.arguments))
    return tools.readWorkPaperSummary(args.range)
  }

  const args = setInputCellInputSchema.parse(JSON.parse(call.arguments))
  return tools.setWorkPaperInputCell(args)
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

    setWorkPaperInputCell({ sheetName, address, value }: SetInputCellArgs) {
      if (sheetName !== 'Inputs') {
        throw new Error(`This example only permits Inputs edits, received ${sheetName}`)
      }

      const target = requireCellAddress(workpaper, sheetName, address)
      const before = readSummary(workpaper, summarySheet)
      const formulaContracts = readFormulaContracts(workpaper, summarySheet)
      const previousValue = workpaper.getCellSerialized(target)

      workpaper.setCellContents(target, value)

      const after = readSummary(workpaper, summarySheet)
      const serialized = serializeWorkbook(workpaper)
      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
      const restoredSummarySheet = requireSheet(restored, 'Summary')
      const restoredSummary = readSummary(restored, restoredSummarySheet)
      const restoredFormulaContracts = readFormulaContracts(restored, restoredSummarySheet)

      return {
        editedCell: workpaper.simpleCellAddressToString(target, {
          includeSheetName: true,
        }),
        before,
        after,
        restored: restoredSummary,
        formulaContracts,
        checks: {
          previousValue,
          newValue: workpaper.getCellSerialized(target),
          formulasPersisted: sameJson(formulaContracts, restoredFormulaContracts),
          restoredMatchesAfter: sameJson(after, restoredSummary),
          expectedArrChanged: after.expectedArr > before.expectedArr,
          serializedBytes: Buffer.byteLength(serialized, 'utf8'),
        },
      }
    },
  }
}

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

function requireWriteResult(result: WorkPaperToolResult | undefined): WorkPaperWriteResult {
  if (result === undefined || !('editedCell' in result)) {
    throw new Error(`Expected WorkPaper write result, received ${JSON.stringify(result)}`)
  }
  return result
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
    !sameJson(actual.toolNames, ['read_workpaper_summary', 'set_workpaper_input_cell']) ||
    !sameJson(actual.followupInputTypes, ['user', 'function_call', 'function_call', 'function_call_output', 'function_call_output']) ||
    actual.toolOutputs.length !== 2 ||
    actual.toolOutputs.some((item) => item.type !== 'function_call_output') ||
    actual.writeResult.editedCell !== 'Inputs!B3' ||
    !sameJson(actual.writeResult.before, expectedBefore) ||
    !sameJson(actual.writeResult.after, expectedAfter) ||
    !sameJson(actual.writeResult.restored, expectedAfter) ||
    !sameJson(actual.writeResult.formulaContracts, expectedFormulaContracts) ||
    actual.writeResult.checks.previousValue !== 0.25 ||
    actual.writeResult.checks.newValue !== 0.4 ||
    !actual.writeResult.checks.formulasPersisted ||
    !actual.writeResult.checks.restoredMatchesAfter ||
    !actual.writeResult.checks.expectedArrChanged ||
    actual.writeResult.checks.serializedBytes <= 0 ||
    !actual.finalAssistantMessage.includes('Expected ARR moved from 60000 to 96000')
  ) {
    throw new Error(`Unexpected OpenAI Responses wrapper output: ${JSON.stringify(actual)}`)
  }
}
