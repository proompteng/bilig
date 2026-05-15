import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'
import { z } from 'zod'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>
type SetInputCellArgs = {
  sheetName: string
  address: string
  value: string | number | boolean | null
}
type ReadSummaryArgs = {
  range?: string
}
type WorkPaperToolSet = ReturnType<typeof createWorkPaperTools>
type WorkPaperReadResult = ReturnType<WorkPaperToolSet['readWorkPaperSummary']>
type WorkPaperWriteResult = ReturnType<WorkPaperToolSet['setWorkPaperInputCell']>
type OpenAiResponsesAdapter = ReturnType<typeof createOpenAiResponsesTools>
type OpenAiFunctionTool = {
  type: 'function'
  name: 'read_workpaper_summary' | 'set_workpaper_input_cell'
  description: string
  parameters: Record<string, unknown>
  strict: true
}
type OpenAiFunctionCall =
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
type OpenAiFunctionCallOutput = {
  type: 'function_call_output'
  call_id: string
  output: string
}
type LangChainReadTool = {
  name: 'read_workpaper_summary'
  description: string
  schema: typeof readSummaryInputSchema
  invoke(args?: ReadSummaryArgs): ReturnType<WorkPaperToolSet['readWorkPaperSummary']>
}
type LangChainWriteTool = {
  name: 'set_workpaper_input_cell'
  description: string
  schema: typeof setInputCellInputSchema
  invoke(args: SetInputCellArgs): WorkPaperWriteResult
}
type LangChainTool = LangChainReadTool | LangChainWriteTool
type LlamaIndexTool =
  | {
      name: 'read_workpaper_summary'
      description: string
      parameters: typeof readSummaryInputSchema
      call(args?: ReadSummaryArgs): WorkPaperReadResult
    }
  | {
      name: 'set_workpaper_input_cell'
      description: string
      parameters: typeof setInputCellInputSchema
      call(args: SetInputCellArgs): WorkPaperWriteResult
    }
type CopilotKitAction =
  | {
      name: 'readWorkPaperSummary'
      description: string
      parameters: CopilotKitParameter[]
      handler(args?: ReadSummaryArgs): WorkPaperReadResult
    }
  | {
      name: 'setWorkPaperInputCell'
      description: string
      parameters: CopilotKitParameter[]
      handler(args: SetInputCellArgs): WorkPaperWriteResult
    }
type CopilotKitParameter = {
  name: string
  type: 'string' | 'number' | 'boolean'
  description: string
  required?: boolean
}
type LangGraphToolMessage = {
  type: 'tool'
  name: LangChainTool['name']
  tool_call_id: string
  content: string
  result: WorkPaperReadResult | WorkPaperWriteResult
}
type LangGraphToolNode = {
  nodeName: 'tools'
  tools: LangChainTool[]
  invoke(input: { messages: Array<{ tool_calls?: LangGraphToolCall[] }> }): { messages: LangGraphToolMessage[] }
}
type LangGraphToolCall =
  | {
      id: string
      name: 'read_workpaper_summary'
      args: ReadSummaryArgs
    }
  | {
      id: string
      name: 'set_workpaper_input_cell'
      args: SetInputCellArgs
    }
type CrewAiTool =
  | {
      name: 'read_workpaper_summary'
      description: string
      argsSchema: typeof readSummaryInputSchema
      run(args?: ReadSummaryArgs): WorkPaperReadResult
    }
  | {
      name: 'set_workpaper_input_cell'
      description: string
      argsSchema: typeof setInputCellInputSchema
      run(args: SetInputCellArgs): WorkPaperWriteResult
    }

const readSummaryInputSchema = z.object({
  range: z.string().default('Summary!A1:B5'),
})

const setInputCellInputSchema = z.object({
  sheetName: z.literal('Inputs'),
  address: z.string().regex(/^[A-Z]+[1-9][0-9]*$/),
  value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
})

const workPaperSummaryOutputSchema = z
  .object({
    range: z.string(),
    values: z.array(z.array(z.unknown())),
    serialized: z.array(z.array(z.unknown())),
  })
  .passthrough()

const workPaperWriteOutputSchema = z
  .object({
    editedCell: z.string(),
    checks: z.object({
      formulasPersisted: z.boolean(),
      restoredMatchesAfter: z.boolean(),
      expectedArrChanged: z.boolean(),
      serializedBytes: z.number(),
    }),
  })
  .passthrough()

const aiSdkWorkbook = buildWorkbook()
const openAiResponsesWorkbook = buildWorkbook()
const langChainWorkbook = buildWorkbook()
const mastraWorkbook = buildWorkbook()
const llamaIndexWorkbook = buildWorkbook()
const langGraphWorkbook = buildWorkbook()
const copilotKitWorkbook = buildWorkbook()
const cloudflareAgentsWorkbook = buildWorkbook()
const crewAiWorkbook = buildWorkbook()
const aiSdkTools = createAiSdkTools(createWorkPaperTools(aiSdkWorkbook))
const openAiResponsesTools = createOpenAiResponsesTools(createWorkPaperTools(openAiResponsesWorkbook))
const langChainTools = createLangChainTools(createWorkPaperTools(langChainWorkbook))
const langChainStructuredToolSmoke = runLangChainStructuredToolSmoke(langChainTools)
const mastraTools = createMastraTools(createWorkPaperTools(mastraWorkbook))
const llamaIndexTools = createLlamaIndexTools(createWorkPaperTools(llamaIndexWorkbook))
const langGraphToolNode = createLangGraphToolNode(createWorkPaperTools(langGraphWorkbook))
const copilotKitActions = createCopilotKitActions(createWorkPaperTools(copilotKitWorkbook))
const cloudflareAgentTools = createCloudflareAgentTools(createWorkPaperTools(cloudflareAgentsWorkbook))
const crewAiTools = createCrewAiTools(createWorkPaperTools(crewAiWorkbook))

const output = {
  aiSdk: {
    toolNames: Object.keys(aiSdkTools),
    readResult: aiSdkTools.readWorkPaperSummary.execute({
      range: 'Summary!A1:B5',
    }),
    writeResult: aiSdkTools.setWorkPaperInputCell.execute({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    }),
  },
  openAiResponses: runOpenAiResponsesToolLoop(openAiResponsesTools),
  langChain: langChainStructuredToolSmoke,
  mastra: {
    toolIds: [mastraTools.readWorkPaperSummary.id, mastraTools.setWorkPaperInputCell.id],
    readResult: mastraTools.readWorkPaperSummary.execute({
      context: {
        range: 'Summary!A1:B5',
      },
    }),
    writeResult: mastraTools.setWorkPaperInputCell.execute({
      context: {
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      },
    }),
  },
  llamaIndex: {
    toolNames: llamaIndexTools.map((tool) => tool.name),
    readResult: requireLlamaIndexTool(llamaIndexTools, 'read_workpaper_summary').call({
      range: 'Summary!A1:B5',
    }),
    writeResult: requireLlamaIndexTool(llamaIndexTools, 'set_workpaper_input_cell').call({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    }),
  },
  langGraph: {
    nodeName: langGraphToolNode.nodeName,
    toolNames: langGraphToolNode.tools.map((tool) => tool.name),
    writeResult: requireWorkPaperWriteResult(
      requireLangGraphToolMessage(
        langGraphToolNode.invoke({
          messages: [
            {
              tool_calls: [
                {
                  id: 'call_set_input_b3',
                  name: 'set_workpaper_input_cell',
                  args: {
                    sheetName: 'Inputs',
                    address: 'B3',
                    value: 0.4,
                  },
                },
              ],
            },
          ],
        }).messages,
        'call_set_input_b3',
      ).result,
    ),
  },
  copilotKit: {
    actionNames: copilotKitActions.map((action) => action.name),
    readResult: requireCopilotKitAction(copilotKitActions, 'readWorkPaperSummary').handler({
      range: 'Summary!A1:B5',
    }),
    writeResult: requireCopilotKitAction(copilotKitActions, 'setWorkPaperInputCell').handler({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    }),
  },
  cloudflareAgents: {
    toolNames: Object.keys(cloudflareAgentTools),
    readResult: cloudflareAgentTools.readWorkPaperSummary.execute({
      range: 'Summary!A1:B5',
    }),
    writeResult: cloudflareAgentTools.setWorkPaperInputCell.execute({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    }),
  },
  crewAi: {
    toolNames: crewAiTools.map((tool) => tool.name),
    contract: {
      inputPayload: 'validated JSON args',
      formulaReadback: 'before/after computed Summary values',
      errorShape: '{ ok: false, error: string }',
    },
    writeResult: requireWorkPaperWriteResult(
      requireCrewAiTool(crewAiTools, 'set_workpaper_input_cell').run({
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      }),
    ),
  },
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

function createAiSdkTools(workPaperTools: WorkPaperToolSet) {
  return {
    readWorkPaperSummary: {
      description: 'Read computed WorkPaper summary values for a small range.',
      inputSchema: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            default: 'Summary!A1:B5',
          },
        },
      },
      execute({ range = 'Summary!A1:B5' }: ReadSummaryArgs = {}) {
        return workPaperTools.readWorkPaperSummary(range)
      },
    },

    setWorkPaperInputCell: {
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      inputSchema: {
        type: 'object',
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: {
            type: 'string',
          },
          address: {
            type: 'string',
          },
          value: {
            oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }],
          },
        },
      },
      execute(args: SetInputCellArgs) {
        return workPaperTools.setWorkPaperInputCell(args)
      },
    },
  }
}

function createOpenAiResponsesTools(workPaperTools: WorkPaperToolSet) {
  const functionTools: OpenAiFunctionTool[] = [
    {
      type: 'function',
      name: 'read_workpaper_summary',
      description: 'Read computed WorkPaper summary values for a small range.',
      parameters: {
        type: 'object',
        required: ['range'],
        properties: {
          range: {
            type: 'string',
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
          },
          address: {
            type: 'string',
          },
          value: {
            type: ['string', 'number', 'boolean', 'null'],
          },
        },
        additionalProperties: false,
      },
      strict: true,
    },
  ]

  return {
    tools: functionTools,
    dispatch(call: OpenAiFunctionCall): WorkPaperReadResult | WorkPaperWriteResult {
      if (call.name === 'read_workpaper_summary') {
        const args = readSummaryInputSchema.parse(JSON.parse(call.arguments))
        return workPaperTools.readWorkPaperSummary(args.range)
      }

      const args = setInputCellInputSchema.parse(JSON.parse(call.arguments))
      return workPaperTools.setWorkPaperInputCell(args)
    },
  }
}

function runOpenAiResponsesToolLoop(adapter: OpenAiResponsesAdapter) {
  const functionCalls: OpenAiFunctionCall[] = [
    {
      type: 'function_call',
      call_id: 'call_read_summary',
      name: 'read_workpaper_summary',
      arguments: JSON.stringify({ range: 'Summary!A1:B5' }),
    },
    {
      type: 'function_call',
      call_id: 'call_set_input_b3',
      name: 'set_workpaper_input_cell',
      arguments: JSON.stringify({
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      }),
    },
  ]
  const [readResult, writeResult] = functionCalls.map((call) => adapter.dispatch(call))
  const toolOutputs: OpenAiFunctionCallOutput[] = functionCalls.map((call, index) => ({
    type: 'function_call_output',
    call_id: call.call_id,
    output: JSON.stringify(index === 0 ? readResult : writeResult),
  }))

  return {
    toolNames: adapter.tools.map((tool) => tool.name),
    toolOutputTypes: toolOutputs.map((item) => item.type),
    readResult: requireWorkPaperReadResult(readResult),
    writeResult: requireWorkPaperWriteResult(writeResult),
  }
}

function runLangChainStructuredToolSmoke(tools: LangChainTool[]) {
  const readResult = requireTool(tools, 'read_workpaper_summary').invoke({
    range: 'Summary!A1:B5',
  })
  const writeResult = requireTool(tools, 'set_workpaper_input_cell').invoke({
    sheetName: 'Inputs',
    address: 'B3',
    value: 0.4,
  })

  return {
    toolNames: tools.map((tool) => tool.name),
    structuredFields: ['editedCell', 'before', 'after', 'checks'],
    readResult,
    writeResult,
  }
}

function createLangChainTools(workPaperTools: WorkPaperToolSet): LangChainTool[] {
  return [
    {
      name: 'read_workpaper_summary',
      description: 'Read computed WorkPaper summary values for a small range.',
      schema: readSummaryInputSchema,
      invoke({ range = 'Summary!A1:B5' }: ReadSummaryArgs = {}) {
        return workPaperTools.readWorkPaperSummary(readSummaryInputSchema.parse({ range }).range)
      },
    },
    {
      name: 'set_workpaper_input_cell',
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      schema: setInputCellInputSchema,
      invoke(args: SetInputCellArgs) {
        return workPaperTools.setWorkPaperInputCell(setInputCellInputSchema.parse(args))
      },
    },
  ]
}

function createMastraTools(workPaperTools: WorkPaperToolSet) {
  return {
    readWorkPaperSummary: {
      id: 'read-workpaper-summary',
      description: 'Read computed WorkPaper summary values for a small range.',
      inputSchema: readSummaryInputSchema,
      outputSchema: workPaperSummaryOutputSchema,
      execute({ context }: { context?: ReadSummaryArgs } = {}) {
        return workPaperTools.readWorkPaperSummary(readSummaryInputSchema.parse(context ?? {}).range)
      },
    },
    setWorkPaperInputCell: {
      id: 'set-workpaper-input-cell',
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      inputSchema: setInputCellInputSchema,
      outputSchema: workPaperWriteOutputSchema,
      execute({ context }: { context: SetInputCellArgs }) {
        return workPaperTools.setWorkPaperInputCell(setInputCellInputSchema.parse(context))
      },
    },
  }
}

function createLlamaIndexTools(workPaperTools: WorkPaperToolSet): LlamaIndexTool[] {
  return [
    {
      name: 'read_workpaper_summary',
      description: 'Read computed WorkPaper summary values for a small range.',
      parameters: readSummaryInputSchema,
      call(args: ReadSummaryArgs = {}) {
        return workPaperTools.readWorkPaperSummary(readSummaryInputSchema.parse(args).range)
      },
    },
    {
      name: 'set_workpaper_input_cell',
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      parameters: setInputCellInputSchema,
      call(args: SetInputCellArgs) {
        return workPaperTools.setWorkPaperInputCell(setInputCellInputSchema.parse(args))
      },
    },
  ]
}

function createLangGraphToolNode(workPaperTools: WorkPaperToolSet): LangGraphToolNode {
  const tools = createLangChainTools(workPaperTools)
  return {
    nodeName: 'tools',
    tools,
    invoke(input) {
      const lastMessage = input.messages[input.messages.length - 1]
      const toolCalls = lastMessage?.tool_calls ?? []
      return {
        messages: toolCalls.map((toolCall) => {
          if (toolCall.name === 'read_workpaper_summary') {
            const tool = requireTool(tools, 'read_workpaper_summary')
            const result = tool.invoke(toolCall.args)
            return {
              type: 'tool',
              name: tool.name,
              tool_call_id: toolCall.id,
              content: JSON.stringify(result),
              result,
            }
          }
          const tool = requireTool(tools, 'set_workpaper_input_cell')
          const result = tool.invoke(toolCall.args)
          return {
            type: 'tool',
            name: tool.name,
            tool_call_id: toolCall.id,
            content: JSON.stringify(result),
            result,
          }
        }),
      }
    },
  }
}

function createCopilotKitActions(workPaperTools: WorkPaperToolSet): CopilotKitAction[] {
  return [
    {
      name: 'readWorkPaperSummary',
      description: 'Read computed WorkPaper summary values for a small range.',
      parameters: [
        {
          name: 'range',
          type: 'string',
          description: 'A1-style summary range such as Summary!A1:B5.',
        },
      ],
      handler(args: ReadSummaryArgs = {}) {
        return workPaperTools.readWorkPaperSummary(readSummaryInputSchema.parse(args).range)
      },
    },
    {
      name: 'setWorkPaperInputCell',
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      parameters: [
        {
          name: 'sheetName',
          type: 'string',
          description: 'The editable sheet. This example allows Inputs only.',
          required: true,
        },
        {
          name: 'address',
          type: 'string',
          description: 'A1-style input address such as B3.',
          required: true,
        },
        {
          name: 'value',
          type: 'number',
          description: 'The replacement value for the input cell.',
          required: true,
        },
      ],
      handler(args: SetInputCellArgs) {
        return workPaperTools.setWorkPaperInputCell(setInputCellInputSchema.parse(args))
      },
    },
  ]
}

function createCloudflareAgentTools(workPaperTools: WorkPaperToolSet) {
  return {
    readWorkPaperSummary: {
      description: 'Read computed WorkPaper summary values inside a Cloudflare Agent turn.',
      inputSchema: readSummaryInputSchema,
      execute(args: ReadSummaryArgs = {}) {
        return workPaperTools.readWorkPaperSummary(readSummaryInputSchema.parse(args).range)
      },
    },
    setWorkPaperInputCell: {
      description: 'Set one validated WorkPaper input cell inside a Cloudflare Agent turn.',
      inputSchema: setInputCellInputSchema,
      execute(args: SetInputCellArgs) {
        return workPaperTools.setWorkPaperInputCell(setInputCellInputSchema.parse(args))
      },
    },
  }
}

function createCrewAiTools(workPaperTools: WorkPaperToolSet): CrewAiTool[] {
  return [
    {
      name: 'read_workpaper_summary',
      description: 'Read computed WorkPaper summary values for a CrewAI tool call.',
      argsSchema: readSummaryInputSchema,
      run(args: ReadSummaryArgs = {}) {
        return workPaperTools.readWorkPaperSummary(readSummaryInputSchema.parse(args).range)
      },
    },
    {
      name: 'set_workpaper_input_cell',
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      argsSchema: setInputCellInputSchema,
      run(args: SetInputCellArgs) {
        return workPaperTools.setWorkPaperInputCell(setInputCellInputSchema.parse(args))
      },
    },
  ]
}

function createWorkPaperTools(workbook: WorkPaperInstance) {
  const summarySheet = requireSheet(workbook, 'Summary')

  return {
    readWorkPaperSummary(range = 'Summary!A1:B5') {
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

    setWorkPaperInputCell({ sheetName, address, value }: SetInputCellArgs) {
      if (sheetName !== 'Inputs') {
        throw new Error(`This example only permits Inputs edits, received ${sheetName}`)
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
        before,
        after,
        restored: restoredSummary,
        formulaContracts,
        checks: {
          previousValue,
          newValue: workbook.getCellSerialized(target),
          formulasPersisted: sameJson(formulaContracts, restoredFormulaContracts),
          restoredMatchesAfter: sameJson(after, restoredSummary),
          expectedArrChanged: after.expectedArr > before.expectedArr,
          serializedBytes: Buffer.byteLength(serialized, 'utf8'),
        },
      }
    },
  }
}

function requireTool(tools: LangChainTool[], name: 'read_workpaper_summary'): LangChainReadTool
function requireTool(tools: LangChainTool[], name: 'set_workpaper_input_cell'): LangChainWriteTool
function requireTool(tools: LangChainTool[], name: LangChainTool['name']): LangChainTool
function requireTool(tools: LangChainTool[], name: LangChainTool['name']): LangChainTool {
  const tool = tools.find((candidate) => candidate.name === name)
  if (tool === undefined) {
    throw new Error(`Missing framework tool: ${name}`)
  }
  return tool
}

function requireLlamaIndexTool(
  tools: LlamaIndexTool[],
  name: 'read_workpaper_summary',
): Extract<LlamaIndexTool, { name: 'read_workpaper_summary' }>
function requireLlamaIndexTool(
  tools: LlamaIndexTool[],
  name: 'set_workpaper_input_cell',
): Extract<LlamaIndexTool, { name: 'set_workpaper_input_cell' }>
function requireLlamaIndexTool(tools: LlamaIndexTool[], name: LlamaIndexTool['name']): LlamaIndexTool {
  const tool = tools.find((candidate) => candidate.name === name)
  if (tool === undefined) {
    throw new Error(`Missing LlamaIndex tool: ${name}`)
  }
  return tool
}

function requireCopilotKitAction(
  actions: CopilotKitAction[],
  name: 'readWorkPaperSummary',
): Extract<CopilotKitAction, { name: 'readWorkPaperSummary' }>
function requireCopilotKitAction(
  actions: CopilotKitAction[],
  name: 'setWorkPaperInputCell',
): Extract<CopilotKitAction, { name: 'setWorkPaperInputCell' }>
function requireCopilotKitAction(actions: CopilotKitAction[], name: CopilotKitAction['name']): CopilotKitAction {
  const action = actions.find((candidate) => candidate.name === name)
  if (action === undefined) {
    throw new Error(`Missing CopilotKit action: ${name}`)
  }
  return action
}

function requireCrewAiTool(tools: CrewAiTool[], name: 'read_workpaper_summary'): Extract<CrewAiTool, { name: 'read_workpaper_summary' }>
function requireCrewAiTool(tools: CrewAiTool[], name: 'set_workpaper_input_cell'): Extract<CrewAiTool, { name: 'set_workpaper_input_cell' }>
function requireCrewAiTool(tools: CrewAiTool[], name: CrewAiTool['name']): CrewAiTool {
  const tool = tools.find((candidate) => candidate.name === name)
  if (tool === undefined) {
    throw new Error(`Missing CrewAI tool: ${name}`)
  }
  return tool
}

function requireLangGraphToolMessage(messages: LangGraphToolMessage[], toolCallId: string): LangGraphToolMessage {
  const message = messages.find((candidate) => candidate.tool_call_id === toolCallId)
  if (message === undefined) {
    throw new Error(`Missing LangGraph tool message: ${toolCallId}`)
  }
  return message
}

function requireWorkPaperWriteResult(result: WorkPaperReadResult | WorkPaperWriteResult): WorkPaperWriteResult {
  if (!('editedCell' in result)) {
    throw new Error(`Expected WorkPaper write result, received ${JSON.stringify(result)}`)
  }
  return result
}

function requireWorkPaperReadResult(result: WorkPaperReadResult | WorkPaperWriteResult): WorkPaperReadResult {
  if ('editedCell' in result) {
    throw new Error(`Expected WorkPaper read result, received ${JSON.stringify(result)}`)
  }
  return result
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

  const writeResults: [string, WorkPaperWriteResult][] = [
    ['aiSdk', actual.aiSdk.writeResult],
    ['openAiResponses', actual.openAiResponses.writeResult],
    ['langChain', actual.langChain.writeResult],
    ['mastra', actual.mastra.writeResult],
    ['llamaIndex', actual.llamaIndex.writeResult],
    ['langGraph', actual.langGraph.writeResult],
    ['copilotKit', actual.copilotKit.writeResult],
    ['cloudflareAgents', actual.cloudflareAgents.writeResult],
    ['crewAi', actual.crewAi.writeResult],
  ]

  for (const [framework, result] of writeResults) {
    if (
      result.editedCell !== 'Inputs!B3' ||
      !sameJson(result.before, expectedBefore) ||
      !sameJson(result.after, expectedAfter) ||
      !sameJson(result.restored, expectedAfter) ||
      !sameJson(result.formulaContracts, expectedFormulaContracts) ||
      result.checks.previousValue !== 0.25 ||
      result.checks.newValue !== 0.4 ||
      !result.checks.formulasPersisted ||
      !result.checks.restoredMatchesAfter ||
      !result.checks.expectedArrChanged ||
      result.checks.serializedBytes <= 0
    ) {
      throw new Error(`Unexpected ${framework} adapter result: ${JSON.stringify(result)}`)
    }
  }
}
