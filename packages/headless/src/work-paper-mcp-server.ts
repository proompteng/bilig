import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import { WorkPaper } from './work-paper.js'
import type { RawCellContent, WorkPaperCellAddress } from './work-paper-types.js'
import { ValueTag } from '@bilig/protocol'

type JsonObject = Record<string, unknown>
type JsonRpcId = string | number | null

interface JsonRpcRequest {
  jsonrpc: '2.0'
  id: JsonRpcId | undefined
  method: string
  params?: JsonObject
}

interface JsonRpcSuccess<Result> {
  jsonrpc: '2.0'
  id: JsonRpcId | undefined
  result: Result
}

interface WorkPaperMcpCapabilities {
  tools: {
    listChanged: false
  }
  resources?: {
    listChanged: false
    subscribe?: false
  }
  prompts?: {
    listChanged: false
  }
}

interface WorkPaperMcpToolDefinition {
  name: 'read_workpaper_summary' | 'set_workpaper_input_cell'
  title: string
  description: string
  inputSchema: JsonObject
  outputSchema: JsonObject
  annotations: WorkPaperMcpToolAnnotations
}

interface WorkPaperMcpToolAnnotations {
  title: string
  readOnlyHint: boolean
  destructiveHint: boolean
  idempotentHint: boolean
  openWorldHint: false
}

interface WorkPaperMcpToolsListResult {
  tools: WorkPaperMcpToolDefinition[]
}

interface WorkPaperMcpToolCallResult {
  content: {
    type: 'text'
    text: string
  }[]
  structuredContent: WorkPaperSummaryReadback | WorkPaperInputEditReadback
  isError: false
}

type WorkPaperMcpJsonRpcResponse = JsonRpcSuccess<unknown>

interface WorkPaperMcpToolServer {
  capabilities: WorkPaperMcpCapabilities
  handleJsonRpc(request: unknown): WorkPaperMcpJsonRpcResponse
}

interface WorkPaperSummary {
  expectedCustomers: number
  expectedArr: number
  expansionArr: number
  targetGap: number
}

interface WorkPaperFormulaContracts {
  expectedCustomers: string
  expectedArr: string
  expansionArr: string
  targetGap: string
}

interface WorkPaperSummaryReadback {
  range: string
  values: unknown[][]
  serialized: RawCellContent[][]
}

interface WorkPaperInputCellArgs {
  sheetName: string
  address: string
  value: RawCellContent
}

interface WorkPaperInputEditReadback {
  editedCell: string
  before: WorkPaperSummary
  after: WorkPaperSummary
  restored: WorkPaperSummary
  formulaContracts: WorkPaperFormulaContracts
  checks: {
    previousValue: RawCellContent
    newValue: RawCellContent
    formulasPersisted: boolean
    restoredMatchesAfter: boolean
    expectedArrChanged: boolean
    serializedBytes: number
  }
}

interface WorkPaperMcpDemoOutput {
  capabilities: WorkPaperMcpCapabilities
  listResponse: JsonRpcSuccess<WorkPaperMcpToolsListResult>
  readResponse: JsonRpcSuccess<WorkPaperMcpToolCallResult>
  writeResponse: JsonRpcSuccess<WorkPaperMcpToolCallResult>
}

function createWorkPaperMcpDemoOutput(): WorkPaperMcpDemoOutput {
  const server = createWorkPaperMcpToolServer(buildDemoWorkPaper())
  const listResponse = server.handleJsonRpc({
    jsonrpc: '2.0',
    id: 1,
    method: 'tools/list',
  })
  const readResponse = server.handleJsonRpc({
    jsonrpc: '2.0',
    id: 2,
    method: 'tools/call',
    params: {
      name: 'read_workpaper_summary',
      arguments: {
        range: 'Summary!A1:B5',
      },
    },
  })
  const writeResponse = server.handleJsonRpc({
    jsonrpc: '2.0',
    id: 3,
    method: 'tools/call',
    params: {
      name: 'set_workpaper_input_cell',
      arguments: {
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      },
    },
  })

  return {
    capabilities: server.capabilities,
    listResponse: requireToolsListResponse(listResponse),
    readResponse: requireToolCallResponse(readResponse, 'read_workpaper_summary'),
    writeResponse: requireToolCallResponse(writeResponse, 'set_workpaper_input_cell'),
  }
}

function buildDemoWorkPaper(): WorkPaper {
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

function createWorkPaperMcpToolServer(workbook: WorkPaper): WorkPaperMcpToolServer {
  const workPaperTools = createWorkPaperTools(workbook)
  const toolDefinitions: WorkPaperMcpToolDefinition[] = [
    {
      name: 'read_workpaper_summary',
      title: 'Read WorkPaper Summary',
      description:
        'Read calculated demo WorkPaper summary values and serialized formula contents. Use this read-only tool to verify workbook formulas without opening Excel.',
      inputSchema: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            description: 'A1 range with an optional sheet name. Defaults to Summary!A1:B5.',
            default: 'Summary!A1:B5',
          },
        },
        additionalProperties: false,
      },
      outputSchema: {
        type: 'object',
        required: ['range', 'values', 'serialized'],
        properties: {
          range: {
            type: 'string',
          },
          values: {
            type: 'array',
            description: 'Two-dimensional array of calculated values.',
          },
          serialized: {
            type: 'array',
            description: 'Two-dimensional array of serialized literals and formulas.',
          },
        },
        additionalProperties: false,
      },
      annotations: {
        title: 'Read WorkPaper Summary',
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    {
      name: 'set_workpaper_input_cell',
      title: 'Set WorkPaper Input Cell',
      description:
        'Set one demo Inputs cell, recalculate dependent formulas, then return before/after/restored readback. Use only for the packaged demo workbook.',
      inputSchema: {
        type: 'object',
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: {
            type: 'string',
            const: 'Inputs',
            description: 'Must be Inputs for the packaged demo workbook.',
          },
          address: {
            type: 'string',
            description: 'Single A1 cell address in the Inputs sheet, for example B3.',
          },
          value: {
            type: ['string', 'number', 'boolean', 'null'],
            description: 'Raw replacement value. Formula strings must start with =.',
          },
        },
        additionalProperties: false,
      },
      outputSchema: {
        type: 'object',
        required: ['editedCell', 'before', 'after', 'restored', 'formulaContracts', 'checks'],
        properties: {
          editedCell: {
            type: 'string',
          },
          before: workPaperSummaryOutputSchema(),
          after: workPaperSummaryOutputSchema(),
          restored: workPaperSummaryOutputSchema(),
          formulaContracts: {
            type: 'object',
            required: ['expectedCustomers', 'expectedArr', 'expansionArr', 'targetGap'],
            properties: {
              expectedCustomers: {
                type: 'string',
              },
              expectedArr: {
                type: 'string',
              },
              expansionArr: {
                type: 'string',
              },
              targetGap: {
                type: 'string',
              },
            },
            additionalProperties: false,
          },
          checks: {
            type: 'object',
            required: ['previousValue', 'newValue', 'formulasPersisted', 'restoredMatchesAfter', 'expectedArrChanged', 'serializedBytes'],
            properties: {
              previousValue: rawCellContentSchema(),
              newValue: rawCellContentSchema(),
              formulasPersisted: {
                type: 'boolean',
              },
              restoredMatchesAfter: {
                type: 'boolean',
              },
              expectedArrChanged: {
                type: 'boolean',
              },
              serializedBytes: {
                type: 'number',
              },
            },
            additionalProperties: false,
          },
        },
        additionalProperties: false,
      },
      annotations: {
        title: 'Set WorkPaper Input Cell',
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
  ]

  return {
    capabilities: {
      tools: {
        listChanged: false,
      },
    },

    handleJsonRpc(request: unknown): WorkPaperMcpJsonRpcResponse {
      const parsedRequest = parseJsonRpcRequest(request)

      if (parsedRequest.method === 'tools/list') {
        return {
          jsonrpc: '2.0',
          id: parsedRequest.id,
          result: {
            tools: toolDefinitions,
          },
        }
      }

      if (parsedRequest.method === 'tools/call') {
        const structuredContent = callTool(workPaperTools, parsedRequest.params)

        return {
          jsonrpc: '2.0',
          id: parsedRequest.id,
          result: {
            content: [
              {
                type: 'text',
                text: JSON.stringify(structuredContent),
              },
            ],
            structuredContent,
            isError: false,
          },
        }
      }

      throw new Error(`Unsupported MCP method: ${parsedRequest.method}`)
    },
  }
}

function callTool(
  workPaperTools: ReturnType<typeof createWorkPaperTools>,
  params: JsonObject | undefined,
): WorkPaperSummaryReadback | WorkPaperInputEditReadback {
  const parsedParams = requireRecord(params ?? {}, 'MCP tool call params')
  const toolName = parsedParams['name']
  const args = parsedParams['arguments']

  if (toolName === 'read_workpaper_summary') {
    const readArgs = requireRecord(args ?? {}, 'read_workpaper_summary arguments')
    const range = readArgs['range']
    return workPaperTools.readWorkPaperSummary(range === undefined ? undefined : requireString(range, 'range'))
  }

  if (toolName === 'set_workpaper_input_cell') {
    return workPaperTools.setWorkPaperInputCell(parseInputCellArgs(args))
  }

  throw new Error(`Unknown WorkPaper tool: ${String(toolName)}`)
}

function createWorkPaperTools(workbook: WorkPaper): {
  readWorkPaperSummary(range?: string): WorkPaperSummaryReadback
  setWorkPaperInputCell(args: WorkPaperInputCellArgs): WorkPaperInputEditReadback
} {
  const summarySheet = requireSheet(workbook, 'Summary')

  return {
    readWorkPaperSummary(range = 'Summary!A1:B5'): WorkPaperSummaryReadback {
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

    setWorkPaperInputCell({ sheetName, address, value }: WorkPaperInputCellArgs): WorkPaperInputEditReadback {
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

function parseJsonRpcRequest(value: unknown): JsonRpcRequest {
  const request = requireRecord(value, 'JSON-RPC request')
  if (request['jsonrpc'] !== '2.0' || typeof request['method'] !== 'string') {
    throw new Error('Expected JSON-RPC 2.0 request')
  }

  const id = request['id']
  if (id !== undefined && id !== null && typeof id !== 'string' && typeof id !== 'number') {
    throw new Error(`Unsupported JSON-RPC id: ${JSON.stringify(id)}`)
  }

  const parsed: JsonRpcRequest = {
    jsonrpc: '2.0',
    id,
    method: request['method'],
  }
  const params = request['params']
  if (params !== undefined) {
    parsed.params = requireRecord(params, 'JSON-RPC params')
  }
  return parsed
}

function parseInputCellArgs(value: unknown): WorkPaperInputCellArgs {
  const args = requireRecord(value, 'set_workpaper_input_cell arguments')
  const sheetName = requireString(args['sheetName'], 'sheetName')
  if (sheetName !== 'Inputs') {
    throw new Error(`This example only permits Inputs edits, received ${sheetName}`)
  }

  const cellValue = args['value']
  if (cellValue !== null && typeof cellValue !== 'string' && typeof cellValue !== 'number' && typeof cellValue !== 'boolean') {
    throw new Error(`Unsupported cell value: ${JSON.stringify(cellValue)}`)
  }

  return {
    sheetName,
    address: requireString(args['address'], 'address'),
    value: cellValue,
  }
}

function requireRecord(value: unknown, label: string): JsonObject {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object`)
  }
  return value
}

function requireString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${label} to be a string`)
  }
  return value
}

function requireSheet(workpaper: WorkPaper, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function requireCellAddress(workpaper: WorkPaper, sheetName: string, a1Address: string): WorkPaperCellAddress {
  const sheetId = requireSheet(workpaper, sheetName)
  const parsed = workpaper.simpleCellAddressFromString(a1Address, sheetId)

  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid cell address: ${sheetName}!${a1Address}`)
  }

  return parsed
}

function readSummary(workpaper: WorkPaper, summary: number): WorkPaperSummary {
  return {
    expectedCustomers: readNumber(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readNumber(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readNumber(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readNumber(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readFormulaContracts(workpaper: WorkPaper, summary: number): WorkPaperFormulaContracts {
  return {
    expectedCustomers: readFormula(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readFormula(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readFormula(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readFormula(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readNumber(workpaper: WorkPaper, sheet: number, row: number, col: number, label: string): number {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (cell.tag !== ValueTag.Number) {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readFormula(workpaper: WorkPaper, sheet: number, row: number, col: number, label: string): string {
  const formula = workpaper.getCellFormula({ sheet, row, col })
  if (formula === undefined) {
    throw new Error(`Expected ${label} to be a formula`)
  }
  return formula
}

function workPaperSummaryOutputSchema(): JsonObject {
  return {
    type: 'object',
    required: ['expectedCustomers', 'expectedArr', 'expansionArr', 'targetGap'],
    properties: {
      expectedCustomers: {
        type: 'number',
      },
      expectedArr: {
        type: 'number',
      },
      expansionArr: {
        type: 'number',
      },
      targetGap: {
        type: 'number',
      },
    },
    additionalProperties: false,
  }
}

function rawCellContentSchema(): JsonObject {
  return {
    type: ['string', 'number', 'boolean', 'null'],
    description: 'Raw serialized cell content; formulas are strings that start with =.',
  }
}

function serializeWorkbook(workpaper: WorkPaper): string {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workpaper, {
      includeConfig: true,
    }),
  )
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function isRecord(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function requireToolsListResponse(response: WorkPaperMcpJsonRpcResponse): JsonRpcSuccess<WorkPaperMcpToolsListResult> {
  const result = response.result
  if (!isToolsListResult(result)) {
    throw new Error(`Expected tools/list response, received ${JSON.stringify(response)}`)
  }

  return {
    jsonrpc: response.jsonrpc,
    id: response.id,
    result,
  }
}

function requireToolCallResponse(
  response: WorkPaperMcpJsonRpcResponse,
  toolName: WorkPaperMcpToolDefinition['name'],
): JsonRpcSuccess<WorkPaperMcpToolCallResult> {
  const result = response.result
  if (!isToolCallResult(result)) {
    throw new Error(`Expected ${toolName} tool-call response, received ${JSON.stringify(response)}`)
  }

  return {
    jsonrpc: response.jsonrpc,
    id: response.id,
    result,
  }
}

function isToolsListResult(result: unknown): result is WorkPaperMcpToolsListResult {
  return isRecord(result) && Array.isArray(result['tools'])
}

function isToolCallResult(result: unknown): result is WorkPaperMcpToolCallResult {
  return isRecord(result) && Array.isArray(result['content']) && result['isError'] === false && isRecord(result['structuredContent'])
}

function assertWorkPaperMcpDemoOutput(actual: WorkPaperMcpDemoOutput): void {
  const toolNames = actual.listResponse.result.tools.map((tool) => tool.name)
  if (!sameJson(toolNames, ['read_workpaper_summary', 'set_workpaper_input_cell'])) {
    throw new Error(`Unexpected MCP tool list: ${JSON.stringify(toolNames)}`)
  }

  const writeResult = actual.writeResponse.result.structuredContent
  if (!isInputEditReadback(writeResult)) {
    throw new Error(`Unexpected MCP write result: ${JSON.stringify(writeResult)}`)
  }

  const expectedBefore: WorkPaperSummary = {
    expectedCustomers: 5,
    expectedArr: 60000,
    expansionArr: 66000,
    targetGap: -34000,
  }
  const expectedAfter: WorkPaperSummary = {
    expectedCustomers: 8,
    expectedArr: 96000,
    expansionArr: 105600,
    targetGap: 5600,
  }
  const expectedFormulaContracts: WorkPaperFormulaContracts = {
    expectedCustomers: '=Inputs!B2*Inputs!B3',
    expectedArr: '=B2*Inputs!B4',
    expansionArr: '=B3*Inputs!B5',
    targetGap: '=B4-100000',
  }

  if (
    actual.capabilities.tools.listChanged ||
    actual.readResponse.result.content[0]?.type !== 'text' ||
    actual.writeResponse.result.content[0]?.type !== 'text' ||
    writeResult.editedCell !== 'Inputs!B3' ||
    !sameJson(writeResult.before, expectedBefore) ||
    !sameJson(writeResult.after, expectedAfter) ||
    !sameJson(writeResult.restored, expectedAfter) ||
    !sameJson(writeResult.formulaContracts, expectedFormulaContracts) ||
    writeResult.checks.previousValue !== 0.25 ||
    writeResult.checks.newValue !== 0.4 ||
    !writeResult.checks.formulasPersisted ||
    !writeResult.checks.restoredMatchesAfter ||
    !writeResult.checks.expectedArrChanged ||
    writeResult.checks.serializedBytes <= 0
  ) {
    throw new Error(`Unexpected MCP adapter result: ${JSON.stringify(actual)}`)
  }
}

function isInputEditReadback(value: WorkPaperSummaryReadback | WorkPaperInputEditReadback): value is WorkPaperInputEditReadback {
  return 'editedCell' in value
}

export {
  assertWorkPaperMcpDemoOutput,
  buildDemoWorkPaper,
  createWorkPaperMcpDemoOutput,
  createWorkPaperMcpToolServer,
  type JsonRpcId,
  type JsonRpcRequest,
  type WorkPaperInputEditReadback,
  type WorkPaperMcpCapabilities,
  type WorkPaperMcpJsonRpcResponse,
  type WorkPaperMcpToolCallResult,
  type WorkPaperMcpToolDefinition,
  type WorkPaperMcpToolServer,
  type WorkPaperMcpToolsListResult,
  type WorkPaperSummaryReadback,
}
