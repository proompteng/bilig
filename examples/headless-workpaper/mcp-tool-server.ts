import { pathToFileURL } from 'node:url'

import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  type RawCellContent,
  type WorkPaperCellAddress,
} from '@bilig/workpaper'

type JsonObject = Record<string, unknown>
type JsonRpcId = string | number | null

interface JsonRpcRequest {
  jsonrpc: '2.0'
  id?: JsonRpcId
  method: string
  params?: JsonObject
}

interface JsonRpcSuccess<Result> {
  jsonrpc: '2.0'
  id: JsonRpcId | undefined
  result: Result
}

interface McpCapabilities {
  tools: {
    listChanged: false
  }
}

interface McpToolDefinition {
  name: 'read_workpaper_summary' | 'set_workpaper_input_cell'
  title: string
  description: string
  inputSchema: JsonObject
  annotations: McpToolAnnotations
}

interface McpToolAnnotations {
  title: string
  readOnlyHint: boolean
  destructiveHint: boolean
  idempotentHint: boolean
  openWorldHint: false
}

interface McpToolsListResult {
  tools: McpToolDefinition[]
}

interface McpToolCallResult {
  content: {
    type: 'text'
    text: string
  }[]
  structuredContent: WorkPaperSummaryReadback | WorkPaperInputEditReadback
  isError: false
}

type McpJsonRpcResponse = JsonRpcSuccess<McpToolsListResult | McpToolCallResult>

interface McpWorkPaperToolServer {
  capabilities: McpCapabilities
  handleJsonRpc(request: unknown): McpJsonRpcResponse
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
  sheetName: 'Inputs'
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

interface McpDemoOutput {
  capabilities: McpCapabilities
  listResponse: JsonRpcSuccess<McpToolsListResult>
  readResponse: JsonRpcSuccess<McpToolCallResult>
  writeResponse: JsonRpcSuccess<McpToolCallResult>
}

if (isMainModule()) {
  const output = createMcpDemoOutput()
  assertOutput(output)
  console.log(JSON.stringify(output, null, 2))
}

function createMcpDemoOutput(): McpDemoOutput {
  const server = createMcpWorkPaperToolServer(buildWorkbook())
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

function buildWorkbook(): WorkPaper {
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

function createMcpWorkPaperToolServer(workbook: WorkPaper): McpWorkPaperToolServer {
  const workPaperTools = createWorkPaperTools(workbook)
  const toolDefinitions: McpToolDefinition[] = [
    {
      name: 'read_workpaper_summary',
      title: 'Read WorkPaper Summary',
      description: 'Read computed WorkPaper summary values for a small range.',
      inputSchema: {
        type: 'object',
        required: [],
        properties: {
          range: {
            type: 'string',
            description: 'A1 range with an optional sheet name.',
            default: 'Summary!A1:B5',
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
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      inputSchema: {
        type: 'object',
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: {
            type: 'string',
            const: 'Inputs',
          },
          address: {
            type: 'string',
            description: 'A1 cell address in the Inputs sheet.',
          },
          value: {
            type: ['string', 'number', 'boolean', 'null'],
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

    handleJsonRpc(request: unknown): McpJsonRpcResponse {
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
  const toolName = parsedParams.name
  const args = parsedParams.arguments

  if (toolName === 'read_workpaper_summary') {
    const readArgs = requireRecord(args ?? {}, 'read_workpaper_summary arguments')
    const range = readArgs.range
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
  if (request.jsonrpc !== '2.0' || typeof request.method !== 'string') {
    throw new Error('Expected JSON-RPC 2.0 request')
  }

  const id = request.id
  if (id !== undefined && id !== null && typeof id !== 'string' && typeof id !== 'number') {
    throw new Error(`Unsupported JSON-RPC id: ${JSON.stringify(id)}`)
  }

  return {
    jsonrpc: '2.0',
    id,
    method: request.method,
    params: request.params === undefined ? undefined : requireRecord(request.params, 'JSON-RPC params'),
  }
}

function parseInputCellArgs(value: unknown): WorkPaperInputCellArgs {
  const args = requireRecord(value, 'set_workpaper_input_cell arguments')
  const sheetName = requireString(args.sheetName, 'sheetName')
  if (sheetName !== 'Inputs') {
    throw new Error(`This example only permits Inputs edits, received ${sheetName}`)
  }

  const cellValue = args.value
  if (cellValue !== null && typeof cellValue !== 'string' && typeof cellValue !== 'number' && typeof cellValue !== 'boolean') {
    throw new Error(`Unsupported cell value: ${JSON.stringify(cellValue)}`)
  }

  return {
    sheetName,
    address: requireString(args.address, 'address'),
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
  if (typeof cell !== 'object' || cell === null || !('value' in cell) || typeof cell.value !== 'number') {
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

function requireToolsListResponse(response: McpJsonRpcResponse): JsonRpcSuccess<McpToolsListResult> {
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

function requireToolCallResponse(response: McpJsonRpcResponse, toolName: McpToolDefinition['name']): JsonRpcSuccess<McpToolCallResult> {
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

function isToolsListResult(result: McpToolsListResult | McpToolCallResult): result is McpToolsListResult {
  return isRecord(result) && Array.isArray(result['tools'])
}

function isToolCallResult(result: McpToolsListResult | McpToolCallResult): result is McpToolCallResult {
  return isRecord(result) && Array.isArray(result['content']) && result['isError'] === false && isRecord(result['structuredContent'])
}

function assertOutput(actual: McpDemoOutput): void {
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

function isMainModule(): boolean {
  return process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href
}

export {
  assertOutput,
  buildWorkbook,
  createMcpDemoOutput,
  createMcpWorkPaperToolServer,
  type JsonRpcId,
  type JsonRpcRequest,
  type McpCapabilities,
  type McpJsonRpcResponse,
}
