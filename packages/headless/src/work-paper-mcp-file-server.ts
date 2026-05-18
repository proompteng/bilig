import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  type PersistedWorkPaperDocument,
} from './persistence.js'
import type { WorkPaper } from './work-paper.js'
import type { WorkPaperMcpCapabilities, WorkPaperMcpToolServer } from './work-paper-mcp-server.js'
import { buildDemoWorkPaper } from './work-paper-mcp-server.js'
import type { RawCellContent, WorkPaperCellAddress } from './work-paper-types.js'
import { formatCellDisplayValue } from '@bilig/protocol'
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { basename, dirname, resolve } from 'node:path'

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

interface FileBackedWorkPaperMcpOptions {
  workbook: WorkPaper
  writable?: boolean
  sourcePath?: string
  persist?: (workbook: WorkPaper) => FileBackedWorkPaperPersistResult
}

interface FileBackedWorkPaperPersistResult {
  persisted: boolean
  path?: string
  serializedBytes: number
}

interface WorkPaperMcpToolAnnotations {
  title: string
  readOnlyHint: boolean
  destructiveHint: boolean
  idempotentHint: boolean
  openWorldHint: false
}

interface WorkPaperMcpToolDefinition {
  name:
    | 'list_sheets'
    | 'read_range'
    | 'read_cell'
    | 'set_cell_contents'
    | 'get_cell_display_value'
    | 'export_workpaper_document'
    | 'validate_formula'
  title: string
  description: string
  inputSchema: JsonObject
  outputSchema: JsonObject
  annotations: WorkPaperMcpToolAnnotations
}

interface FileBackedToolCallResult {
  content: {
    type: 'text'
    text: string
  }[]
  structuredContent: JsonObject
  isError: false
}

const capabilities: WorkPaperMcpCapabilities = {
  tools: {
    listChanged: false,
  },
}

function createFileBackedWorkPaperMcpToolServer(options: FileBackedWorkPaperMcpOptions): WorkPaperMcpToolServer {
  const { workbook, writable = false, sourcePath } = options
  const persist = options.persist ?? createMemoryPersist(workbook)
  const toolDefinitions = createFileBackedToolDefinitions(writable)

  return {
    capabilities,

    handleJsonRpc(request: unknown): JsonRpcSuccess<unknown> {
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
        const structuredContent = callFileBackedTool({
          workbook,
          writable,
          sourcePath,
          persist,
          params: parsedRequest.params,
        })

        return {
          jsonrpc: '2.0',
          id: parsedRequest.id,
          result: toolResult(structuredContent),
        }
      }

      throw new Error(`Unsupported MCP method: ${parsedRequest.method}`)
    },
  }
}

function createFileBackedWorkPaperMcpToolServerFromFile(input: {
  workpaperPath: string
  writable?: boolean
  initDemoWorkPaper?: boolean
}): WorkPaperMcpToolServer {
  const workpaperPath = resolve(input.workpaperPath)
  if (input.initDemoWorkPaper && !existsSync(workpaperPath)) {
    mkdirSync(dirname(workpaperPath), { recursive: true })
    writeFileAtomically(workpaperPath, serializeWorkPaperDocument(exportWorkPaperDocument(buildDemoWorkPaper(), { includeConfig: true })))
  }
  const workbook = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(workpaperPath, 'utf8')))

  return createFileBackedWorkPaperMcpToolServer({
    workbook,
    writable: input.writable ?? false,
    sourcePath: workpaperPath,
    persist(updatedWorkbook) {
      const serialized = serializeWorkbook(updatedWorkbook)
      if (input.writable) {
        writeFileAtomically(workpaperPath, serialized)
      }
      return {
        persisted: input.writable ?? false,
        path: workpaperPath,
        serializedBytes: Buffer.byteLength(serialized, 'utf8'),
      }
    },
  })
}

function createFileBackedToolDefinitions(writable: boolean): WorkPaperMcpToolDefinition[] {
  return [
    {
      name: 'list_sheets',
      title: 'List WorkPaper Sheets',
      description:
        'Discover sheet names and used dimensions before reading or editing a WorkPaper. Returns metadata only; use read_range or read_cell for values.',
      inputSchema: emptySchema(),
      outputSchema: {
        type: 'object',
        required: ['writable', 'sheets'],
        properties: {
          sourcePath: {
            type: 'string',
            description: 'Absolute JSON file path when the server was started with --workpaper.',
          },
          writable: {
            type: 'boolean',
            description: 'Whether set_cell_contents persists edits back to the source JSON file.',
          },
          sheets: {
            type: 'array',
            items: {
              type: 'object',
              required: ['id', 'name', 'dimensions'],
              properties: {
                id: {
                  type: 'number',
                },
                name: {
                  type: 'string',
                },
                dimensions: {
                  type: 'object',
                  description: 'Current used rows and columns for the sheet.',
                },
              },
            },
          },
        },
        additionalProperties: false,
      },
      annotations: readOnlyAnnotation('List WorkPaper Sheets'),
    },
    {
      name: 'read_range',
      title: 'Read WorkPaper Range',
      description:
        'Read calculated values plus serialized formulas/inputs for an A1 range. Use for audit readback after edits; use read_cell for one address.',
      inputSchema: {
        type: 'object',
        required: ['range'],
        properties: {
          range: {
            type: 'string',
            description: 'A1 range such as Summary!A1:B5. If omitted from the range, pass sheetName separately.',
          },
          sheetName: {
            type: 'string',
            description: 'Default sheet name when range omits a sheet name, for example Summary.',
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
            description: 'Canonical A1 range including the sheet name.',
          },
          values: {
            type: 'array',
            description: 'Two-dimensional array of evaluated cell values.',
          },
          serialized: {
            type: 'array',
            description: 'Two-dimensional array of raw serialized cell contents, including formulas.',
          },
        },
        additionalProperties: false,
      },
      annotations: readOnlyAnnotation('Read WorkPaper Range'),
    },
    {
      name: 'read_cell',
      title: 'Read WorkPaper Cell',
      description:
        'Read one cell with calculated value, display text, formula text, and serialized content. Use after set_cell_contents to verify readback.',
      inputSchema: cellAddressSchema(['sheetName', 'address']),
      outputSchema: cellReadOutputSchema(),
      annotations: readOnlyAnnotation('Read WorkPaper Cell'),
    },
    {
      name: 'set_cell_contents',
      title: 'Set WorkPaper Cell Contents',
      description: writable
        ? 'Write raw content to one cell, recalculate dependents, atomically persist the WorkPaper JSON file, and return before/after/restored readback.'
        : 'Write raw content to one cell and recalculate dependents in memory only. Start with --writable when the edit should persist to JSON.',
      inputSchema: {
        type: 'object',
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: {
            type: 'string',
            description: 'Existing sheet name, for example Inputs.',
          },
          address: {
            type: 'string',
            description: 'Single A1 cell address such as B3. Ranges are not accepted.',
          },
          value: {
            type: ['string', 'number', 'boolean', 'null'],
            description: 'Raw cell content. Formula strings must start with =; plain strings are stored as literals.',
          },
        },
        additionalProperties: false,
      },
      outputSchema: {
        type: 'object',
        required: ['editedCell', 'before', 'after', 'restored', 'persistence', 'checks'],
        properties: {
          editedCell: {
            type: 'string',
            description: 'Canonical sheet-qualified address that was edited.',
          },
          before: cellReadOutputSchema(),
          after: cellReadOutputSchema(),
          restored: cellReadOutputSchema(),
          persistence: {
            type: 'object',
            required: ['persisted', 'serializedBytes'],
            properties: {
              persisted: {
                type: 'boolean',
              },
              path: {
                type: 'string',
              },
              serializedBytes: {
                type: 'number',
              },
            },
          },
          checks: {
            type: 'object',
            required: ['persisted', 'restoredMatchesAfter', 'previousSerialized', 'newSerialized'],
            properties: {
              persisted: {
                type: 'boolean',
              },
              restoredMatchesAfter: {
                type: 'boolean',
                description: 'True when exported and re-imported JSON preserves the edited cell readback.',
              },
              previousSerialized: rawCellContentSchema(),
              newSerialized: rawCellContentSchema(),
            },
          },
        },
        additionalProperties: false,
      },
      annotations: {
        title: 'Set WorkPaper Cell Contents',
        readOnlyHint: false,
        destructiveHint: writable,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    {
      name: 'get_cell_display_value',
      title: 'Get WorkPaper Cell Display Value',
      description:
        'Return the formatted display string for one cell. Use when an agent needs what a user would see, not the raw numeric value.',
      inputSchema: cellAddressSchema(['sheetName', 'address']),
      outputSchema: {
        type: 'object',
        required: ['address', 'displayValue'],
        properties: {
          address: {
            type: 'string',
          },
          displayValue: {
            type: 'string',
          },
        },
        additionalProperties: false,
      },
      annotations: readOnlyAnnotation('Get WorkPaper Cell Display Value'),
    },
    {
      name: 'export_workpaper_document',
      title: 'Export WorkPaper Document',
      description:
        'Export the current WorkPaper JSON document for persistence, review, or handoff to another agent. Does not write files by itself.',
      inputSchema: {
        type: 'object',
        properties: {
          includeConfig: {
            type: 'boolean',
            default: true,
            description: 'Include workbook configuration metadata in the exported JSON. Defaults to true.',
          },
        },
        additionalProperties: false,
      },
      outputSchema: {
        type: 'object',
        required: ['document', 'serializedBytes'],
        properties: {
          sourcePath: {
            type: 'string',
          },
          document: {
            type: 'object',
            description: 'Persisted WorkPaper JSON document.',
          },
          serializedBytes: {
            type: 'number',
          },
        },
        additionalProperties: false,
      },
      annotations: readOnlyAnnotation('Export WorkPaper Document'),
    },
    {
      name: 'validate_formula',
      title: 'Validate WorkPaper Formula',
      description:
        'Validate formula syntax with the WorkPaper parser before writing it to a cell. This checks syntax only; use set_cell_contents plus readback to evaluate.',
      inputSchema: {
        type: 'object',
        required: ['formula'],
        properties: {
          formula: {
            type: 'string',
            description: 'Formula string including the leading =, for example =SUM(Inputs!B2:B4).',
          },
        },
        additionalProperties: false,
      },
      outputSchema: {
        type: 'object',
        required: ['formula', 'valid'],
        properties: {
          formula: {
            type: 'string',
          },
          valid: {
            type: 'boolean',
          },
        },
        additionalProperties: false,
      },
      annotations: readOnlyAnnotation('Validate WorkPaper Formula'),
    },
  ]
}

function callFileBackedTool(input: {
  workbook: WorkPaper
  writable: boolean
  sourcePath: string | undefined
  persist: (workbook: WorkPaper) => FileBackedWorkPaperPersistResult
  params: JsonObject | undefined
}): JsonObject {
  const parsedParams = requireRecord(input.params ?? {}, 'MCP tool call params')
  const toolName = parsedParams['name']
  const args = requireRecord(parsedParams['arguments'] ?? {}, `${String(toolName)} arguments`)

  if (toolName === 'list_sheets') {
    return {
      sourcePath: input.sourcePath,
      writable: input.writable,
      sheets: input.workbook.getSheetNames().map((name) => {
        const sheetId = requireSheet(input.workbook, name)
        return {
          id: sheetId,
          name,
          dimensions: input.workbook.getSheetDimensions(sheetId),
        }
      }),
    }
  }

  if (toolName === 'read_range') {
    const range = requireString(args['range'], 'range')
    const defaultSheet = optionalSheetId(input.workbook, args['sheetName'])
    const parsedRange = input.workbook.simpleCellRangeFromString(range, defaultSheet)
    if (parsedRange === undefined) {
      throw new Error(`Invalid range: ${range}`)
    }
    return {
      range: input.workbook.simpleCellRangeToString(parsedRange, { includeSheetName: true }),
      values: input.workbook.getRangeValues(parsedRange),
      serialized: input.workbook.getRangeSerialized(parsedRange),
    }
  }

  if (toolName === 'read_cell') {
    return readCell(input.workbook, parseCellArgs(input.workbook, args))
  }

  if (toolName === 'set_cell_contents') {
    const address = parseCellArgs(input.workbook, args)
    const value = parseRawCellContent(args['value'])
    const before = readCell(input.workbook, address)

    input.workbook.setCellContents(address, value)

    const after = readCell(input.workbook, address)
    const persistence = input.persist(input.workbook)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkbook(input.workbook)))
    const restoredAddress = requireCellAddress(
      restored,
      requireString(args['sheetName'], 'sheetName'),
      requireString(args['address'], 'address'),
    )
    const restoredCell = readCell(restored, restoredAddress)

    return {
      editedCell: input.workbook.simpleCellAddressToString(address, { includeSheetName: true }),
      before,
      after,
      restored: restoredCell,
      persistence,
      checks: {
        persisted: persistence.persisted,
        restoredMatchesAfter: JSON.stringify(after) === JSON.stringify(restoredCell),
        previousSerialized: before['serialized'],
        newSerialized: after['serialized'],
      },
    }
  }

  if (toolName === 'get_cell_display_value') {
    const cell = readCell(input.workbook, parseCellArgs(input.workbook, args))
    return {
      address: cell['address'],
      displayValue: cell['displayValue'],
    }
  }

  if (toolName === 'export_workpaper_document') {
    const includeConfig = args['includeConfig'] === undefined ? true : requireBoolean(args['includeConfig'], 'includeConfig')
    const document = exportWorkPaperDocument(input.workbook, { includeConfig })
    const serialized = serializeWorkPaperDocument(document)
    return {
      sourcePath: input.sourcePath,
      document,
      serializedBytes: Buffer.byteLength(serialized, 'utf8'),
    }
  }

  if (toolName === 'validate_formula') {
    const formula = requireString(args['formula'], 'formula')
    return {
      formula,
      valid: input.workbook.validateFormula(formula),
    }
  }

  throw new Error(`Unknown WorkPaper tool: ${String(toolName)}`)
}

function readCell(workbook: WorkPaper, address: WorkPaperCellAddress): JsonObject {
  const value = workbook.getCellValue(address)
  const format = workbook.getCellValueFormat(address)
  return {
    address: workbook.simpleCellAddressToString(address, { includeSheetName: true }),
    value,
    serialized: workbook.getCellSerialized(address),
    formula: workbook.getCellFormula(address) ?? null,
    displayValue: formatCellDisplayValue(value, format),
  }
}

function toolResult(structuredContent: JsonObject): FileBackedToolCallResult {
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify(structuredContent),
      },
    ],
    structuredContent,
    isError: false,
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

  return {
    jsonrpc: '2.0',
    id,
    method: request['method'],
    ...(request['params'] !== undefined ? { params: requireRecord(request['params'], 'JSON-RPC params') } : {}),
  }
}

function parseCellArgs(workbook: WorkPaper, args: JsonObject): WorkPaperCellAddress {
  return requireCellAddress(workbook, requireString(args['sheetName'], 'sheetName'), requireString(args['address'], 'address'))
}

function requireCellAddress(workbook: WorkPaper, sheetName: string, a1Address: string): WorkPaperCellAddress {
  const sheetId = requireSheet(workbook, sheetName)
  const parsed = workbook.simpleCellAddressFromString(a1Address, sheetId)
  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid cell address: ${sheetName}!${a1Address}`)
  }
  return parsed
}

function optionalSheetId(workbook: WorkPaper, value: unknown): number | undefined {
  if (value === undefined) {
    return undefined
  }
  return requireSheet(workbook, requireString(value, 'sheetName'))
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function parseRawCellContent(value: unknown): RawCellContent {
  if (value !== null && typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') {
    throw new Error(`Unsupported cell value: ${JSON.stringify(value)}`)
  }
  return value
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

function requireBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`Expected ${label} to be a boolean`)
  }
  return value
}

function isRecord(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function serializeWorkbook(workbook: WorkPaper): string {
  return serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
}

function createMemoryPersist(workbook: WorkPaper): () => FileBackedWorkPaperPersistResult {
  return () => {
    const serialized = serializeWorkbook(workbook)
    return {
      persisted: false,
      serializedBytes: Buffer.byteLength(serialized, 'utf8'),
    }
  }
}

function writeFileAtomically(path: string, contents: string): void {
  const tempPath = resolve(dirname(path), `.${basename(path)}.${process.pid.toString()}.tmp`)
  writeFileSync(tempPath, contents)
  renameSync(tempPath, path)
}

function emptySchema(): JsonObject {
  return {
    type: 'object',
    additionalProperties: false,
  }
}

function cellAddressSchema(required: string[]): JsonObject {
  return {
    type: 'object',
    required,
    properties: {
      sheetName: {
        type: 'string',
        description: 'Existing sheet name.',
      },
      address: {
        type: 'string',
        description: 'Single A1 cell address such as B3.',
      },
    },
    additionalProperties: false,
  }
}

function cellReadOutputSchema(): JsonObject {
  return {
    type: 'object',
    required: ['address', 'value', 'serialized', 'formula', 'displayValue'],
    properties: {
      address: {
        type: 'string',
        description: 'Canonical sheet-qualified A1 address.',
      },
      value: {
        description: 'Calculated cell value.',
      },
      serialized: rawCellContentSchema(),
      formula: {
        type: ['string', 'null'],
        description: 'Formula text without losing the original calculated value context, or null for literal cells.',
      },
      displayValue: {
        type: 'string',
        description: 'Formatted value as a user would see it.',
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

function readOnlyAnnotation(title: string): WorkPaperMcpToolAnnotations {
  return {
    title,
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  }
}

export {
  createFileBackedWorkPaperMcpToolServer,
  createFileBackedWorkPaperMcpToolServerFromFile,
  type FileBackedWorkPaperMcpOptions,
  type FileBackedWorkPaperPersistResult,
  type PersistedWorkPaperDocument,
}
