import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { createFileBackedWorkPaperMcpToolServerFromFile } from './work-paper-mcp-file-server.js'
import {
  WORKPAPER_MCP_PROTOCOL_VERSION,
  dispatchWorkPaperMcpJsonRpc,
  type WorkPaperMcpJsonRpcDispatchResult,
} from './work-paper-mcp-json-rpc.js'
import type { WorkPaperMcpToolServer } from './work-paper-mcp-server.js'

type JsonObject = Record<string, unknown>

export interface McpChallengeProof {
  readonly transport: 'stdio-json-rpc'
  readonly protocolVersion: string
  readonly serverName: string
  readonly workpaperPath?: string
  readonly tools: readonly string[]
  readonly resources: readonly string[]
  readonly prompts: readonly string[]
  readonly editedCell: 'Inputs!B3'
  readonly dependentCell: 'Summary!B3'
  readonly before: number
  readonly after: number
  readonly afterRestart: number
  readonly displayValue: string
  readonly persistence: {
    readonly persisted: boolean
    readonly serializedBytes: number
  }
  readonly checks: {
    readonly listedFileBackedTools: boolean
    readonly listedResourcesAndPrompts: boolean
    readonly formulaValidationPassed: boolean
    readonly dependentCellChanged: boolean
    readonly persistedToDisk: boolean
    readonly exportContainsWorkPaperDocument: boolean
    readonly restartReadbackMatchesAfter: boolean
    readonly displayValueRead: boolean
  }
  readonly verified: boolean
  readonly limitations: readonly string[]
  readonly nextStep: string
}

export interface McpChallengeCliHost {
  readonly argv: readonly string[]
  readonly writeStderr?: (text: string) => void
  readonly writeStdout?: (text: string) => void
}

type McpChallengeOutputMode = 'json' | 'markdown'

interface McpChallengeCliOptions {
  readonly help: boolean
  readonly keepTemp: boolean
  readonly outputMode: McpChallengeOutputMode
}

interface McpChallengeBuildOptions {
  readonly keepTemp?: boolean
}

const expectedFileBackedTools = [
  'list_sheets',
  'read_range',
  'read_cell',
  'set_cell_contents',
  'get_cell_display_value',
  'export_workpaper_document',
  'validate_formula',
] as const

const expectedResources = [
  'bilig://workpaper/manifest',
  'bilig://workpaper/agent-handoff',
  'bilig://workpaper/sheets',
  'bilig://workpaper/current-document',
] as const

const expectedPrompts = ['edit_and_verify_workpaper', 'debug_workpaper_formula'] as const

export function runMcpChallengeCli(host: McpChallengeCliHost): number {
  const writeStdout = host.writeStdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = host.writeStderr ?? ((text: string) => process.stderr.write(text))
  let options: McpChallengeCliOptions

  try {
    options = parseMcpChallengeCliArgs(host.argv)
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n\n${mcpChallengeHelpText()}`)
    return 1
  }

  if (options.help) {
    writeStdout(mcpChallengeHelpText())
    return 0
  }

  try {
    const proof = buildMcpChallengeProof({ keepTemp: options.keepTemp })
    writeStdout(renderMcpChallengeProof(proof, options.outputMode))
    return proof.verified ? 0 : 1
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    return 1
  }
}

export function parseMcpChallengeCliArgs(args: readonly string[]): McpChallengeCliOptions {
  let help = false
  let keepTemp = false
  let outputMode: McpChallengeOutputMode = 'json'

  for (const arg of args) {
    if (arg === '--help' || arg === '-h') {
      help = true
      continue
    }
    if (arg === '--json') {
      outputMode = 'json'
      continue
    }
    if (arg === '--markdown') {
      outputMode = 'markdown'
      continue
    }
    if (arg === '--keep-temp') {
      keepTemp = true
      continue
    }
    throw new Error(`Unknown bilig-mcp-challenge argument: ${arg}`)
  }

  return { help, keepTemp, outputMode }
}

export function mcpChallengeHelpText(): string {
  return [
    'Usage: bilig-mcp-challenge [--json|--markdown] [--keep-temp]',
    '',
    'Runs the Bilig file-backed MCP challenge without cloning the repository:',
    'initialize MCP JSON-RPC, list the writable WorkPaper tools/resources/prompts,',
    'edit Inputs!B3, read recalculated Summary!B3, export JSON, restart from disk,',
    'and print a proof object with verified: true.',
    '',
    'Options:',
    '  --json       Print machine-readable JSON. Default.',
    '  --markdown   Print a paste-ready Markdown report.',
    '  --keep-temp  Keep the temporary WorkPaper JSON file and include its path.',
    '  -h, --help   Print this help text.',
    '',
  ].join('\n')
}

export function buildMcpChallengeProof(options: McpChallengeBuildOptions = {}): McpChallengeProof {
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-mcp-challenge-'))
  const workpaperPath = join(tempDir, 'pricing.workpaper.json')
  const keepTemp = options.keepTemp ?? false

  try {
    const server = createFileBackedWorkPaperMcpToolServerFromFile({
      workpaperPath,
      writable: true,
      initDemoWorkPaper: true,
    })
    const initialize = rpcResult(
      dispatchWorkPaperMcpJsonRpc(
        {
          jsonrpc: '2.0',
          id: 1,
          method: 'initialize',
        },
        { server, protocolVersion: WORKPAPER_MCP_PROTOCOL_VERSION },
      ),
      'initialize',
    )
    const initialized = requireRecord(initialize, 'initialize result')
    const tools = readToolNames(rpcResult(callJsonRpc(server, 2, 'tools/list'), 'tools/list result'))
    const resources = readResourceUris(rpcResult(callJsonRpc(server, 3, 'resources/list'), 'resources/list result'))
    const prompts = readPromptNames(rpcResult(callJsonRpc(server, 4, 'prompts/list'), 'prompts/list result'))
    const beforeCell = toolStructuredContent(
      callTool(server, 5, 'read_cell', {
        sheetName: 'Summary',
        address: 'B3',
      }),
      'read_cell before',
    )
    const formulaValidation = toolStructuredContent(
      callTool(server, 6, 'validate_formula', {
        formula: '=SUM(1,2)',
      }),
      'validate_formula',
    )
    const write = toolStructuredContent(
      callTool(server, 7, 'set_cell_contents', {
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      }),
      'set_cell_contents',
    )
    const afterCell = toolStructuredContent(
      callTool(server, 8, 'read_cell', {
        sheetName: 'Summary',
        address: 'B3',
      }),
      'read_cell after',
    )
    const display = toolStructuredContent(
      callTool(server, 9, 'get_cell_display_value', {
        sheetName: 'Summary',
        address: 'B3',
      }),
      'get_cell_display_value',
    )
    const exported = toolStructuredContent(
      callTool(server, 10, 'export_workpaper_document', {
        includeConfig: true,
      }),
      'export_workpaper_document',
    )
    const restartedServer = createFileBackedWorkPaperMcpToolServerFromFile({
      workpaperPath,
      writable: false,
    })
    const restartedCell = toolStructuredContent(
      callTool(restartedServer, 11, 'read_cell', {
        sheetName: 'Summary',
        address: 'B3',
      }),
      'read_cell after restart',
    )
    const serverInfo = requireRecord(initialized['serverInfo'], 'initialize serverInfo')
    const before = numericCellValue(beforeCell)
    const after = numericCellValue(afterCell)
    const afterRestart = numericCellValue(restartedCell)
    const displayValue = requireString(display['displayValue'], 'displayValue')
    const persistence = requireRecord(write['persistence'], 'set_cell_contents persistence')
    const serializedBytes = requireNumber(persistence['serializedBytes'], 'persistence.serializedBytes')
    const checks = {
      listedFileBackedTools: arraysEqual(tools, expectedFileBackedTools),
      listedResourcesAndPrompts: arraysEqual(resources, expectedResources) && arraysEqual(prompts, expectedPrompts),
      formulaValidationPassed: formulaValidation['valid'] === true,
      dependentCellChanged: before === 60_000 && after === 96_000,
      persistedToDisk:
        write['editedCell'] === 'Inputs!B3' &&
        requireRecord(write['checks'], 'set_cell_contents checks')['persisted'] === true &&
        persistence['persisted'] === true &&
        serializedBytes > 0,
      exportContainsWorkPaperDocument:
        isRecord(exported['document']) && requireNumber(exported['serializedBytes'], 'exported.serializedBytes') > 0,
      restartReadbackMatchesAfter: afterRestart === after,
      displayValueRead: displayValue === '96000',
    }

    const proof: McpChallengeProof = {
      transport: 'stdio-json-rpc',
      protocolVersion: requireString(initialized['protocolVersion'], 'protocolVersion'),
      serverName: requireString(serverInfo['name'], 'serverInfo.name'),
      tools,
      resources,
      prompts,
      editedCell: 'Inputs!B3',
      dependentCell: 'Summary!B3',
      before,
      after,
      afterRestart,
      displayValue,
      persistence: {
        persisted: persistence['persisted'] === true,
        serializedBytes,
      },
      checks,
      verified:
        checks.listedFileBackedTools &&
        checks.listedResourcesAndPrompts &&
        checks.formulaValidationPassed &&
        checks.dependentCellChanged &&
        checks.persistedToDisk &&
        checks.exportContainsWorkPaperDocument &&
        checks.restartReadbackMatchesAfter &&
        checks.displayValueRead,
      limitations: [
        'This challenge proves the file-backed MCP WorkPaper tool surface, not Excel desktop UI automation.',
        'For XLSX-specific behavior, run bilig-formula-clinic or the XLSX recalculation example with a real workbook fixture.',
      ],
      nextStep: 'If this proof matches your agent workflow, star or watch Bilig: https://github.com/proompteng/bilig/stargazers',
    }

    return keepTemp ? { ...proof, workpaperPath } : proof
  } finally {
    if (!keepTemp) {
      rmSync(tempDir, { recursive: true, force: true })
    }
  }
}

export function renderMcpChallengeProof(proof: McpChallengeProof, outputMode: McpChallengeOutputMode): string {
  if (outputMode === 'markdown') {
    return renderMcpChallengeMarkdown(proof)
  }
  return `${JSON.stringify(proof, null, 2)}\n`
}

function renderMcpChallengeMarkdown(proof: McpChallengeProof): string {
  return `# Bilig MCP challenge

\`\`\`json
${JSON.stringify(proof, null, 2)}
\`\`\`

Result: ${proof.verified ? 'verified' : 'failed'}.

The important invariant is that \`${proof.editedCell}\` changed the dependent formula cell \`${proof.dependentCell}\`, the edit persisted to WorkPaper JSON, and a restarted file-backed MCP server read the same computed value.
`
}

function callJsonRpc(server: WorkPaperMcpToolServer, id: number, method: string, params?: JsonObject): WorkPaperMcpJsonRpcDispatchResult {
  return dispatchWorkPaperMcpJsonRpc(
    {
      jsonrpc: '2.0',
      id,
      method,
      params,
    },
    { server, protocolVersion: WORKPAPER_MCP_PROTOCOL_VERSION },
  )
}

function callTool(server: WorkPaperMcpToolServer, id: number, name: string, args: JsonObject): WorkPaperMcpJsonRpcDispatchResult {
  return callJsonRpc(server, id, 'tools/call', {
    name,
    arguments: args,
  })
}

function rpcResult(result: WorkPaperMcpJsonRpcDispatchResult, label: string): unknown {
  if (result.kind !== 'response') {
    throw new Error(`Expected ${label} to return a JSON-RPC response`)
  }
  const response = requireRecord(result.response, label)
  if (isRecord(response['error'])) {
    throw new Error(`${label} failed: ${JSON.stringify(response['error'])}`)
  }
  return response['result']
}

function toolStructuredContent(result: WorkPaperMcpJsonRpcDispatchResult, label: string): JsonObject {
  const responseResult = requireRecord(rpcResult(result, label), label)
  return requireRecord(responseResult['structuredContent'], `${label} structuredContent`)
}

function readToolNames(value: unknown): string[] {
  const result = requireRecord(value, 'tools/list result')
  const tools = requireArray(result['tools'], 'tools')
  return tools.map((tool) => requireString(requireRecord(tool, 'tool')['name'], 'tool.name'))
}

function readResourceUris(value: unknown): string[] {
  const result = requireRecord(value, 'resources/list result')
  const resources = requireArray(result['resources'], 'resources')
  return resources.map((resource) => requireString(requireRecord(resource, 'resource')['uri'], 'resource.uri'))
}

function readPromptNames(value: unknown): string[] {
  const result = requireRecord(value, 'prompts/list result')
  const prompts = requireArray(result['prompts'], 'prompts')
  return prompts.map((prompt) => requireString(requireRecord(prompt, 'prompt')['name'], 'prompt.name'))
}

function numericCellValue(cell: JsonObject): number {
  const value = cell['value']
  if (isRecord(value) && typeof value['value'] === 'number') {
    return value['value']
  }
  if (typeof value === 'number') {
    return value
  }
  throw new Error(`Expected numeric cell value, got ${JSON.stringify(cell)}`)
}

function arraysEqual(actual: readonly string[], expected: readonly string[]): boolean {
  return actual.length === expected.length && actual.every((item, index) => item === expected[index])
}

function requireRecord(value: unknown, label: string): JsonObject {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object, got ${JSON.stringify(value)}`)
  }
  return value
}

function requireArray(value: unknown, label: string): unknown[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected ${label} to be an array, got ${JSON.stringify(value)}`)
  }
  return value
}

function requireString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${label} to be a string, got ${JSON.stringify(value)}`)
  }
  return value
}

function requireNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`Expected ${label} to be a number, got ${JSON.stringify(value)}`)
  }
  return value
}

function isRecord(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
