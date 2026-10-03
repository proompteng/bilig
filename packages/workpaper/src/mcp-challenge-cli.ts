import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { clearExternalFunctionAdapters, installExternalFunctionAdapter } from '@bilig/formula'
import { ValueTag } from '@bilig/protocol'
import { exportWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import { createFileBackedWorkPaperMcpToolServerFromFile } from './work-paper-mcp-file-server.js'
import {
  WORKPAPER_MCP_PROTOCOL_VERSION,
  dispatchWorkPaperMcpJsonRpc,
  type WorkPaperMcpJsonRpcDispatchResult,
} from './work-paper-mcp-json-rpc.js'
import { WorkPaper } from './work-paper-runtime.js'
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
  readonly afterRestore: number
  readonly afterRestart: number
  readonly persistedDocumentBytes: number
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
}

export interface McpRevenuePlanReadback {
  readonly totalRevenue: number
  readonly westCustomers: number
  readonly enterpriseArpa: number
  readonly targetRevenue: number
  readonly qualifiedCustomerCounts: readonly number[]
}

export interface McpRevenuePlanChallengeProof {
  readonly transport: 'stdio-json-rpc'
  readonly scenario: 'revenue-plan'
  readonly protocolVersion: string
  readonly serverName: string
  readonly workpaperPath?: string
  readonly tools: readonly string[]
  readonly resources: readonly string[]
  readonly prompts: readonly string[]
  readonly formulaFamilies: readonly string[]
  readonly editedCell: 'Deals!C2'
  readonly readbackRange: 'Summary!B2:B8'
  readonly before: McpRevenuePlanReadback
  readonly after: McpRevenuePlanReadback
  readonly afterRestart: McpRevenuePlanReadback
  readonly persistedDocumentBytes: number
  readonly persistence: {
    readonly persisted: boolean
    readonly serializedBytes: number
  }
  readonly checks: {
    readonly listedFileBackedTools: boolean
    readonly listedResourcesAndPrompts: boolean
    readonly formulaValidationPassed: boolean
    readonly totalRevenueRecalculated: boolean
    readonly sumifReadbackChanged: boolean
    readonly xlookupReadbackStable: boolean
    readonly filterSpillUpdated: boolean
    readonly namedExpressionApplied: boolean
    readonly persistedToDisk: boolean
    readonly restoredReadbackMatchesAfter: boolean
    readonly exportContainsWorkPaperDocument: boolean
    readonly restartReadbackMatchesAfter: boolean
  }
  readonly verified: boolean
  readonly limitations: readonly string[]
}

export interface McpProviderBackedChallengeProof {
  readonly transport: 'stdio-json-rpc'
  readonly scenario: 'provider-backed'
  readonly providerFunction: 'IMPORTRANGE'
  readonly adapterSurface: 'web'
  readonly protocolVersion: string
  readonly serverName: string
  readonly workpaperPath?: string
  readonly tools: readonly string[]
  readonly resources: readonly string[]
  readonly prompts: readonly string[]
  readonly target: 'Imports!B2'
  readonly formula: '=IMPORTRANGE("source","Revenue!B2")'
  readonly adapterFormula: '=IMPORTRANGE("source","Revenue!B2")+0'
  readonly before: JsonObject
  readonly after: JsonObject
  readonly afterRestart: JsonObject
  readonly diagnostics: readonly JsonObject[]
  readonly persistedDocumentBytes: number
  readonly checks: {
    readonly listedFileBackedTools: boolean
    readonly listedResourcesAndPrompts: boolean
    readonly formulaValidationPassed: boolean
    readonly blockedReadbackIsBlocked: boolean
    readonly blockedDiagnosticExplainsAdapter: boolean
    readonly adapterBackedReadbackIsFresh: boolean
    readonly adapterBackedDiagnosticsCleared: boolean
    readonly exportContainsWorkPaperDocument: boolean
    readonly restartReadbackMatchesAfter: boolean
  }
  readonly verified: boolean
  readonly limitations: readonly string[]
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
  'set_cell_contents_and_readback',
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

const revenuePlanFormulaFamilies = ['SUM', 'SUMIF', 'XLOOKUP', 'FILTER', 'named-expression', 'persistence', 'restart'] as const
const providerBackedBlockedFormula = '=IMPORTRANGE("source","Revenue!B2")' as const
const providerBackedAdapterFormula = '=IMPORTRANGE("source","Revenue!B2")+0' as const

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
    '  --markdown   Print a Markdown report.',
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
      afterRestore: afterRestart,
      afterRestart,
      persistedDocumentBytes: serializedBytes,
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
        'This challenge proves the file-backed MCP WorkPaper tool surface, not every desktop spreadsheet feature.',
        'Use saved-file compatibility evaluators only when a workbook file is the integration contract.',
      ],
    }

    return keepTemp ? { ...proof, workpaperPath } : proof
  } finally {
    if (!keepTemp) {
      rmSync(tempDir, { recursive: true, force: true })
    }
  }
}

export function buildMcpRevenuePlanChallengeProof(options: McpChallengeBuildOptions = {}): McpRevenuePlanChallengeProof {
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-mcp-revenue-plan-'))
  const workpaperPath = join(tempDir, 'revenue-plan.workpaper.json')
  const keepTemp = options.keepTemp ?? false

  try {
    writeFileSync(workpaperPath, serializeWorkPaperDocument(exportWorkPaperDocument(createRevenuePlanWorkPaper(), { includeConfig: true })))

    const server = createFileBackedWorkPaperMcpToolServerFromFile({
      workpaperPath,
      writable: true,
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
    const formulaValidation = toolStructuredContent(
      callTool(server, 5, 'validate_formula', {
        formula: '=SUM(Deals!E2:E4)*(100+GrowthRatePercent)/100',
      }),
      'validate_formula',
    )
    const write = toolStructuredContent(
      callTool(server, 6, 'set_cell_contents_and_readback', {
        sheetName: 'Deals',
        address: 'C2',
        value: 20,
        readbackSheetName: 'Summary',
        readbackRange: 'B2:B8',
      }),
      'set_cell_contents_and_readback',
    )
    const exported = toolStructuredContent(
      callTool(server, 7, 'export_workpaper_document', {
        includeConfig: true,
      }),
      'export_workpaper_document',
    )
    const restartedServer = createFileBackedWorkPaperMcpToolServerFromFile({
      workpaperPath,
      writable: false,
    })
    const restartedReadback = toolStructuredContent(
      callTool(restartedServer, 8, 'read_range', {
        sheetName: 'Summary',
        range: 'B2:B8',
      }),
      'read_range after restart',
    )
    const serverInfo = requireRecord(initialized['serverInfo'], 'initialize serverInfo')
    const before = revenuePlanReadback(requireRecord(write['beforeReadback'], 'beforeReadback'))
    const after = revenuePlanReadback(requireRecord(write['afterReadback'], 'afterReadback'))
    const afterRestart = revenuePlanReadback(restartedReadback)
    const persistence = requireRecord(write['persistence'], 'set_cell_contents_and_readback persistence')
    const serializedBytes = requireNumber(persistence['serializedBytes'], 'persistence.serializedBytes')
    const writeChecks = requireRecord(write['checks'], 'set_cell_contents_and_readback checks')
    const checks = {
      listedFileBackedTools: arraysEqual(tools, expectedFileBackedTools),
      listedResourcesAndPrompts: arraysEqual(resources, expectedResources) && arraysEqual(prompts, expectedPrompts),
      formulaValidationPassed: formulaValidation['valid'] === true,
      totalRevenueRecalculated: before.totalRevenue === 27_300 && after.totalRevenue === 36_900,
      sumifReadbackChanged: before.westCustomers === 30 && after.westCustomers === 38,
      xlookupReadbackStable: before.enterpriseArpa === 1_200 && after.enterpriseArpa === 1_200,
      filterSpillUpdated:
        numberArraysEqual(before.qualifiedCustomerCounts, [30, 18]) && numberArraysEqual(after.qualifiedCustomerCounts, [20, 30, 18]),
      namedExpressionApplied: before.targetRevenue === 30_576 && after.targetRevenue === 41_328,
      persistedToDisk:
        write['editedCell'] === 'Deals!C2' && writeChecks['persisted'] === true && persistence['persisted'] === true && serializedBytes > 0,
      restoredReadbackMatchesAfter: writeChecks['restoredReadbackMatchesAfter'] === true,
      exportContainsWorkPaperDocument:
        isRecord(exported['document']) && requireNumber(exported['serializedBytes'], 'exported.serializedBytes') > 0,
      restartReadbackMatchesAfter: JSON.stringify(afterRestart) === JSON.stringify(after),
    }
    const proof: McpRevenuePlanChallengeProof = {
      transport: 'stdio-json-rpc',
      scenario: 'revenue-plan',
      protocolVersion: requireString(initialized['protocolVersion'], 'protocolVersion'),
      serverName: requireString(serverInfo['name'], 'serverInfo.name'),
      tools,
      resources,
      prompts,
      formulaFamilies: revenuePlanFormulaFamilies,
      editedCell: 'Deals!C2',
      readbackRange: 'Summary!B2:B8',
      before,
      after,
      afterRestart,
      persistedDocumentBytes: serializedBytes,
      persistence: {
        persisted: persistence['persisted'] === true,
        serializedBytes,
      },
      checks,
      verified:
        checks.listedFileBackedTools &&
        checks.listedResourcesAndPrompts &&
        checks.formulaValidationPassed &&
        checks.totalRevenueRecalculated &&
        checks.sumifReadbackChanged &&
        checks.xlookupReadbackStable &&
        checks.filterSpillUpdated &&
        checks.namedExpressionApplied &&
        checks.persistedToDisk &&
        checks.restoredReadbackMatchesAfter &&
        checks.exportContainsWorkPaperDocument &&
        checks.restartReadbackMatchesAfter,
      limitations: [
        'This revenue-plan challenge proves a realistic file-backed MCP WorkPaper path, not every desktop spreadsheet feature.',
        'It covers formula readback for SUM, SUMIF, XLOOKUP, FILTER spills, and a named expression; use real workbook fixtures for domain-specific parity.',
      ],
    }

    return keepTemp ? { ...proof, workpaperPath } : proof
  } finally {
    if (!keepTemp) {
      rmSync(tempDir, { recursive: true, force: true })
    }
  }
}

export function buildMcpProviderBackedChallengeProof(options: McpChallengeBuildOptions = {}): McpProviderBackedChallengeProof {
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-mcp-provider-backed-'))
  const workpaperPath = join(tempDir, 'provider-backed.workpaper.json')
  const keepTemp = options.keepTemp ?? false

  clearExternalFunctionAdapters()
  try {
    writeFileSync(
      workpaperPath,
      serializeWorkPaperDocument(exportWorkPaperDocument(createProviderBackedWorkPaper(), { includeConfig: true })),
    )

    const blockedServer = createFileBackedWorkPaperMcpToolServerFromFile({
      workpaperPath,
      writable: false,
    })
    const initialize = rpcResult(
      dispatchWorkPaperMcpJsonRpc(
        {
          jsonrpc: '2.0',
          id: 1,
          method: 'initialize',
        },
        { server: blockedServer, protocolVersion: WORKPAPER_MCP_PROTOCOL_VERSION },
      ),
      'initialize',
    )
    const initialized = requireRecord(initialize, 'initialize result')
    const tools = readToolNames(rpcResult(callJsonRpc(blockedServer, 2, 'tools/list'), 'tools/list result'))
    const resources = readResourceUris(rpcResult(callJsonRpc(blockedServer, 3, 'resources/list'), 'resources/list result'))
    const prompts = readPromptNames(rpcResult(callJsonRpc(blockedServer, 4, 'prompts/list'), 'prompts/list result'))
    const formulaValidation = toolStructuredContent(
      callTool(blockedServer, 5, 'validate_formula', {
        formula: providerBackedBlockedFormula,
      }),
      'validate_formula',
    )
    const before = toolStructuredContent(
      callTool(blockedServer, 6, 'read_cell', {
        sheetName: 'Imports',
        address: 'B2',
      }),
      'read_cell provider-backed blocked',
    )

    installExternalFunctionAdapter({
      surface: 'web',
      resolveFunction(name) {
        if (name !== 'IMPORTRANGE') {
          return undefined
        }
        return {
          kind: 'lookup',
          implementation: () => ({
            tag: ValueTag.Number,
            value: 96_000,
          }),
        }
      },
    })

    const adapterServer = createFileBackedWorkPaperMcpToolServerFromFile({
      workpaperPath,
      writable: true,
    })
    toolStructuredContent(
      callTool(adapterServer, 7, 'set_cell_contents', {
        sheetName: 'Imports',
        address: 'B2',
        value: 0,
      }),
      'set_cell_contents provider-backed reset',
    )
    const adapterWrite = toolStructuredContent(
      callTool(adapterServer, 8, 'set_cell_contents', {
        sheetName: 'Imports',
        address: 'B2',
        value: providerBackedAdapterFormula,
      }),
      'set_cell_contents provider-backed adapter',
    )
    const after = requireRecord(adapterWrite['after'], 'set_cell_contents provider-backed after')
    const exported = toolStructuredContent(
      callTool(adapterServer, 9, 'export_workpaper_document', {
        includeConfig: true,
      }),
      'export_workpaper_document',
    )
    const restartedServer = createFileBackedWorkPaperMcpToolServerFromFile({
      workpaperPath,
      writable: false,
    })
    const afterRestart = toolStructuredContent(
      callTool(restartedServer, 10, 'read_cell', {
        sheetName: 'Imports',
        address: 'B2',
      }),
      'read_cell provider-backed restart',
    )
    const serverInfo = requireRecord(initialized['serverInfo'], 'initialize serverInfo')
    const diagnostics = requireArray(before['formulaDiagnostics'], 'formulaDiagnostics').map((entry) =>
      requireRecord(entry, 'formula diagnostic'),
    )
    const afterDiagnostics = requireArray(after['formulaDiagnostics'], 'after formulaDiagnostics')
    const checks = {
      listedFileBackedTools: arraysEqual(tools, expectedFileBackedTools),
      listedResourcesAndPrompts: arraysEqual(resources, expectedResources) && arraysEqual(prompts, expectedPrompts),
      formulaValidationPassed: formulaValidation['valid'] === true,
      blockedReadbackIsBlocked: before['displayValue'] === '#BLOCKED!',
      blockedDiagnosticExplainsAdapter: diagnostics.some(
        (diagnostic) =>
          diagnostic['code'] === 'provider-backed-adapter-missing' &&
          diagnostic['functionName'] === 'IMPORTRANGE' &&
          diagnostic['adapterSurface'] === 'web' &&
          diagnostic['errorText'] === '#BLOCKED!',
      ),
      adapterBackedReadbackIsFresh: numericCellValue(after) === 96_000 && after['displayValue'] === '96000',
      adapterBackedDiagnosticsCleared: afterDiagnostics.length === 0,
      exportContainsWorkPaperDocument:
        isRecord(exported['document']) && requireNumber(exported['serializedBytes'], 'exported.serializedBytes') > 0,
      restartReadbackMatchesAfter: JSON.stringify(afterRestart) === JSON.stringify(after),
    }
    const proof: McpProviderBackedChallengeProof = {
      transport: 'stdio-json-rpc',
      scenario: 'provider-backed',
      providerFunction: 'IMPORTRANGE',
      adapterSurface: 'web',
      protocolVersion: requireString(initialized['protocolVersion'], 'protocolVersion'),
      serverName: requireString(serverInfo['name'], 'serverInfo.name'),
      tools,
      resources,
      prompts,
      target: 'Imports!B2',
      formula: providerBackedBlockedFormula,
      adapterFormula: providerBackedAdapterFormula,
      before,
      after,
      afterRestart,
      diagnostics,
      persistedDocumentBytes: requireNumber(exported['serializedBytes'], 'exported.serializedBytes'),
      checks,
      verified:
        checks.listedFileBackedTools &&
        checks.listedResourcesAndPrompts &&
        checks.formulaValidationPassed &&
        checks.blockedReadbackIsBlocked &&
        checks.blockedDiagnosticExplainsAdapter &&
        checks.adapterBackedReadbackIsFresh &&
        checks.adapterBackedDiagnosticsCleared &&
        checks.exportContainsWorkPaperDocument &&
        checks.restartReadbackMatchesAfter,
      limitations: [
        'This provider-backed challenge proves an explicit adapter boundary for IMPORTRANGE without calling Google Sheets.',
        'The adapter is synthetic and local; production IMPORTRANGE use needs a host adapter that owns authorization and remote range fetching.',
      ],
    }

    return keepTemp ? { ...proof, workpaperPath } : proof
  } finally {
    clearExternalFunctionAdapters()
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

If this proof matched your workflow, keep the repository and release feed nearby:
https://github.com/proompteng/bilig

If it almost worked, open the concrete MCP or workbook blocker:
https://github.com/proompteng/bilig/discussions/new?category=general
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

function createRevenuePlanWorkPaper(): WorkPaper {
  const workbook = WorkPaper.buildFromSheets({
    Deals: [
      ['Region', 'Segment', 'Customers', 'ARPA', 'Revenue'],
      ['West', 'Enterprise', 12, 1200, '=C2*D2'],
      ['East', 'SMB', 30, 250, '=C3*D3'],
      ['West', 'SMB', 18, 300, '=C4*D4'],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Total revenue', '=SUM(Deals!E2:E4)'],
      ['West customers', '=SUMIF(Deals!A2:A4,"West",Deals!C2:C4)'],
      ['Enterprise ARPA', '=XLOOKUP("Enterprise",Deals!B2:B4,Deals!D2:D4)'],
      ['Target revenue', null],
      ['Qualified customer counts', '=FILTER(Deals!C2:C4,Deals!C2:C4>=18)'],
    ],
  })
  const summarySheet = workbook.getSheetId('Summary')
  if (summarySheet === undefined) {
    throw new Error('Expected Summary sheet to exist.')
  }
  workbook.addNamedExpression('GrowthRatePercent', 12)
  workbook.setCellContents({ sheet: summarySheet, row: 4, col: 1 }, '=SUM(Deals!E2:E4)*(100+GrowthRatePercent)/100')
  return workbook
}

function createProviderBackedWorkPaper(): WorkPaper {
  return WorkPaper.buildFromSheets({
    Imports: [
      ['Metric', 'Value'],
      ['Remote ARR', providerBackedBlockedFormula],
    ],
  })
}

function revenuePlanReadback(range: JsonObject): McpRevenuePlanReadback {
  const values = requireArray(range['values'], 'readback values')
  const numbers = values.map((row, index) => optionalColumnNumber(row, index))
  return {
    totalRevenue: requireReadbackNumber(numbers[0], 'total revenue'),
    westCustomers: requireReadbackNumber(numbers[1], 'west customers'),
    enterpriseArpa: requireReadbackNumber(numbers[2], 'enterprise ARPA'),
    targetRevenue: requireReadbackNumber(numbers[3], 'target revenue'),
    qualifiedCustomerCounts: numbers.slice(4).filter((value): value is number => value !== undefined),
  }
}

function optionalColumnNumber(row: unknown, index: number): number | undefined {
  const cells = requireArray(row, `readback row ${String(index + 1)}`)
  const cell = cells[0]
  if (isRecord(cell) && typeof cell['value'] === 'number') {
    return cell['value']
  }
  if (typeof cell === 'number') {
    return cell
  }
  return undefined
}

function requireReadbackNumber(value: number | undefined, label: string): number {
  if (value === undefined) {
    throw new Error(`Expected ${label} readback to be numeric.`)
  }
  return value
}

function numberArraysEqual(actual: readonly number[], expected: readonly number[]): boolean {
  return actual.length === expected.length && actual.every((item, index) => item === expected[index])
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
