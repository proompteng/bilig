import { WorkPaper } from './work-paper.js'
import {
  buildAgentWorkbookChallengeProof,
  buildMcpChallengeProof,
  buildMcpProviderBackedChallengeProof,
  buildMcpRevenuePlanChallengeProof,
} from './cli.js'

export const biligEvaluatorSchemaVersion = 'bilig-evaluator.v1'

export type BiligEvaluatorDoor = 'workpaper-service' | 'agent-mcp'
export type BiligEvaluatorScenario = 'default' | 'revenue-plan' | 'provider-backed'

export interface BiligEvaluatorCliContext {
  readonly stdout?: (text: string) => void
  readonly stderr?: (text: string) => void
}

export interface BiligEvaluatorDoorSummary {
  readonly door: BiligEvaluatorDoor
  readonly label: string
  readonly command: string
  readonly docs: string
}

export interface BiligEvaluatorEvidence {
  readonly scenario?: BiligEvaluatorScenario | undefined
  readonly target?: string | undefined
  readonly editedCell?: string | undefined
  readonly dependentCell?: string | undefined
  readonly dependentCells?: readonly string[] | undefined
  readonly readbackRange?: string | undefined
  readonly formulaFamilies?: readonly string[] | undefined
  readonly providerFunction?: string | undefined
  readonly adapterSurface?: string | undefined
  readonly formula?: string | undefined
  readonly adapterFormula?: string | undefined
  readonly diagnostics?: readonly unknown[] | undefined
  readonly before?: unknown
  readonly after?: unknown
  readonly afterRestore?: unknown
  readonly afterRestart?: unknown
  readonly persistedDocumentBytes?: number | undefined
  readonly toolCount?: number | undefined
  readonly tools?: readonly string[] | undefined
  readonly checks: Readonly<Record<string, boolean>>
}

export interface BiligEvaluatorProof {
  readonly schemaVersion: typeof biligEvaluatorSchemaVersion
  readonly door: BiligEvaluatorDoor
  readonly doorName: string
  readonly command: string
  readonly packageVersions: Readonly<Record<string, string>>
  readonly scenario?: BiligEvaluatorScenario | undefined
  readonly evidence: BiligEvaluatorEvidence
  readonly verified: boolean
  readonly durationMs: number
  readonly limitations: readonly string[]
  readonly next: {
    readonly docs: string
    readonly command: string
  }
  readonly sourceProof: unknown
}

interface ParsedEvaluatorArgs {
  readonly door: BiligEvaluatorDoor | undefined
  readonly scenario: BiligEvaluatorScenario
  readonly json: boolean
  readonly markdown: boolean
  readonly list: boolean
  readonly help: boolean
}

const workpaperDocs = 'https://proompteng.github.io/bilig/eval-workpaper-service.html'
const mcpDocs = 'https://proompteng.github.io/bilig/eval-agent-mcp.html'

const doorSummaries: readonly BiligEvaluatorDoorSummary[] = [
  {
    door: 'workpaper-service',
    label: 'Edit, recalculate, restore, and persist a WorkPaper service document.',
    command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json',
    docs: workpaperDocs,
  },
  {
    door: 'agent-mcp',
    label: 'Exercise the WorkPaper MCP tools through a local file-backed server restart.',
    command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    docs: mcpDocs,
  },
]

export function listBiligEvaluatorDoors(): readonly BiligEvaluatorDoorSummary[] {
  return doorSummaries
}

export function buildBiligEvaluatorProof(
  door: BiligEvaluatorDoor,
  options: { readonly scenario?: BiligEvaluatorScenario } = {},
): Promise<BiligEvaluatorProof> {
  const startedAt = Date.now()
  return buildDoorProof(door, options.scenario ?? 'default').then((proof) => ({
    ...proof,
    durationMs: Math.max(0, Date.now() - startedAt),
  }))
}

export async function runBiligEvaluatorCli(args: readonly string[], context: BiligEvaluatorCliContext = {}): Promise<number> {
  const writeStdout = context.stdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = context.stderr ?? ((text: string) => process.stderr.write(text))

  try {
    const parsed = parseEvaluatorArgs(args)
    if (parsed.help) {
      writeStdout(renderEvaluatorHelp())
      return 0
    }
    if (parsed.list) {
      const summary = {
        schemaVersion: biligEvaluatorSchemaVersion,
        doors: listBiligEvaluatorDoors(),
      }
      writeStdout(parsed.markdown ? renderDoorListMarkdown(summary.doors) : `${JSON.stringify(summary, null, 2)}\n`)
      return 0
    }
    if (parsed.door === undefined) {
      throw new Error('Expected --door workpaper-service or --door agent-mcp.')
    }

    const proof = await buildBiligEvaluatorProof(parsed.door, { scenario: parsed.scenario })
    writeStdout(parsed.markdown ? renderProofMarkdown(proof) : `${JSON.stringify(proof, null, 2)}\n`)
    return proof.verified ? 0 : 1
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    return 1
  }
}

async function buildDoorProof(
  door: BiligEvaluatorDoor,
  scenario: BiligEvaluatorScenario,
): Promise<Omit<BiligEvaluatorProof, 'durationMs'>> {
  if (scenario !== 'default' && door !== 'agent-mcp') {
    throw new Error(`Scenario "${scenario}" is only available for --door agent-mcp.`)
  }
  switch (door) {
    case 'workpaper-service':
      return buildWorkpaperServiceEvaluatorProof()
    case 'agent-mcp':
      return buildAgentMcpEvaluatorProof(scenario)
  }
}

function buildWorkpaperServiceEvaluatorProof(): Omit<BiligEvaluatorProof, 'durationMs'> {
  const sourceProof = buildAgentWorkbookChallengeProof()
  const checks = normalizeBooleanRecord(recordValue(sourceProof, 'checks'))
  const summary = requireDoorSummary('workpaper-service')

  return {
    schemaVersion: biligEvaluatorSchemaVersion,
    door: 'workpaper-service',
    doorName: 'WorkPaper service proof',
    command: summary.command,
    packageVersions: {
      '@bilig/workpaper': WorkPaper.version,
    },
    evidence: {
      editedCell: stringValue(sourceProof, 'editedCell'),
      dependentCell: stringValue(sourceProof, 'dependentCell'),
      before: propertyValue(sourceProof, 'before'),
      after: propertyValue(sourceProof, 'after'),
      afterRestore: propertyValue(sourceProof, 'afterRestore'),
      persistedDocumentBytes: numberValue(sourceProof, 'persistedDocumentBytes'),
      checks,
    },
    verified: booleanValue(sourceProof, 'verified'),
    limitations: stringArrayValue(sourceProof, 'limitations'),
    next: {
      docs: summary.docs,
      command: summary.command,
    },
    sourceProof,
  }
}

function buildAgentMcpEvaluatorProof(scenario: BiligEvaluatorScenario): Omit<BiligEvaluatorProof, 'durationMs'> {
  const sourceProof =
    scenario === 'revenue-plan'
      ? buildMcpRevenuePlanChallengeProof()
      : scenario === 'provider-backed'
        ? buildMcpProviderBackedChallengeProof()
        : buildMcpChallengeProof()
  const checks = normalizeBooleanRecord(recordValue(sourceProof, 'checks'))
  const tools = stringArrayValue(sourceProof, 'tools')
  const summary = requireDoorSummary('agent-mcp')
  const command = commandForScenario(summary, scenario)
  const formulaFamilies = stringArrayValue(sourceProof, 'formulaFamilies')
  const doorName =
    scenario === 'revenue-plan'
      ? 'Agent MCP revenue-plan proof'
      : scenario === 'provider-backed'
        ? 'Agent MCP provider-backed boundary proof'
        : 'Agent MCP proof'

  return {
    schemaVersion: biligEvaluatorSchemaVersion,
    door: 'agent-mcp',
    doorName,
    command,
    packageVersions: {
      '@bilig/workpaper': WorkPaper.version,
    },
    ...(scenario === 'default' ? {} : { scenario }),
    evidence: {
      ...(scenario === 'default' ? {} : { scenario }),
      editedCell: stringValue(sourceProof, 'editedCell'),
      dependentCell: stringValue(sourceProof, 'dependentCell'),
      ...(scenario === 'revenue-plan'
        ? {
            dependentCells: ['Summary!B2', 'Summary!B3', 'Summary!B4', 'Summary!B5', 'Summary!B6:B8'],
            readbackRange: stringValue(sourceProof, 'readbackRange'),
            formulaFamilies,
          }
        : {}),
      ...(scenario === 'provider-backed'
        ? {
            providerFunction: stringValue(sourceProof, 'providerFunction'),
            adapterSurface: stringValue(sourceProof, 'adapterSurface'),
            target: stringValue(sourceProof, 'target'),
            formula: stringValue(sourceProof, 'formula'),
            adapterFormula: stringValue(sourceProof, 'adapterFormula'),
            diagnostics: arrayValue(sourceProof, 'diagnostics'),
          }
        : {}),
      before: propertyValue(sourceProof, 'before'),
      after: propertyValue(sourceProof, 'after'),
      afterRestore: propertyValue(sourceProof, 'afterRestore'),
      afterRestart: propertyValue(sourceProof, 'afterRestart'),
      persistedDocumentBytes: numberValue(sourceProof, 'persistedDocumentBytes'),
      toolCount: tools.length,
      tools,
      checks,
    },
    verified: booleanValue(sourceProof, 'verified'),
    limitations: stringArrayValue(sourceProof, 'limitations'),
    next: {
      docs: summary.docs,
      command,
    },
    sourceProof,
  }
}

function parseEvaluatorArgs(args: readonly string[]): ParsedEvaluatorArgs {
  let door: BiligEvaluatorDoor | undefined
  let scenario: BiligEvaluatorScenario = 'default'
  let json = true
  let markdown = false
  let list = false
  let help = false

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      throw new Error('Unexpected missing bilig-evaluate argument.')
    }
    if (arg === '--door') {
      door = parseDoor(requireNextArg(args, index, '--door'))
      index += 1
      continue
    }
    if (arg.startsWith('--door=')) {
      door = parseDoor(arg.slice('--door='.length))
      continue
    }
    if (arg === '--scenario') {
      scenario = parseScenario(requireNextArg(args, index, '--scenario'))
      index += 1
      continue
    }
    if (arg.startsWith('--scenario=')) {
      scenario = parseScenario(arg.slice('--scenario='.length))
      continue
    }
    switch (arg) {
      case '--json':
        json = true
        markdown = false
        break
      case '--markdown':
        markdown = true
        json = false
        break
      case '--list':
        list = true
        break
      case '--help':
      case '-h':
        help = true
        break
      default:
        throw new Error(`Unknown bilig-evaluate option: ${arg}`)
    }
  }

  return { door, scenario, json, markdown, list, help }
}

function parseDoor(value: string): BiligEvaluatorDoor {
  switch (value) {
    case 'workpaper':
    case 'workpaper-service':
    case 'service':
      return 'workpaper-service'
    case 'agent':
    case 'agent-mcp':
    case 'mcp':
      return 'agent-mcp'
    default:
      throw new Error(`Unknown bilig-evaluate door: ${value}`)
  }
}

function parseScenario(value: string): BiligEvaluatorScenario {
  switch (value) {
    case 'default':
    case 'basic':
      return 'default'
    case 'revenue':
    case 'revenue-plan':
      return 'revenue-plan'
    case 'provider':
    case 'provider-backed':
      return 'provider-backed'
    default:
      throw new Error(`Unknown bilig-evaluate scenario: ${value}`)
  }
}

function requireNextArg(args: readonly string[], index: number, option: string): string {
  const value = args[index + 1]
  if (value === undefined || value.startsWith('-')) {
    throw new Error(`Expected value after ${option}.`)
  }
  return value
}

function renderEvaluatorHelp(): string {
  return `Usage: bilig-evaluate --door <door> [--scenario <scenario>] [--json|--markdown]

Doors:
  workpaper-service   Prove WorkPaper edit, recalc, restore, and persistence.
  agent-mcp           Prove MCP tools, restart persistence, and readback.

Options:
  --door <door>        Door to run.
  --scenario <name>    Optional scenario. Use revenue-plan or provider-backed with agent-mcp.
  --list              Print available doors.
  --json              Print JSON proof (default).
  --markdown          Print compact Markdown proof.
  -h, --help          Show this help.
`
}

function renderDoorListMarkdown(doors: readonly BiligEvaluatorDoorSummary[]): string {
  return ['# Bilig Evaluator Doors', '', ...doors.map((door) => `- \`${door.door}\`: ${door.label}\n  \`${door.command}\``), ''].join('\n')
}

function renderProofMarkdown(proof: BiligEvaluatorProof): string {
  return [
    `# ${proof.doorName}`,
    '',
    `- Door: \`${proof.door}\``,
    `- Verified: \`${String(proof.verified)}\``,
    `- Command: \`${proof.command}\``,
    `- Docs: ${proof.next.docs}`,
    `- Evidence: \`${JSON.stringify(proof.evidence)}\``,
    '',
  ].join('\n')
}

function requireDoorSummary(door: BiligEvaluatorDoor): BiligEvaluatorDoorSummary {
  const summary = doorSummaries.find((entry) => entry.door === door)
  if (summary === undefined) {
    throw new Error(`Missing evaluator door summary for ${door}.`)
  }
  return summary
}

function commandForScenario(summary: BiligEvaluatorDoorSummary, scenario: BiligEvaluatorScenario): string {
  if (summary.door !== 'agent-mcp' || scenario === 'default') {
    return summary.command
  }
  if (scenario === 'provider-backed') {
    return 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json'
  }
  return 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json'
}

function requireRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object.`)
  }
  return value
}

function propertyValue(source: unknown, key: string): unknown {
  return requireRecord(source, 'proof')[key]
}

function recordValue(source: unknown, key: string): Record<string, unknown> {
  const value = propertyValue(source, key)
  if (value === undefined) {
    return {}
  }
  return requireRecord(value, key)
}

function normalizeBooleanRecord(source: Record<string, unknown>): Readonly<Record<string, boolean>> {
  const normalized: Record<string, boolean> = {}
  for (const [key, value] of Object.entries(source)) {
    if (typeof value === 'boolean') {
      normalized[key] = value
    }
  }
  return normalized
}

function stringValue(source: unknown, key: string): string | undefined {
  return optionalString(propertyValue(source, key))
}

function numberValue(source: unknown, key: string): number | undefined {
  return optionalNumber(propertyValue(source, key))
}

function booleanValue(source: unknown, key: string): boolean {
  return propertyValue(source, key) === true
}

function arrayValue(source: unknown, key: string): readonly unknown[] {
  const value = propertyValue(source, key)
  return Array.isArray(value) ? value : []
}

function stringArrayValue(source: unknown, key: string): readonly string[] {
  return stringArray(propertyValue(source, key))
}

function stringArray(value: unknown): readonly string[] {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === 'string') : []
}

function optionalString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined
}

function optionalNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
