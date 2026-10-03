import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { runXlsxFormulaRecalcCli } from './cli-api.js'
import type {
  WorkbookCompatibilityNamedCount,
  WorkbookCompatibilityReport,
  WorkbookCompatibilityRiskLevel,
} from './workbook-compatibility-report.js'

export const biligEvaluatorSchemaVersion = 'bilig-evaluator.v1'

export type BiligEvaluatorDoor = 'xlsx-cache' | 'workbook-compatibility'
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
  readonly target?: string | undefined
  readonly riskLevel?: WorkbookCompatibilityRiskLevel | undefined
  readonly riskReasons?: readonly string[] | undefined
  readonly unsupportedFunctions?: readonly WorkbookCompatibilityNamedCount[] | undefined
  readonly externalLinks?: number | undefined
  readonly macroModules?: number | undefined
  readonly volatileFunctions?: readonly WorkbookCompatibilityNamedCount[] | undefined
  readonly pivotTables?: number | undefined
  readonly formulaCellCount?: number | undefined
  readonly before?: unknown
  readonly after?: unknown
  readonly staleCachedFormulaCount?: number | undefined
  readonly suggestedReads?: readonly string[] | undefined
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

interface XlsxCacheDoctorFormulaSummary {
  readonly target?: string | undefined
  readonly formula?: string | undefined
  readonly cachedValue?: unknown
  readonly literalRecalculatedValue?: unknown
  readonly stale?: boolean | undefined
}

interface XlsxCacheDoctorProofSummary {
  readonly commandSucceeded?: boolean | undefined
  readonly inspectionCompleted?: boolean | undefined
  readonly recalculationCompleted?: boolean | undefined
  readonly formulaCellCount?: number | undefined
  readonly staleCachedFormulaCount?: number | undefined
  readonly suggestedReads?: readonly string[] | undefined
  readonly formulas?: readonly XlsxCacheDoctorFormulaSummary[] | undefined
  readonly [key: string]: unknown
}

const xlsxCacheDocs = 'https://proompteng.github.io/bilig/eval-xlsx-cache-doctor.html'
const workbookCompatibilityDocs = 'https://proompteng.github.io/bilig/workbook-compatibility-report.html'

const doorSummaries: readonly BiligEvaluatorDoorSummary[] = [
  {
    door: 'xlsx-cache',
    label: 'Detect and recalculate a stale cached XLSX formula without Excel or LibreOffice.',
    command: 'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
    docs: xlsxCacheDocs,
  },
  {
    door: 'workbook-compatibility',
    label: 'Inspect workbook risks before trusting Bilig in a Node service or agent workflow.',
    command: 'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json',
    docs: workbookCompatibilityDocs,
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
      throw new Error('Expected --door xlsx-cache or --door workbook-compatibility.')
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
  if (scenario !== 'default') {
    throw new Error(`Scenario "${scenario}" is only available from @bilig/workpaper with --door agent-mcp.`)
  }
  switch (door) {
    case 'xlsx-cache':
      return buildXlsxCacheEvaluatorProof()
    case 'workbook-compatibility':
      return await buildWorkbookCompatibilityEvaluatorProof()
  }
}

function buildXlsxCacheEvaluatorProof(): Omit<BiligEvaluatorProof, 'durationMs'> {
  const sourceProof = buildXlsxCacheDoctorSourceProof()
  const firstFormula = sourceProof.formulas?.[0]
  const checks = {
    commandSucceeded: sourceProof.commandSucceeded === true,
    inspectionCompleted: sourceProof.inspectionCompleted === true,
    recalculationCompleted: sourceProof.recalculationCompleted === true,
    staleCachedFormulaFound: (sourceProof.staleCachedFormulaCount ?? 0) > 0,
    readbackSuggested: (sourceProof.suggestedReads?.length ?? 0) > 0,
  }
  const verified = Object.values(checks).every(Boolean)
  const summary = requireDoorSummary('xlsx-cache')

  return {
    schemaVersion: biligEvaluatorSchemaVersion,
    door: 'xlsx-cache',
    doorName: 'XLSX stale-cache proof',
    command: summary.command,
    packageVersions: {
      '@bilig/xlsx-formula-recalc': readLocalPackageVersion(),
    },
    evidence: {
      target: firstFormula?.target,
      before: firstFormula?.cachedValue,
      after: firstFormula?.literalRecalculatedValue,
      staleCachedFormulaCount: sourceProof.staleCachedFormulaCount,
      suggestedReads: sourceProof.suggestedReads,
      checks,
    },
    verified,
    limitations: [
      'This door proves stale cached formula detection and Bilig recalculation readback; it does not prove Excel desktop parity.',
      'Use Excel, LibreOffice, or an oracle harness for macros, pivots, charts, unsupported functions, and UI-specific workbook behavior.',
    ],
    next: {
      docs: summary.docs,
      command: summary.command,
    },
    sourceProof,
  }
}

async function buildWorkbookCompatibilityEvaluatorProof(): Promise<Omit<BiligEvaluatorProof, 'durationMs'>> {
  const { buildWorkbookCompatibilityDemoBytes, buildWorkbookCompatibilityReport } = await import('./workbook-compatibility-report.js')
  const sourceProof = buildWorkbookCompatibilityReport(buildWorkbookCompatibilityDemoBytes(), {
    fileName: 'bilig-workbook-compatibility-demo.xlsx',
    inspectLimit: 'all',
  })
  const checks = {
    commandSucceeded: sourceProof.commandSucceeded,
    inspectionCompleted: sourceProof.inspectionCompleted,
    metadataScanCompleted: sourceProof.inspectionCompleted && !sourceProof.recalculationCompleted,
    riskReasonsExplainFindings: sourceProof.risk.reasons.length > 0,
    noCompatibilityScore: !reportHasCompatibilityScore(sourceProof),
    unsupportedFunctionsReported: sourceProof.findings.unsupportedFunctions.length > 0,
  }
  const verified = Object.values(checks).every(Boolean)
  const summary = requireDoorSummary('workbook-compatibility')

  return {
    schemaVersion: biligEvaluatorSchemaVersion,
    door: 'workbook-compatibility',
    doorName: 'Workbook compatibility risk report',
    command: summary.command,
    packageVersions: {
      '@bilig/xlsx-formula-recalc': readLocalPackageVersion(),
    },
    evidence: {
      riskLevel: sourceProof.risk.level,
      riskReasons: sourceProof.risk.reasons,
      unsupportedFunctions: sourceProof.findings.unsupportedFunctions,
      externalLinks: sourceProof.findings.externalLinks.count,
      macroModules: sourceProof.findings.macroModules.count,
      volatileFunctions: sourceProof.findings.volatileFunctions,
      pivotTables: sourceProof.findings.pivotTables.count,
      formulaCellCount: sourceProof.workbook.formulaCellCount,
      staleCachedFormulaCount: sourceProof.findings.staleCachedFormulas.count,
      suggestedReads: sourceProof.cacheInspection.suggestedReads,
      checks,
    },
    verified,
    limitations: sourceProof.limitations,
    next: {
      docs: summary.docs,
      command: summary.command,
    },
    sourceProof,
  }
}

function buildXlsxCacheDoctorSourceProof(): XlsxCacheDoctorProofSummary {
  let stdout = ''
  const exitCode = runXlsxFormulaRecalcCli(['--demo', '--json'], {
    commandName: 'xlsx-cache-doctor',
    stdout: (text) => {
      stdout += text
    },
  })
  if (exitCode !== 0) {
    throw new Error('xlsx-cache-doctor evaluator door failed.')
  }
  return parseXlsxCacheDoctorProof(stdout)
}

function parseXlsxCacheDoctorProof(text: string): XlsxCacheDoctorProofSummary {
  const parsed: unknown = JSON.parse(text)
  const source = requireRecord(parsed, 'xlsx-cache-doctor JSON output')
  const formulasRaw = arrayValue(source, 'formulas')
  const formulas = formulasRaw.map((formula) => {
    const record = requireRecord(formula, 'xlsx-cache-doctor formula summary')
    return {
      target: optionalString(record['target']),
      formula: optionalString(record['formula']),
      cachedValue: record['cachedValue'],
      literalRecalculatedValue: record['literalRecalculatedValue'],
      stale: optionalBoolean(record['stale']),
    }
  })

  return {
    ...source,
    commandSucceeded: optionalBoolean(source['commandSucceeded']),
    inspectionCompleted: optionalBoolean(source['inspectionCompleted']),
    recalculationCompleted: optionalBoolean(source['recalculationCompleted']),
    formulaCellCount: optionalNumber(source['formulaCellCount']),
    staleCachedFormulaCount: optionalNumber(source['staleCachedFormulaCount']),
    suggestedReads: stringArray(source['suggestedReads']),
    formulas,
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
    case 'xlsx':
    case 'xlsx-cache':
    case 'xlsx-cache-doctor':
      return 'xlsx-cache'
    case 'workbook':
    case 'compatibility':
    case 'workbook-compatibility':
    case 'workbook-compatibility-report':
      return 'workbook-compatibility'
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
  return `Usage: bilig-evaluate --door <door> [--json|--markdown]

Doors:
  xlsx-cache          Detect stale cached XLSX formulas and prove recalculated readback.
  workbook-compatibility
                      Inspect workbook risks without claiming Excel compatibility.

Options:
  --door <door>        Door to run.
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

function reportHasCompatibilityScore(report: WorkbookCompatibilityReport): boolean {
  const text = JSON.stringify(report)
  return /compatibilityScore|excelCompatibilityPercent/iu.test(text)
}

function requireDoorSummary(door: BiligEvaluatorDoor): BiligEvaluatorDoorSummary {
  const summary = doorSummaries.find((entry) => entry.door === door)
  if (summary === undefined) {
    throw new Error(`Missing evaluator door summary for ${door}.`)
  }
  return summary
}

function readLocalPackageVersion(): string {
  const packageJsonPath = join(dirname(fileURLToPath(import.meta.url)), '..', 'package.json')
  const parsed: unknown = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  const record = requireRecord(parsed, packageJsonPath)
  const version = record['version']
  if (typeof version !== 'string' || version.length === 0) {
    throw new Error(`Expected package version in ${packageJsonPath}.`)
  }
  return version
}

function requireRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object.`)
  }
  return value
}

function arrayValue(source: unknown, key: string): readonly unknown[] {
  const value = requireRecord(source, 'source')[key]
  return Array.isArray(value) ? value : []
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

function optionalBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
