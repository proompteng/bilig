import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import { WorkPaper } from './work-paper.js'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

export interface AgentWorkbookChallengeProof {
  readonly editedCell: 'Inputs!B2'
  readonly dependentCell: 'Summary!B2'
  readonly before: number
  readonly after: number
  readonly afterRestore: number
  readonly persistedDocumentBytes: number
  readonly sheets: readonly string[]
  readonly checks: {
    readonly formulaReadbackChanged: boolean
    readonly exportedWorkPaperDocument: boolean
    readonly restoredMatchesAfter: boolean
  }
  readonly verified: boolean
  readonly limitations: readonly string[]
  readonly nextStep: string
}

export interface AgentWorkbookChallengeCliHost {
  readonly argv: readonly string[]
  readonly writeStderr?: (text: string) => void
  readonly writeStdout?: (text: string) => void
}

type AgentWorkbookChallengeOutputMode = 'json' | 'markdown'

interface AgentWorkbookChallengeCliOptions {
  readonly help: boolean
  readonly outputMode: AgentWorkbookChallengeOutputMode
}

export function runAgentWorkbookChallengeCli(host: AgentWorkbookChallengeCliHost): number {
  const writeStdout = host.writeStdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = host.writeStderr ?? ((text: string) => process.stderr.write(text))
  let options: AgentWorkbookChallengeCliOptions

  try {
    options = parseAgentWorkbookChallengeCliArgs(host.argv)
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n\n${agentWorkbookChallengeHelpText()}`)
    return 1
  }

  if (options.help) {
    writeStdout(agentWorkbookChallengeHelpText())
    return 0
  }

  const proof = buildAgentWorkbookChallengeProof()
  writeStdout(renderAgentWorkbookChallengeProof(proof, options.outputMode))
  return proof.verified ? 0 : 1
}

export function parseAgentWorkbookChallengeCliArgs(args: readonly string[]): AgentWorkbookChallengeCliOptions {
  let help = false
  let outputMode: AgentWorkbookChallengeOutputMode = 'json'

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
    throw new Error(`Unknown bilig-agent-challenge argument: ${arg}`)
  }

  return { help, outputMode }
}

export function agentWorkbookChallengeHelpText(): string {
  return [
    'Usage: bilig-agent-challenge [--json|--markdown]',
    '',
    'Runs the Bilig agent workbook challenge without cloning the repository:',
    'build a two-sheet WorkPaper, edit Inputs!B2, read Summary!B2, export JSON,',
    'restore the document, and print a proof object with verified: true.',
    '',
    'Options:',
    '  --json       Print machine-readable JSON. Default.',
    '  --markdown   Print a paste-ready Markdown report.',
    '  -h, --help   Print this help text.',
    '',
  ].join('\n')
}

export function buildAgentWorkbookChallengeProof(): AgentWorkbookChallengeProof {
  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Customers', 20],
      ['Average revenue', 1200],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Revenue', '=Inputs!B2*Inputs!B3'],
    ],
  })

  const inputs = requireSheet(workbook, 'Inputs')
  const summary = requireSheet(workbook, 'Summary')
  const summaryRevenue = { sheet: summary, row: 1, col: 1 }

  const before = numberValue(workbook.getCellValue(summaryRevenue))
  workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)

  const after = numberValue(workbook.getCellValue(summaryRevenue))
  const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
  const restoredSummary = requireSheet(restored, 'Summary')
  const afterRestore = numberValue(restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }))
  const checks = {
    formulaReadbackChanged: before !== after && before === 24_000 && after === 38_400,
    exportedWorkPaperDocument: serialized.length > 0,
    restoredMatchesAfter: afterRestore === after,
  }

  return {
    editedCell: 'Inputs!B2',
    dependentCell: 'Summary!B2',
    before,
    after,
    afterRestore,
    persistedDocumentBytes: serialized.length,
    sheets: restored.getSheetNames(),
    checks,
    verified: checks.formulaReadbackChanged && checks.exportedWorkPaperDocument && checks.restoredMatchesAfter,
    limitations: [
      'This challenge proves the WorkPaper write/read/persist loop, not full Excel desktop compatibility.',
      'For XLSX-specific behavior, run bilig-formula-clinic or the XLSX recalculation example with a real workbook fixture.',
    ],
    nextStep:
      'If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers',
  }
}

export function renderAgentWorkbookChallengeProof(
  proof: AgentWorkbookChallengeProof,
  outputMode: AgentWorkbookChallengeOutputMode,
): string {
  if (outputMode === 'markdown') {
    return renderAgentWorkbookChallengeMarkdown(proof)
  }
  return `${JSON.stringify(proof, null, 2)}\n`
}

function renderAgentWorkbookChallengeMarkdown(proof: AgentWorkbookChallengeProof): string {
  return `# Bilig agent workbook challenge

\`\`\`json
${JSON.stringify(proof, null, 2)}
\`\`\`

Result: ${proof.verified ? 'verified' : 'failed'}.

The important invariant is that \`${proof.editedCell}\` changed the dependent formula cell \`${proof.dependentCell}\`, and a serialized WorkPaper restore kept the same computed value.
`
}

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheet = workpaper.getSheetId(sheetName)
  if (sheet === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheet
}

function numberValue(cell: unknown): number {
  if (isRecord(cell) && typeof cell['value'] === 'number') {
    return cell['value']
  }
  throw new Error(`Expected numeric cell value, got ${JSON.stringify(cell)}`)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
