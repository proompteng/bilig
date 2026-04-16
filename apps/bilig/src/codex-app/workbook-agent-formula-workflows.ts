import { parseCellAddress, translateFormulaReferences } from '@bilig/formula'
import type { WorkbookAgentCommand } from '@bilig/agent-api'
import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowTemplate,
} from '@bilig/contracts'
import type { ZeroSyncService } from '../zero/service.js'
import { findWorkbookFormulaIssues, type WorkbookFormulaIssue } from './workbook-agent-comprehension.js'
import { throwIfWorkflowCancelled } from './workbook-agent-workflow-abort.js'

export interface FormulaWorkflowExecutionInput {
  readonly sheetName?: string
  readonly limit?: number
}

const MAX_FORMULA_REPAIR_ROW_DISTANCE = 5

interface FormulaWorkflowStepPlan {
  readonly stepId: string
  readonly label: string
  readonly runningSummary: string
  readonly pendingSummary: string
}

interface FormulaWorkflowStepResult {
  readonly stepId: string
  readonly label: string
  readonly summary: string
}

interface FormulaWorkflowTemplateMetadata {
  readonly title: string
  readonly runningSummary: string
  readonly stepPlans: readonly FormulaWorkflowStepPlan[]
}

export interface FormulaWorkflowExecutionResult {
  readonly title: string
  readonly summary: string
  readonly artifact: WorkbookAgentWorkflowArtifact
  readonly citations: readonly WorkbookAgentTimelineCitation[]
  readonly steps: readonly FormulaWorkflowStepResult[]
  readonly commands?: readonly WorkbookAgentCommand[]
  readonly goalText?: string
}

interface FormulaRepairRecommendation {
  readonly issue: WorkbookFormulaIssue
  readonly sourceAddress: string
  readonly sourceFormula: string
  readonly repairedFormula: string
}

interface FormulaRepairSkip {
  readonly issue: WorkbookFormulaIssue
  readonly reason: string
}

function summarizeFormulaIssueKinds(issueKinds: readonly ('error' | 'cycle' | 'unsupported')[]): string {
  return issueKinds.join(', ')
}

function summarizeFormulaIssuesMarkdown(report: ReturnType<typeof findWorkbookFormulaIssues>): string {
  const lines = [
    '## Formula Issues',
    '',
    `Scanned formula cells: ${String(report.summary.scannedFormulaCells)}`,
    `Issues found: ${String(report.summary.issueCount)}`,
    `Errors: ${String(report.summary.errorCount)}`,
    `Cycles: ${String(report.summary.cycleCount)}`,
    `JS-only fallbacks: ${String(report.summary.unsupportedCount)}`,
  ]
  if (report.summary.truncated) {
    lines.push('Showing the highest-risk issues from the requested limit.')
  }
  lines.push('')
  if (report.issues.length === 0) {
    lines.push('No formula issues were detected in the current workbook.')
    return lines.join('\n')
  }
  lines.push('### Highest-Risk Issues')
  report.issues.forEach((issue) => {
    const valueSuffix = issue.valueText.length > 0 ? ` -> ${issue.valueText}` : ''
    const errorSuffix = issue.errorText ? ` (${issue.errorText})` : ''
    lines.push(
      `- ${issue.sheetName}!${issue.address}: =${issue.formula}${valueSuffix} [${summarizeFormulaIssueKinds(issue.issueKinds)}]${errorSuffix}`,
    )
  })
  return lines.join('\n')
}

function summarizeHighlightedFormulaIssuesMarkdown(report: ReturnType<typeof findWorkbookFormulaIssues>): string {
  const lines = [
    '## Highlighted Formula Issues',
    '',
    `Scanned formula cells: ${String(report.summary.scannedFormulaCells)}`,
    `Issues highlighted: ${String(report.issues.length)}`,
  ]
  if (report.issues.length === 0) {
    lines.push('', 'Formula scan complete. Issue count: 0.')
    return lines.join('\n')
  }
  lines.push('', '### Highlighted Cells')
  report.issues.forEach((issue) => {
    lines.push(`- ${issue.sheetName}!${issue.address} [${summarizeFormulaIssueKinds(issue.issueKinds)}]`)
  })
  lines.push(
    '',
    'The workbook change set applies visible formatting to the listed cells. Review flows surface it in the workbook panel, and autonomous sessions apply it directly.',
  )
  return lines.join('\n')
}

function summarizeRepairFormulaIssuesMarkdown(input: {
  readonly report: ReturnType<typeof findWorkbookFormulaIssues>
  readonly repaired: readonly FormulaRepairRecommendation[]
  readonly skipped: readonly FormulaRepairSkip[]
}): string {
  const lines = [
    '## Formula Repair Preview',
    '',
    `Scanned formula cells: ${String(input.report.summary.scannedFormulaCells)}`,
    `Issues found: ${String(input.report.summary.issueCount)}`,
    `Repairs staged: ${String(input.repaired.length)}`,
    `Issues skipped: ${String(input.skipped.length)}`,
    '',
  ]
  if (input.repaired.length === 0) {
    lines.push('No safe formula repairs were inferred from neighboring healthy formulas.')
  } else {
    lines.push('### Staged repairs')
    input.repaired.forEach((repair) => {
      lines.push(
        `- ${repair.issue.sheetName}!${repair.issue.address}: replace ${repair.issue.formula} with =${repair.repairedFormula} using ${repair.sourceAddress} (${repair.sourceFormula}) as the source pattern`,
      )
    })
  }
  if (input.skipped.length > 0) {
    lines.push('', '### Skipped issues')
    input.skipped.forEach((skip) => {
      lines.push(`- ${skip.issue.sheetName}!${skip.issue.address}: ${skip.reason}`)
    })
  }
  lines.push(
    '',
    `Only issue cells with a nearby healthy formula in the same column, within ${String(MAX_FORMULA_REPAIR_ROW_DISTANCE)} rows, are staged for repair.`,
  )
  return lines.join('\n')
}

function createIssueCitations(
  report: ReturnType<typeof findWorkbookFormulaIssues>,
  role: 'source' | 'target',
): WorkbookAgentTimelineCitation[] {
  return report.issues.map((issue) => ({
    kind: 'range',
    sheetName: issue.sheetName,
    startAddress: issue.address,
    endAddress: issue.address,
    role,
  }))
}

function createRepairCitations(repaired: readonly FormulaRepairRecommendation[]): WorkbookAgentTimelineCitation[] {
  return repaired.flatMap((repair) => [
    {
      kind: 'range' as const,
      sheetName: repair.issue.sheetName,
      startAddress: repair.sourceAddress,
      endAddress: repair.sourceAddress,
      role: 'source' as const,
    },
    {
      kind: 'range' as const,
      sheetName: repair.issue.sheetName,
      startAddress: repair.issue.address,
      endAddress: repair.issue.address,
      role: 'target' as const,
    },
  ])
}

function resolveWorkflowSheetName(input: {
  readonly workflowInput?: FormulaWorkflowExecutionInput | null
  readonly context?: WorkbookAgentUiContext | null
}): string | null {
  const explicitName = input.workflowInput?.sheetName?.trim()
  if (explicitName && explicitName.length > 0) {
    return explicitName
  }
  return input.context?.selection.sheetName ?? null
}

function buildFormulaRepairPlan(input: {
  readonly runtime: {
    readonly engine: {
      exportSnapshot: () => {
        readonly sheets: readonly {
          readonly name: string
          readonly cells: readonly {
            readonly address: string
            readonly formula?: string
          }[]
        }[]
      }
    }
  }
  readonly report: ReturnType<typeof findWorkbookFormulaIssues>
}): {
  readonly repaired: readonly FormulaRepairRecommendation[]
  readonly skipped: readonly FormulaRepairSkip[]
} {
  const snapshot = input.runtime.engine.exportSnapshot()
  const sheetByName = new Map(snapshot.sheets.map((sheet) => [sheet.name, sheet] as const))
  const issueAddressSet = new Set(input.report.issues.map((issue) => `${issue.sheetName}!${issue.address}`))
  const repaired: FormulaRepairRecommendation[] = []
  const skipped: FormulaRepairSkip[] = []

  for (const issue of input.report.issues) {
    if (issue.issueKinds.includes('cycle')) {
      skipped.push({
        issue,
        reason: 'Skipped because cyclic formulas need manual review instead of inferred rewrites.',
      })
      continue
    }
    const sheet = sheetByName.get(issue.sheetName)
    if (!sheet) {
      skipped.push({
        issue,
        reason: 'Skipped because the durable snapshot for the issue sheet was unavailable.',
      })
      continue
    }
    const target = parseCellAddress(issue.address, issue.sheetName)
    const candidate = sheet.cells
      .flatMap((cell) => {
        if (!cell.formula || cell.address === issue.address) {
          return []
        }
        if (issueAddressSet.has(`${issue.sheetName}!${cell.address}`)) {
          return []
        }
        const source = parseCellAddress(cell.address, issue.sheetName)
        if (source.col !== target.col) {
          return []
        }
        const rowDistance = Math.abs(source.row - target.row)
        if (rowDistance === 0 || rowDistance > MAX_FORMULA_REPAIR_ROW_DISTANCE) {
          return []
        }
        return [{ cell, source, rowDistance }] as const
      })
      .toSorted((left, right) => {
        if (left.rowDistance !== right.rowDistance) {
          return left.rowDistance - right.rowDistance
        }
        const leftAbove = left.source.row < target.row ? 0 : 1
        const rightAbove = right.source.row < target.row ? 0 : 1
        if (leftAbove !== rightAbove) {
          return leftAbove - rightAbove
        }
        return left.source.row - right.source.row
      })[0]
    if (!candidate) {
      skipped.push({
        issue,
        reason: `Skipped because there is no healthy formula in the same column within ${String(MAX_FORMULA_REPAIR_ROW_DISTANCE)} rows.`,
      })
      continue
    }
    try {
      const sourceFormula = candidate.cell.formula
      if (!sourceFormula) {
        skipped.push({
          issue,
          reason: `Skipped because the source formula at ${candidate.cell.address} is missing.`,
        })
        continue
      }
      const repairedFormula = translateFormulaReferences(
        sourceFormula,
        target.row - candidate.source.row,
        target.col - candidate.source.col,
      )
      if (`=${repairedFormula}` === issue.formula) {
        skipped.push({
          issue,
          reason: `Skipped because the translated formula from ${candidate.cell.address} matches the current formula text.`,
        })
        continue
      }
      repaired.push({
        issue,
        sourceAddress: candidate.cell.address,
        sourceFormula: `=${sourceFormula}`,
        repairedFormula,
      })
    } catch (error) {
      skipped.push({
        issue,
        reason:
          error instanceof Error && error.message.length > 0
            ? `Skipped because translating the source formula from ${candidate.cell.address} failed: ${error.message}`
            : `Skipped because translating the source formula from ${candidate.cell.address} failed.`,
      })
    }
  }

  return { repaired, skipped }
}

export function getFormulaWorkflowTemplateMetadata(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  workflowInput?: FormulaWorkflowExecutionInput | null,
): FormulaWorkflowTemplateMetadata | null {
  const scopeLabel = workflowInput?.sheetName ?? 'the workbook'
  if (workflowTemplate === 'findFormulaIssues') {
    return {
      title: 'Find Formula Issues',
      runningSummary: `Running formula issue scan workflow for ${scopeLabel}.`,
      stepPlans: [
        {
          stepId: 'scan-formula-cells',
          label: 'Scan formula cells',
          runningSummary: workflowInput?.sheetName
            ? `Scanning ${workflowInput.sheetName} formulas for errors, cycles, and JS-only fallbacks.`
            : 'Scanning workbook formulas for errors, cycles, and JS-only fallbacks.',
          pendingSummary: workflowInput?.sheetName
            ? `Waiting to scan ${workflowInput.sheetName} formulas for errors, cycles, and JS-only fallbacks.`
            : 'Waiting to scan workbook formulas for errors, cycles, and JS-only fallbacks.',
        },
        {
          stepId: 'draft-issue-report',
          label: 'Draft issue report',
          runningSummary: 'Drafting the durable formula issue report.',
          pendingSummary: 'Waiting to assemble the durable formula issue report.',
        },
      ],
    }
  }
  if (workflowTemplate === 'highlightFormulaIssues') {
    return {
      title: 'Highlight Formula Issues',
      runningSummary: `Running formula highlight workflow for ${scopeLabel}.`,
      stepPlans: [
        {
          stepId: 'scan-formula-cells',
          label: 'Scan formula cells',
          runningSummary: workflowInput?.sheetName
            ? `Scanning ${workflowInput.sheetName} formulas for errors, cycles, and JS-only fallbacks.`
            : 'Scanning workbook formulas for errors, cycles, and JS-only fallbacks.',
          pendingSummary: workflowInput?.sheetName
            ? `Waiting to scan ${workflowInput.sheetName} formulas for errors, cycles, and JS-only fallbacks.`
            : 'Waiting to scan workbook formulas for errors, cycles, and JS-only fallbacks.',
        },
        {
          stepId: 'stage-issue-highlights',
          label: 'Stage issue highlights',
          runningSummary: 'Staging semantic highlight commands for the detected formula issues.',
          pendingSummary: 'Waiting to stage semantic highlight commands for the detected formula issues.',
        },
        {
          stepId: 'draft-highlight-report',
          label: 'Draft highlight report',
          runningSummary: 'Drafting the durable formula highlight report.',
          pendingSummary: 'Waiting to assemble the durable formula highlight report.',
        },
      ],
    }
  }
  if (workflowTemplate === 'repairFormulaIssues') {
    return {
      title: 'Repair Formula Issues',
      runningSummary: `Running formula repair workflow for ${scopeLabel}.`,
      stepPlans: [
        {
          stepId: 'scan-formula-cells',
          label: 'Scan formula cells',
          runningSummary: workflowInput?.sheetName
            ? `Scanning ${workflowInput.sheetName} formulas for repairable errors, cycles, and JS-only fallbacks.`
            : 'Scanning workbook formulas for repairable errors, cycles, and JS-only fallbacks.',
          pendingSummary: workflowInput?.sheetName
            ? `Waiting to scan ${workflowInput.sheetName} formulas for repairable errors, cycles, and JS-only fallbacks.`
            : 'Waiting to scan workbook formulas for repairable errors, cycles, and JS-only fallbacks.',
        },
        {
          stepId: 'infer-formula-repairs',
          label: 'Infer formula repairs',
          runningSummary: 'Comparing issue cells against nearby healthy formulas in the same column.',
          pendingSummary: 'Waiting to compare issue cells against nearby healthy formulas in the same column.',
        },
        {
          stepId: 'stage-formula-repairs',
          label: 'Stage formula repairs',
          runningSummary: 'Preparing semantic write commands for the formula repairs that passed the safety checks.',
          pendingSummary: 'Waiting to prepare semantic write commands for the formula repairs that passed the safety checks.',
        },
        {
          stepId: 'draft-repair-report',
          label: 'Draft repair report',
          runningSummary: 'Drafting the durable formula repair report.',
          pendingSummary: 'Waiting to assemble the durable formula repair report.',
        },
      ],
    }
  }
  return null
}

export async function executeFormulaWorkflow(input: {
  documentId: string
  zeroSyncService: ZeroSyncService
  workflowTemplate: WorkbookAgentWorkflowTemplate
  workflowInput?: FormulaWorkflowExecutionInput | null
  context?: WorkbookAgentUiContext | null
  signal?: AbortSignal
}): Promise<FormulaWorkflowExecutionResult | null> {
  if (
    input.workflowTemplate !== 'findFormulaIssues' &&
    input.workflowTemplate !== 'highlightFormulaIssues' &&
    input.workflowTemplate !== 'repairFormulaIssues'
  ) {
    return null
  }
  throwIfWorkflowCancelled(input.signal)
  const resolvedSheetName = resolveWorkflowSheetName({
    ...(input.workflowInput !== undefined ? { workflowInput: input.workflowInput } : {}),
    ...(input.context !== undefined ? { context: input.context } : {}),
  })
  const formulaInspection = await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) => {
    const report = findWorkbookFormulaIssues(runtime, {
      ...(resolvedSheetName ? { sheetName: resolvedSheetName } : {}),
      ...(input.workflowInput?.limit !== undefined ? { limit: input.workflowInput.limit } : {}),
    })
    return {
      report,
      ...(input.workflowTemplate === 'repairFormulaIssues' ? { repairPlan: buildFormulaRepairPlan({ runtime, report }) } : {}),
    }
  })
  throwIfWorkflowCancelled(input.signal)
  const formulaIssues = formulaInspection.report
  const scopeLabel = resolvedSheetName ? ` on ${resolvedSheetName}` : ''

  if (input.workflowTemplate === 'findFormulaIssues') {
    return {
      title: 'Find Formula Issues',
      summary:
        formulaIssues.summary.issueCount === 0
          ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}${scopeLabel} and found no issues.`
          : `Found ${String(formulaIssues.summary.issueCount)} formula issue${formulaIssues.summary.issueCount === 1 ? '' : 's'}${scopeLabel} across ${String(formulaIssues.summary.scannedFormulaCells)} scanned formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}.`,
      artifact: {
        kind: 'markdown',
        title: 'Formula Issues',
        text: summarizeFormulaIssuesMarkdown(formulaIssues),
      },
      citations: createIssueCitations(formulaIssues, 'source'),
      steps: [
        {
          stepId: 'scan-formula-cells',
          label: 'Scan formula cells',
          summary:
            formulaIssues.summary.issueCount === 0
              ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}${scopeLabel} and found no issues.`
              : `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}${scopeLabel} and found ${String(formulaIssues.summary.issueCount)} issue${formulaIssues.summary.issueCount === 1 ? '' : 's'}.`,
        },
        {
          stepId: 'draft-issue-report',
          label: 'Draft issue report',
          summary: 'Prepared the durable formula issue report for the thread.',
        },
      ],
    }
  }

  if (input.workflowTemplate === 'highlightFormulaIssues') {
    return {
      title: 'Highlight Formula Issues',
      summary:
        formulaIssues.summary.issueCount === 0
          ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}${scopeLabel} and found no issues to highlight.`
          : `Staged highlight formatting for ${String(formulaIssues.summary.issueCount)} formula issue${formulaIssues.summary.issueCount === 1 ? '' : 's'}${scopeLabel}.`,
      artifact: {
        kind: 'markdown',
        title: 'Formula Issue Highlights',
        text: summarizeHighlightedFormulaIssuesMarkdown(formulaIssues),
      },
      citations: createIssueCitations(formulaIssues, 'target'),
      steps: [
        {
          stepId: 'scan-formula-cells',
          label: 'Scan formula cells',
          summary:
            formulaIssues.summary.issueCount === 0
              ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}${scopeLabel} and found no issues.`
              : `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}${scopeLabel} and found ${String(formulaIssues.summary.issueCount)} issue${formulaIssues.summary.issueCount === 1 ? '' : 's'}.`,
        },
        {
          stepId: 'stage-issue-highlights',
          label: 'Stage issue highlights',
          summary:
            formulaIssues.summary.issueCount === 0
              ? 'Formula issue scan complete. Issue count: 0.'
              : `Prepared ${String(formulaIssues.issues.length)} semantic formatting command${formulaIssues.issues.length === 1 ? '' : 's'} to highlight the detected formula issues.`,
        },
        {
          stepId: 'draft-highlight-report',
          label: 'Draft highlight report',
          summary: 'Prepared the durable formula highlight report for the thread.',
        },
      ],
      commands: formulaIssues.issues.map((issue) => ({
        kind: 'formatRange' as const,
        range: {
          sheetName: issue.sheetName,
          startAddress: issue.address,
          endAddress: issue.address,
        },
        patch: {
          fill: {
            backgroundColor: '#FEE2E2',
          },
          font: {
            bold: true,
            color: '#991B1B',
          },
        },
      })),
      goalText: resolvedSheetName ? `Highlight formula issues on ${resolvedSheetName}` : 'Highlight formula issues in the workbook',
    }
  }

  const repairPlan = formulaInspection.repairPlan ?? { repaired: [], skipped: [] }
  return {
    title: 'Repair Formula Issues',
    summary:
      repairPlan.repaired.length === 0
        ? `Scanned ${String(formulaIssues.summary.issueCount)} formula issue${formulaIssues.summary.issueCount === 1 ? '' : 's'}${scopeLabel} and found no safe repairs to stage.`
        : `Staged ${String(repairPlan.repaired.length)} formula repair${repairPlan.repaired.length === 1 ? '' : 's'}${scopeLabel} from nearby healthy formulas.`,
    artifact: {
      kind: 'markdown',
      title: 'Formula Repair Preview',
      text: summarizeRepairFormulaIssuesMarkdown({
        report: formulaIssues,
        repaired: repairPlan.repaired,
        skipped: repairPlan.skipped,
      }),
    },
    citations: createRepairCitations(repairPlan.repaired),
    steps: [
      {
        stepId: 'scan-formula-cells',
        label: 'Scan formula cells',
        summary:
          formulaIssues.summary.issueCount === 0
            ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}${scopeLabel} and found no issues.`
            : `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? '' : 's'}${scopeLabel} and found ${String(formulaIssues.summary.issueCount)} issue${formulaIssues.summary.issueCount === 1 ? '' : 's'}.`,
      },
      {
        stepId: 'infer-formula-repairs',
        label: 'Infer formula repairs',
        summary:
          repairPlan.repaired.length === 0
            ? `Compared issue cells${scopeLabel} against nearby healthy formulas and found no safe repair candidates.`
            : `Matched ${String(repairPlan.repaired.length)} issue cell${repairPlan.repaired.length === 1 ? '' : 's'}${scopeLabel} to nearby healthy formula patterns.`,
      },
      {
        stepId: 'stage-formula-repairs',
        label: 'Stage formula repairs',
        summary:
          repairPlan.repaired.length === 0
            ? 'Formula repair analysis complete. Repair-ready issues: 0.'
            : `Prepared ${String(repairPlan.repaired.length)} semantic write command${repairPlan.repaired.length === 1 ? '' : 's'} for the repair workbook change set.`,
      },
      {
        stepId: 'draft-repair-report',
        label: 'Draft repair report',
        summary: 'Prepared the durable formula repair report for the thread.',
      },
    ],
    ...(repairPlan.repaired.length > 0
      ? {
          commands: repairPlan.repaired.map(
            (repair) =>
              ({
                kind: 'writeRange' as const,
                sheetName: repair.issue.sheetName,
                startAddress: repair.issue.address,
                values: [[{ formula: repair.repairedFormula }]],
              }) satisfies WorkbookAgentCommand,
          ),
          goalText: resolvedSheetName ? `Repair formula issues on ${resolvedSheetName}` : 'Repair formula issues in the workbook',
        }
      : {}),
  }
}
