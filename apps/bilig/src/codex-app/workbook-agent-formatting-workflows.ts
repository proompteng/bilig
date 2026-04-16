import type { WorkbookAgentCommand } from '@bilig/agent-api'
import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowTemplate,
} from '@bilig/contracts'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkbookSnapshot } from '@bilig/protocol'
import type { ZeroSyncService } from '../zero/service.js'
import { throwIfWorkflowCancelled } from './workbook-agent-workflow-abort.js'

type FormattingWorkflowTemplate = 'highlightCurrentSheetOutliers' | 'styleCurrentSheetHeaders'
type SnapshotSheet = WorkbookSnapshot['sheets'][number]
type SnapshotCell = SnapshotSheet['cells'][number]

const DEFAULT_OUTLIER_LIMIT = 25
const MIN_NUMERIC_SAMPLES = 4

export interface FormattingWorkflowExecutionInput {
  readonly sheetName?: string
  readonly limit?: number
}

interface FormattingWorkflowStepPlan {
  readonly stepId: string
  readonly label: string
  readonly runningSummary: string
  readonly pendingSummary: string
}

interface FormattingWorkflowStepResult {
  readonly stepId: string
  readonly label: string
  readonly summary: string
}

interface FormattingWorkflowTemplateMetadata {
  readonly title: string
  readonly runningSummary: string
  readonly stepPlans: readonly FormattingWorkflowStepPlan[]
}

export interface FormattingWorkflowExecutionResult {
  readonly title: string
  readonly summary: string
  readonly artifact: WorkbookAgentWorkflowArtifact
  readonly citations: readonly WorkbookAgentTimelineCitation[]
  readonly steps: readonly FormattingWorkflowStepResult[]
  readonly commands?: readonly WorkbookAgentCommand[]
  readonly goalText?: string
}

interface SheetExtents {
  readonly headerRow: number
  readonly dataStartRow: number
  readonly minCol: number
  readonly maxCol: number
  readonly maxRow: number
  readonly cellByAddress: Map<string, SnapshotCell>
}

interface NumericSample {
  readonly address: string
  readonly value: number
}

interface ColumnOutlierReport {
  readonly columnLabel: string
  readonly sampleCount: number
  readonly lowerBound: number
  readonly upperBound: number
  readonly outliers: readonly NumericSample[]
}

function resolveWorkflowSheetName(input: {
  readonly workflowInput?: FormattingWorkflowExecutionInput | null
  readonly context?: WorkbookAgentUiContext | null
}): string | null {
  const explicitName = input.workflowInput?.sheetName?.trim()
  if (explicitName && explicitName.length > 0) {
    return explicitName
  }
  return input.context?.selection.sheetName ?? null
}

function normalizeColumnLabel(label: string): string {
  const collapsed = label.replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim()
  if (collapsed.length === 0) {
    return ''
  }
  return collapsed
    .split(' ')
    .map((part) => `${part.slice(0, 1).toUpperCase()}${part.slice(1)}`)
    .join(' ')
}

function inspectSheet(sheet: SnapshotSheet): SheetExtents {
  let minRow = Number.POSITIVE_INFINITY
  let minCol = Number.POSITIVE_INFINITY
  let maxRow = Number.NEGATIVE_INFINITY
  let maxCol = Number.NEGATIVE_INFINITY
  const cellByAddress = new Map(sheet.cells.map((cell) => [cell.address, cell] as const))
  for (const cell of sheet.cells) {
    const parsed = parseCellAddress(cell.address, sheet.name)
    minRow = Math.min(minRow, parsed.row)
    minCol = Math.min(minCol, parsed.col)
    maxRow = Math.max(maxRow, parsed.row)
    maxCol = Math.max(maxCol, parsed.col)
  }
  return {
    headerRow: minRow,
    dataStartRow: minRow + 1,
    minCol,
    maxCol,
    maxRow,
    cellByAddress,
  }
}

function quantile(sortedValues: readonly number[], percentile: number): number {
  if (sortedValues.length === 0) {
    return 0
  }
  if (sortedValues.length === 1) {
    return sortedValues[0] ?? 0
  }
  const clampedPercentile = Math.max(0, Math.min(1, percentile))
  const index = (sortedValues.length - 1) * clampedPercentile
  const lowerIndex = Math.floor(index)
  const upperIndex = Math.ceil(index)
  const lower = sortedValues[lowerIndex] ?? sortedValues[0] ?? 0
  const upper = sortedValues[upperIndex] ?? sortedValues[sortedValues.length - 1] ?? 0
  if (lowerIndex === upperIndex) {
    return lower
  }
  const weight = index - lowerIndex
  return lower + (upper - lower) * weight
}

function describeColumnLabel(headerCell: SnapshotCell | undefined, headerRow: number, col: number): string {
  if (typeof headerCell?.value === 'string') {
    const normalized = normalizeColumnLabel(headerCell.value)
    if (normalized.length > 0) {
      return normalized
    }
  }
  return formatAddress(headerRow, col).replace(/[0-9]+$/u, '')
}

function formatNumericValue(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(2)
}

function createOutlierCitation(sheetName: string, address: string, role: 'source' | 'target'): WorkbookAgentTimelineCitation {
  return {
    kind: 'range',
    sheetName,
    startAddress: address,
    endAddress: address,
    role,
  }
}

function summarizeOutlierMarkdown(input: {
  readonly sheetName: string
  readonly analyzedColumnCount: number
  readonly reports: readonly ColumnOutlierReport[]
  readonly highlightedOutlierCount: number
  readonly truncated: boolean
  readonly limit: number
}): string {
  const lines = [
    '## Highlighted Numeric Outliers',
    '',
    `Sheet: ${input.sheetName}`,
    `Numeric columns analyzed: ${String(input.analyzedColumnCount)}`,
    `Outliers highlighted: ${String(input.highlightedOutlierCount)}`,
  ]
  if (input.truncated) {
    lines.push(`Highlighted outlier sample limit: ${String(input.limit)}.`)
  }
  lines.push('')
  if (input.highlightedOutlierCount === 0) {
    lines.push('No numeric outliers were detected on the current sheet.')
    return lines.join('\n')
  }
  lines.push('### Highlighted cells')
  for (const report of input.reports) {
    if (report.outliers.length === 0) {
      continue
    }
    lines.push(
      `- ${report.columnLabel}: ${report.outliers
        .map((sample) => `${sample.address} = ${formatNumericValue(sample.value)}`)
        .join(', ')} (outside ${formatNumericValue(report.lowerBound)} to ${formatNumericValue(report.upperBound)})`,
    )
  }
  lines.push(
    '',
    'The staged workbook change set applies visible formatting to the outlier cells so reviewers can inspect and apply the change through the normal workbook mutation path.',
  )
  return lines.join('\n')
}

export function getFormattingWorkflowTemplateMetadata(
  workflowTemplate: WorkbookAgentWorkflowTemplate | FormattingWorkflowTemplate,
  workflowInput?: FormattingWorkflowExecutionInput | null,
): FormattingWorkflowTemplateMetadata | null {
  const scopeLabel = workflowInput?.sheetName ?? 'the active sheet'
  if (workflowTemplate === 'highlightCurrentSheetOutliers') {
    return {
      title: 'Highlight Current Sheet Outliers',
      runningSummary: `Running outlier highlight workflow for ${scopeLabel}.`,
      stepPlans: [
        {
          stepId: 'inspect-numeric-columns',
          label: 'Inspect numeric columns',
          runningSummary: `Inspecting numeric columns and header labels on ${scopeLabel}.`,
          pendingSummary: `Waiting to inspect numeric columns and header labels on ${scopeLabel}.`,
        },
        {
          stepId: 'compute-outlier-thresholds',
          label: 'Compute outlier thresholds',
          runningSummary: 'Computing numeric outlier thresholds across the inspected columns.',
          pendingSummary: 'Waiting to compute numeric outlier thresholds across the inspected columns.',
        },
        {
          stepId: 'stage-outlier-highlights',
          label: 'Stage outlier highlights',
          runningSummary: 'Staging semantic highlight commands for the detected outlier cells.',
          pendingSummary: 'Waiting to stage semantic highlight commands for the detected outlier cells.',
        },
        {
          stepId: 'draft-outlier-report',
          label: 'Draft outlier report',
          runningSummary: 'Drafting the durable outlier highlight report.',
          pendingSummary: 'Waiting to assemble the durable outlier highlight report.',
        },
      ],
    }
  }
  if (workflowTemplate === 'styleCurrentSheetHeaders') {
    return {
      title: 'Style Current Sheet Headers',
      runningSummary: `Running header-style workflow for ${scopeLabel}.`,
      stepPlans: [
        {
          stepId: 'inspect-header-row',
          label: 'Inspect header row',
          runningSummary: `Inspecting the used range and header row on ${scopeLabel}.`,
          pendingSummary: `Waiting to inspect the used range and header row on ${scopeLabel}.`,
        },
        {
          stepId: 'stage-header-style',
          label: 'Stage header style',
          runningSummary: 'Staging a semantic header-style change set for the current sheet.',
          pendingSummary: 'Waiting to stage a semantic header-style change set for the current sheet.',
        },
        {
          stepId: 'draft-header-style-report',
          label: 'Draft header style report',
          runningSummary: 'Drafting the durable header-style report.',
          pendingSummary: 'Waiting to assemble the durable header-style report.',
        },
      ],
    }
  }
  return null
}

export async function executeFormattingWorkflow(input: {
  readonly documentId: string
  readonly zeroSyncService: ZeroSyncService
  readonly workflowTemplate: WorkbookAgentWorkflowTemplate | FormattingWorkflowTemplate
  readonly context?: WorkbookAgentUiContext | null
  readonly workflowInput?: FormattingWorkflowExecutionInput | null
  readonly signal?: AbortSignal
}): Promise<FormattingWorkflowExecutionResult | null> {
  if (input.workflowTemplate !== 'highlightCurrentSheetOutliers' && input.workflowTemplate !== 'styleCurrentSheetHeaders') {
    return null
  }
  const sheetName = resolveWorkflowSheetName({
    ...(input.workflowInput !== undefined ? { workflowInput: input.workflowInput } : {}),
    ...(input.context !== undefined ? { context: input.context } : {}),
  })
  if (!sheetName) {
    throw new Error('Selection context is required for outlier highlight workflows.')
  }
  const limit = Math.max(1, input.workflowInput?.limit ?? DEFAULT_OUTLIER_LIMIT)
  throwIfWorkflowCancelled(input.signal)
  return await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) => {
    throwIfWorkflowCancelled(input.signal)
    const snapshot = runtime.engine.exportSnapshot()
    const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
    if (!sheet) {
      throw new Error(`Sheet ${sheetName} was not found in the workbook.`)
    }
    if (sheet.cells.length === 0) {
      if (input.workflowTemplate === 'styleCurrentSheetHeaders') {
        return {
          title: 'Style Current Sheet Headers',
          summary: `${sheetName} is empty, so there were no headers to style.`,
          artifact: {
            kind: 'markdown',
            title: 'Header Style Preview',
            text: ['## Header Style Preview', '', `Sheet: ${sheetName}`, '', 'Sheet is empty. Header styling change count: 0.'].join('\n'),
          },
          citations: [],
          steps: [
            {
              stepId: 'inspect-header-row',
              label: 'Inspect header row',
              summary: `Loaded ${sheetName} and found no populated cells.`,
            },
            {
              stepId: 'stage-header-style',
              label: 'Stage header style',
              summary: 'Sheet is empty. Header styling change count: 0.',
            },
            {
              stepId: 'draft-header-style-report',
              label: 'Draft header style report',
              summary: 'Prepared the durable empty-sheet header-style report for the thread.',
            },
          ],
        } satisfies FormattingWorkflowExecutionResult
      }
      return {
        title: 'Highlight Current Sheet Outliers',
        summary: `${sheetName} is empty, so there were no numeric outliers to highlight.`,
        artifact: {
          kind: 'markdown',
          title: 'Current Sheet Outlier Highlights',
          text: ['## Highlighted Numeric Outliers', '', `Sheet: ${sheetName}`, '', 'Sheet is empty. Outlier highlight count: 0.'].join(
            '\n',
          ),
        },
        citations: [],
        steps: [
          {
            stepId: 'inspect-numeric-columns',
            label: 'Inspect numeric columns',
            summary: `Loaded ${sheetName} and found no populated cells.`,
          },
          {
            stepId: 'compute-outlier-thresholds',
            label: 'Compute outlier thresholds',
            summary: 'No numeric outlier thresholds were computed because the sheet is empty.',
          },
          {
            stepId: 'stage-outlier-highlights',
            label: 'Stage outlier highlights',
            summary: 'Sheet is empty. Outlier highlight count: 0.',
          },
          {
            stepId: 'draft-outlier-report',
            label: 'Draft outlier report',
            summary: 'Prepared the durable empty-sheet outlier report for the thread.',
          },
        ],
      } satisfies FormattingWorkflowExecutionResult
    }

    const { headerRow, dataStartRow, minCol, maxCol, maxRow, cellByAddress } = inspectSheet(sheet)
    if (input.workflowTemplate === 'styleCurrentSheetHeaders') {
      const headerStartAddress = formatAddress(headerRow, minCol)
      const headerEndAddress = formatAddress(headerRow, maxCol)
      return {
        title: 'Style Current Sheet Headers',
        summary: `Prepared a consistent header style change set for ${sheetName}.`,
        artifact: {
          kind: 'markdown',
          title: 'Header Style Preview',
          text: [
            '## Header Style Preview',
            '',
            `Sheet: ${sheetName}`,
            `Header row: ${headerStartAddress}:${headerEndAddress}`,
            '',
            'The staged workbook change set applies a bold header style with stronger contrast and bottom borders so the sheet is easier to review.',
          ].join('\n'),
        },
        citations: [
          {
            kind: 'range',
            sheetName,
            startAddress: headerStartAddress,
            endAddress: headerEndAddress,
            role: 'target',
          },
        ],
        steps: [
          {
            stepId: 'inspect-header-row',
            label: 'Inspect header row',
            summary: `Loaded the used range and header row from ${sheetName}.`,
          },
          {
            stepId: 'stage-header-style',
            label: 'Stage header style',
            summary: 'Prepared a semantic format change set for the current sheet header row.',
          },
          {
            stepId: 'draft-header-style-report',
            label: 'Draft header style report',
            summary: 'Prepared the durable header-style report for the thread.',
          },
        ],
        commands: [
          {
            kind: 'formatRange' as const,
            range: {
              sheetName,
              startAddress: headerStartAddress,
              endAddress: headerEndAddress,
            },
            patch: {
              fill: {
                backgroundColor: '#E2E8F0',
              },
              font: {
                bold: true,
                color: '#0F172A',
              },
              alignment: {
                horizontal: 'center',
                vertical: 'middle',
                wrap: true,
              },
              borders: {
                bottom: {
                  style: 'solid',
                  weight: 'medium',
                  color: '#94A3B8',
                },
              },
            },
          },
        ],
        goalText: `Apply consistent header styling on ${sheetName}`,
      } satisfies FormattingWorkflowExecutionResult
    }
    const reports: ColumnOutlierReport[] = []
    let numericSampleCount = 0
    for (let col = minCol; col <= maxCol; col += 1) {
      throwIfWorkflowCancelled(input.signal)
      const headerCell = cellByAddress.get(formatAddress(headerRow, col))
      const numericSamples: NumericSample[] = []
      for (let row = dataStartRow; row <= maxRow; row += 1) {
        const cell = cellByAddress.get(formatAddress(row, col))
        if (typeof cell?.value === 'number' && Number.isFinite(cell.value)) {
          numericSamples.push({
            address: formatAddress(row, col),
            value: cell.value,
          })
        }
      }
      numericSampleCount += numericSamples.length
      if (numericSamples.length < MIN_NUMERIC_SAMPLES) {
        continue
      }
      const sortedValues = numericSamples.map((sample) => sample.value).toSorted((left, right) => left - right)
      const q1 = quantile(sortedValues, 0.25)
      const q3 = quantile(sortedValues, 0.75)
      const iqr = q3 - q1
      const lowerBound = q1 - iqr * 1.5
      const upperBound = q3 + iqr * 1.5
      const outliers = numericSamples.filter((sample) => sample.value < lowerBound || sample.value > upperBound)
      if (outliers.length === 0) {
        continue
      }
      reports.push({
        columnLabel: describeColumnLabel(headerCell, headerRow, col),
        sampleCount: numericSamples.length,
        lowerBound,
        upperBound,
        outliers,
      })
    }

    const flattenedOutliers = reports.flatMap((report) =>
      report.outliers.map((sample) => ({
        ...sample,
        columnLabel: report.columnLabel,
        sampleCount: report.sampleCount,
        lowerBound: report.lowerBound,
        upperBound: report.upperBound,
        severity: sample.value > report.upperBound ? sample.value - report.upperBound : report.lowerBound - sample.value,
      })),
    )
    const stagedOutliers = flattenedOutliers.toSorted((left, right) => right.severity - left.severity).slice(0, limit)
    const stagedOutlierByAddress = new Set(stagedOutliers.map((sample) => sample.address))
    const stagedReports = reports
      .map((report) => ({
        ...report,
        outliers: report.outliers.filter((sample) => stagedOutlierByAddress.has(sample.address)),
      }))
      .filter((report) => report.outliers.length > 0)
    const highlightedOutlierCount = stagedOutliers.length
    const analyzedColumnCount = reports.length
    const truncated = flattenedOutliers.length > highlightedOutlierCount

    return {
      title: 'Highlight Current Sheet Outliers',
      summary:
        highlightedOutlierCount === 0
          ? `Scanned ${String(analyzedColumnCount)} numeric column${analyzedColumnCount === 1 ? '' : 's'} on ${sheetName} and found no outliers to highlight.`
          : `Staged outlier highlights for ${String(highlightedOutlierCount)} cell${highlightedOutlierCount === 1 ? '' : 's'} across ${String(stagedReports.length)} numeric column${stagedReports.length === 1 ? '' : 's'} on ${sheetName}.`,
      artifact: {
        kind: 'markdown',
        title: 'Current Sheet Outlier Highlights',
        text: summarizeOutlierMarkdown({
          sheetName,
          analyzedColumnCount,
          reports: stagedReports,
          highlightedOutlierCount,
          truncated,
          limit,
        }),
      },
      citations: stagedOutliers.map((sample) => createOutlierCitation(sheetName, sample.address, 'target')),
      steps: [
        {
          stepId: 'inspect-numeric-columns',
          label: 'Inspect numeric columns',
          summary:
            numericSampleCount === 0
              ? `Loaded ${sheetName} and found no numeric cells to inspect for outliers.`
              : `Loaded ${String(numericSampleCount)} numeric cell${numericSampleCount === 1 ? '' : 's'} across ${String(analyzedColumnCount)} numeric column${analyzedColumnCount === 1 ? '' : 's'} on ${sheetName}.`,
        },
        {
          stepId: 'compute-outlier-thresholds',
          label: 'Compute outlier thresholds',
          summary:
            highlightedOutlierCount === 0
              ? `Computed outlier thresholds on ${String(analyzedColumnCount)} numeric column${analyzedColumnCount === 1 ? '' : 's'} and found no outliers.`
              : `Computed outlier thresholds and found ${String(flattenedOutliers.length)} outlier cell${flattenedOutliers.length === 1 ? '' : 's'} on ${sheetName}.`,
        },
        {
          stepId: 'stage-outlier-highlights',
          label: 'Stage outlier highlights',
          summary:
            highlightedOutlierCount === 0
              ? 'Outlier scan complete. Highlight count: 0.'
              : `Prepared ${String(highlightedOutlierCount)} semantic formatting command${highlightedOutlierCount === 1 ? '' : 's'} to highlight numeric outliers${truncated ? ` (limited to the top ${String(limit)} cells)` : ''}.`,
        },
        {
          stepId: 'draft-outlier-report',
          label: 'Draft outlier report',
          summary: 'Prepared the durable outlier highlight report for the thread.',
        },
      ],
      commands: stagedOutliers.map((sample) => ({
        kind: 'formatRange' as const,
        range: {
          sheetName,
          startAddress: sample.address,
          endAddress: sample.address,
        },
        patch: {
          fill: {
            backgroundColor: '#FEF3C7',
          },
          font: {
            bold: true,
            color: '#92400E',
          },
        },
      })),
      goalText: `Highlight numeric outliers on ${sheetName}`,
    } satisfies FormattingWorkflowExecutionResult
  })
}
