import type { WorkbookAgentCommand, WorkbookAgentWriteCellInput } from '@bilig/agent-api'
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

type RollupWorkflowTemplate = 'createCurrentSheetRollup' | 'createCurrentSheetReviewTab'
type SnapshotSheet = WorkbookSnapshot['sheets'][number]
type SnapshotCell = SnapshotSheet['cells'][number]

export interface RollupWorkflowExecutionInput {
  readonly sheetName?: string
}

interface RollupWorkflowStepPlan {
  readonly stepId: string
  readonly label: string
  readonly runningSummary: string
  readonly pendingSummary: string
}

interface RollupWorkflowStepResult {
  readonly stepId: string
  readonly label: string
  readonly summary: string
}

interface RollupWorkflowTemplateMetadata {
  readonly title: string
  readonly runningSummary: string
  readonly stepPlans: readonly RollupWorkflowStepPlan[]
}

export interface RollupWorkflowExecutionResult {
  readonly title: string
  readonly summary: string
  readonly artifact: WorkbookAgentWorkflowArtifact
  readonly citations: readonly WorkbookAgentTimelineCitation[]
  readonly steps: readonly RollupWorkflowStepResult[]
  readonly commands?: readonly WorkbookAgentCommand[]
  readonly goalText?: string
}

interface RollupColumnSummary {
  readonly headerLabel: string
  readonly count: number
  readonly sum: number
  readonly average: number
  readonly min: number
  readonly max: number
}

function resolveWorkflowSheetName(input: {
  readonly workflowInput?: RollupWorkflowExecutionInput | null
  readonly context?: WorkbookAgentUiContext | null
}): string | null {
  const explicitName = input.workflowInput?.sheetName?.trim()
  if (explicitName && explicitName.length > 0) {
    return explicitName
  }
  return input.context?.selection.sheetName ?? null
}

function createUniqueSheetName(existingNames: readonly string[], baseName: string): string {
  if (!existingNames.includes(baseName)) {
    return baseName
  }
  for (let suffix = 2; suffix < 1000; suffix += 1) {
    const candidate = `${baseName} ${String(suffix)}`
    if (!existingNames.includes(candidate)) {
      return candidate
    }
  }
  throw new Error(`Could not allocate a unique rollup sheet name from ${baseName}.`)
}

function inspectSheet(sheet: SnapshotSheet): {
  readonly headerRow: number
  readonly dataStartRow: number
  readonly minCol: number
  readonly maxCol: number
  readonly maxRow: number
  readonly cellByAddress: Map<string, SnapshotCell>
} {
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

function summarizeRollupMarkdown(input: {
  readonly sourceSheetName: string
  readonly targetSheetName: string
  readonly sourceStartAddress: string
  readonly sourceEndAddress: string
  readonly columns: readonly RollupColumnSummary[]
}): string {
  const lines = [
    '## Current Sheet Rollup Preview',
    '',
    `Source sheet: ${input.sourceSheetName}`,
    `Source range: ${input.sourceStartAddress}:${input.sourceEndAddress}`,
    `Target sheet: ${input.targetSheetName}`,
    `Numeric columns summarized: ${String(input.columns.length)}`,
    '',
  ]
  if (input.columns.length === 0) {
    lines.push('No numeric columns were available to summarize on the current sheet.')
    return lines.join('\n')
  }
  lines.push('### Column summaries')
  for (const column of input.columns) {
    lines.push(
      `- ${column.headerLabel}: count ${String(column.count)}, sum ${String(column.sum)}, average ${String(column.average)}, min ${String(column.min)}, max ${String(column.max)}`,
    )
  }
  lines.push(
    '',
    'The staged workbook change set creates a new rollup sheet and writes the aggregate table through the normal workbook mutation path.',
  )
  return lines.join('\n')
}

export function getRollupWorkflowTemplateMetadata(
  workflowTemplate: WorkbookAgentWorkflowTemplate | RollupWorkflowTemplate,
  workflowInput?: RollupWorkflowExecutionInput | null,
): RollupWorkflowTemplateMetadata | null {
  if (workflowTemplate !== 'createCurrentSheetRollup' && workflowTemplate !== 'createCurrentSheetReviewTab') {
    return null
  }
  const scopeLabel = workflowInput?.sheetName ?? 'the active sheet'
  if (workflowTemplate === 'createCurrentSheetReviewTab') {
    return {
      title: 'Create Current Sheet Review Tab',
      runningSummary: `Running review-tab workflow for ${scopeLabel}.`,
      stepPlans: [
        {
          stepId: 'inspect-source-sheet',
          label: 'Inspect source sheet',
          runningSummary: `Inspecting the used range on ${scopeLabel}.`,
          pendingSummary: `Waiting to inspect the used range on ${scopeLabel}.`,
        },
        {
          stepId: 'stage-review-tab-preview',
          label: 'Stage review tab preview',
          runningSummary: 'Staging the semantic change set that creates and copies the review tab.',
          pendingSummary: 'Waiting to stage the semantic change set that creates and copies the review tab.',
        },
        {
          stepId: 'draft-review-tab-report',
          label: 'Draft review tab report',
          runningSummary: 'Drafting the durable review-tab report.',
          pendingSummary: 'Waiting to assemble the durable review-tab report.',
        },
      ],
    }
  }
  return {
    title: 'Create Current Sheet Rollup',
    runningSummary: `Running rollup workflow for ${scopeLabel}.`,
    stepPlans: [
      {
        stepId: 'inspect-source-sheet',
        label: 'Inspect source sheet',
        runningSummary: `Inspecting the used range and numeric columns on ${scopeLabel}.`,
        pendingSummary: `Waiting to inspect the used range and numeric columns on ${scopeLabel}.`,
      },
      {
        stepId: 'aggregate-column-stats',
        label: 'Aggregate column stats',
        runningSummary: 'Computing numeric column counts, sums, averages, mins, and maxes.',
        pendingSummary: 'Waiting to compute numeric column counts, sums, averages, mins, and maxes.',
      },
      {
        stepId: 'stage-rollup-preview',
        label: 'Stage rollup change set',
        runningSummary: 'Staging the semantic change set that creates the rollup sheet.',
        pendingSummary: 'Waiting to stage the semantic change set that creates the rollup sheet.',
      },
      {
        stepId: 'draft-rollup-report',
        label: 'Draft rollup report',
        runningSummary: 'Drafting the durable rollup report.',
        pendingSummary: 'Waiting to assemble the durable rollup report.',
      },
    ],
  }
}

export async function executeRollupWorkflow(input: {
  readonly documentId: string
  readonly zeroSyncService: ZeroSyncService
  readonly workflowTemplate: WorkbookAgentWorkflowTemplate | RollupWorkflowTemplate
  readonly context?: WorkbookAgentUiContext | null
  readonly workflowInput?: RollupWorkflowExecutionInput | null
  readonly signal?: AbortSignal
}): Promise<RollupWorkflowExecutionResult | null> {
  if (input.workflowTemplate !== 'createCurrentSheetRollup' && input.workflowTemplate !== 'createCurrentSheetReviewTab') {
    return null
  }
  const sheetName = resolveWorkflowSheetName({
    ...(input.workflowInput !== undefined ? { workflowInput: input.workflowInput } : {}),
    ...(input.context !== undefined ? { context: input.context } : {}),
  })
  if (!sheetName) {
    throw new Error('Selection context is required for current-sheet rollup workflows.')
  }

  throwIfWorkflowCancelled(input.signal)
  return await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) => {
    throwIfWorkflowCancelled(input.signal)
    const snapshot = runtime.engine.exportSnapshot()
    const sourceSheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
    if (!sourceSheet) {
      throw new Error(`Sheet ${sheetName} was not found in the workbook.`)
    }
    if (sourceSheet.cells.length === 0) {
      if (input.workflowTemplate === 'createCurrentSheetReviewTab') {
        return {
          title: 'Create Current Sheet Review Tab',
          summary: `${sheetName} is empty, so there was no review tab content to stage.`,
          artifact: {
            kind: 'markdown',
            title: 'Current Sheet Review Tab Preview',
            text: [
              '## Current Sheet Review Tab Preview',
              '',
              `Source sheet: ${sheetName}`,
              '',
              'Sheet is empty. Review-tab rows copied: 0.',
            ].join('\n'),
          },
          citations: [],
          steps: [
            {
              stepId: 'inspect-source-sheet',
              label: 'Inspect source sheet',
              summary: `Loaded ${sheetName} and found no populated cells.`,
            },
            {
              stepId: 'stage-review-tab-preview',
              label: 'Stage review tab preview',
              summary: 'Sheet is empty. Review-tab rows copied: 0.',
            },
            {
              stepId: 'draft-review-tab-report',
              label: 'Draft review tab report',
              summary: 'Prepared the durable empty-sheet review-tab report for the thread.',
            },
          ],
        } satisfies RollupWorkflowExecutionResult
      }
      return {
        title: 'Create Current Sheet Rollup',
        summary: `${sheetName} is empty, so there were no numeric columns to roll up.`,
        artifact: {
          kind: 'markdown',
          title: 'Current Sheet Rollup Preview',
          text: [
            '## Current Sheet Rollup Preview',
            '',
            `Source sheet: ${sheetName}`,
            '',
            'Sheet is empty. Rollup columns summarized: 0.',
          ].join('\n'),
        },
        citations: [],
        steps: [
          {
            stepId: 'inspect-source-sheet',
            label: 'Inspect source sheet',
            summary: `Loaded ${sheetName} and found no populated cells.`,
          },
          {
            stepId: 'aggregate-column-stats',
            label: 'Aggregate column stats',
            summary: 'No numeric column summaries were computed because the sheet is empty.',
          },
          {
            stepId: 'stage-rollup-preview',
            label: 'Stage rollup change set',
            summary: 'Sheet is empty. Rollup columns summarized: 0.',
          },
          {
            stepId: 'draft-rollup-report',
            label: 'Draft rollup report',
            summary: 'Prepared the durable empty-sheet rollup report for the thread.',
          },
        ],
      } satisfies RollupWorkflowExecutionResult
    }

    const { headerRow, dataStartRow, minCol, maxCol, maxRow, cellByAddress } = inspectSheet(sourceSheet)
    if (input.workflowTemplate === 'createCurrentSheetReviewTab') {
      const sourceStartAddress = formatAddress(headerRow, minCol)
      const sourceEndAddress = formatAddress(maxRow, maxCol)
      const targetSheetName = createUniqueSheetName(
        snapshot.sheets.map((sheet) => sheet.name),
        `${sheetName} Review`,
      )
      const targetEndAddress = formatAddress(maxRow - headerRow, maxCol - minCol)
      return {
        title: 'Create Current Sheet Review Tab',
        summary: `Staged a review-tab change set for ${sheetName} into ${targetSheetName}.`,
        artifact: {
          kind: 'markdown',
          title: 'Current Sheet Review Tab Preview',
          text: [
            '## Current Sheet Review Tab Preview',
            '',
            `Source sheet: ${sheetName}`,
            `Source range: ${sourceStartAddress}:${sourceEndAddress}`,
            `Target sheet: ${targetSheetName}`,
            '',
            "The staged workbook change set creates a new review tab and copies the current sheet's used range into it for review-oriented work.",
          ].join('\n'),
        },
        citations: [
          {
            kind: 'range',
            sheetName,
            startAddress: sourceStartAddress,
            endAddress: sourceEndAddress,
            role: 'source',
          },
          {
            kind: 'range',
            sheetName: targetSheetName,
            startAddress: 'A1',
            endAddress: targetEndAddress,
            role: 'target',
          },
        ],
        steps: [
          {
            stepId: 'inspect-source-sheet',
            label: 'Inspect source sheet',
            summary: `Loaded the used range from ${sheetName}.`,
          },
          {
            stepId: 'stage-review-tab-preview',
            label: 'Stage review tab preview',
            summary: `Prepared the semantic change set that creates ${targetSheetName}.`,
          },
          {
            stepId: 'draft-review-tab-report',
            label: 'Draft review tab report',
            summary: 'Prepared the durable review-tab report for the thread.',
          },
        ],
        commands: [
          {
            kind: 'createSheet',
            name: targetSheetName,
          },
          {
            kind: 'copyRange',
            source: {
              sheetName,
              startAddress: sourceStartAddress,
              endAddress: sourceEndAddress,
            },
            target: {
              sheetName: targetSheetName,
              startAddress: 'A1',
              endAddress: targetEndAddress,
            },
          },
        ],
        goalText: `Create a review tab for ${sheetName}`,
      } satisfies RollupWorkflowExecutionResult
    }
    const columnSummaries: RollupColumnSummary[] = []
    for (let col = minCol; col <= maxCol; col += 1) {
      const headerCell = cellByAddress.get(formatAddress(headerRow, col))
      const headerLabel =
        headerCell && typeof headerCell.value === 'string'
          ? headerCell.value.trim() || formatAddress(0, col).replace(/\d+/gu, '')
          : formatAddress(0, col).replace(/\d+/gu, '')
      const numericValues: number[] = []
      for (let row = dataStartRow; row <= maxRow; row += 1) {
        const cell = cellByAddress.get(formatAddress(row, col))
        if (cell && typeof cell.value === 'number') {
          numericValues.push(cell.value)
        }
      }
      if (numericValues.length === 0) {
        continue
      }
      const sum = numericValues.reduce((total, value) => total + value, 0)
      columnSummaries.push({
        headerLabel,
        count: numericValues.length,
        sum,
        average: sum / numericValues.length,
        min: Math.min(...numericValues),
        max: Math.max(...numericValues),
      })
    }

    const sourceStartAddress = formatAddress(headerRow, minCol)
    const sourceEndAddress = formatAddress(maxRow, maxCol)
    if (columnSummaries.length === 0) {
      return {
        title: 'Create Current Sheet Rollup',
        summary: `No numeric columns were available to roll up on ${sheetName}.`,
        artifact: {
          kind: 'markdown',
          title: 'Current Sheet Rollup Preview',
          text: summarizeRollupMarkdown({
            sourceSheetName: sheetName,
            targetSheetName: `${sheetName} Rollup`,
            sourceStartAddress,
            sourceEndAddress,
            columns: columnSummaries,
          }),
        },
        citations: [
          {
            kind: 'range',
            sheetName,
            startAddress: sourceStartAddress,
            endAddress: sourceEndAddress,
            role: 'source',
          },
        ],
        steps: [
          {
            stepId: 'inspect-source-sheet',
            label: 'Inspect source sheet',
            summary: `Loaded the used range and headers from ${sheetName}.`,
          },
          {
            stepId: 'aggregate-column-stats',
            label: 'Aggregate column stats',
            summary: 'No numeric columns were available to summarize on the current sheet.',
          },
          {
            stepId: 'stage-rollup-preview',
            label: 'Stage rollup change set',
            summary: 'Rollup analysis complete. Numeric columns summarized: 0.',
          },
          {
            stepId: 'draft-rollup-report',
            label: 'Draft rollup report',
            summary: 'Prepared the durable no-op rollup report for the thread.',
          },
        ],
      } satisfies RollupWorkflowExecutionResult
    }

    const targetSheetName = createUniqueSheetName(
      snapshot.sheets.map((sheet) => sheet.name),
      `${sheetName} Rollup`,
    )
    const values: WorkbookAgentWriteCellInput[][] = [
      ['Metric', 'Value', null, null, null, null],
      ['Source Sheet', sheetName, null, null, null, null],
      ['Rows Scanned', maxRow - headerRow, null, null, null, null],
      ['Numeric Columns', columnSummaries.length, null, null, null, null],
      [null, null, null, null, null, null],
      ['Column', 'Count', 'Sum', 'Average', 'Min', 'Max'],
      ...columnSummaries.map((column) => [column.headerLabel, column.count, column.sum, column.average, column.min, column.max]),
    ]
    const targetEndAddress = formatAddress(values.length - 1, values[0]!.length - 1)

    return {
      title: 'Create Current Sheet Rollup',
      summary: `Staged a rollup change set for ${sheetName} into ${targetSheetName}.`,
      artifact: {
        kind: 'markdown',
        title: 'Current Sheet Rollup Preview',
        text: summarizeRollupMarkdown({
          sourceSheetName: sheetName,
          targetSheetName,
          sourceStartAddress,
          sourceEndAddress,
          columns: columnSummaries,
        }),
      },
      citations: [
        {
          kind: 'range',
          sheetName,
          startAddress: sourceStartAddress,
          endAddress: sourceEndAddress,
          role: 'source',
        },
        {
          kind: 'range',
          sheetName: targetSheetName,
          startAddress: 'A1',
          endAddress: targetEndAddress,
          role: 'target',
        },
      ],
      steps: [
        {
          stepId: 'inspect-source-sheet',
          label: 'Inspect source sheet',
          summary: `Loaded the used range and numeric columns from ${sheetName}.`,
        },
        {
          stepId: 'aggregate-column-stats',
          label: 'Aggregate column stats',
          summary: `Computed rollup metrics for ${String(columnSummaries.length)} numeric column${columnSummaries.length === 1 ? '' : 's'}.`,
        },
        {
          stepId: 'stage-rollup-preview',
          label: 'Stage rollup change set',
          summary: `Prepared the semantic change set that creates ${targetSheetName}.`,
        },
        {
          stepId: 'draft-rollup-report',
          label: 'Draft rollup report',
          summary: 'Prepared the durable rollup report for the thread.',
        },
      ],
      commands: [
        {
          kind: 'createSheet',
          name: targetSheetName,
        },
        {
          kind: 'writeRange',
          sheetName: targetSheetName,
          startAddress: 'A1',
          values,
        },
      ],
      goalText: `Create a rollup sheet for ${sheetName}`,
    } satisfies RollupWorkflowExecutionResult
  })
}
