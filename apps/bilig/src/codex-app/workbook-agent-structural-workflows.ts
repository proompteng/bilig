import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkbookAgentCommand } from '@bilig/agent-api'
import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowTemplate,
} from '@bilig/contracts'
import { throwIfWorkflowCancelled } from './workbook-agent-workflow-abort.js'

export type StructuralWorkflowTemplate =
  | 'createSheet'
  | 'renameCurrentSheet'
  | 'hideCurrentRow'
  | 'hideCurrentColumn'
  | 'unhideCurrentRow'
  | 'unhideCurrentColumn'

interface StructuralWorkflowExecutionInput {
  readonly name?: string
}

interface StructuralWorkflowStepPlan {
  readonly stepId: string
  readonly label: string
  readonly runningSummary: string
  readonly pendingSummary: string
}

interface StructuralWorkflowStepResult {
  readonly stepId: string
  readonly label: string
  readonly summary: string
}

export interface StructuralWorkflowTemplateMetadata {
  readonly title: string
  readonly runningSummary: string
  readonly stepPlans: readonly StructuralWorkflowStepPlan[]
}

export interface StructuralWorkflowExecutionResult {
  readonly title: string
  readonly summary: string
  readonly artifact: WorkbookAgentWorkflowArtifact
  readonly citations: readonly WorkbookAgentTimelineCitation[]
  readonly steps: readonly StructuralWorkflowStepResult[]
  readonly commands: readonly WorkbookAgentCommand[]
  readonly goalText: string
}

function requireSheetName(context: WorkbookAgentUiContext | null | undefined): string {
  const sheetName = context?.selection.sheetName
  if (!sheetName) {
    throw new Error('Selection context is required for this structural workflow.')
  }
  return sheetName
}

function requireSelectionAddress(context: WorkbookAgentUiContext | null | undefined): string {
  const address = context?.selection.address
  if (!address) {
    throw new Error('Selection context is required for this structural workflow.')
  }
  return address
}

function requireWorkflowName(workflowInput: StructuralWorkflowExecutionInput | null | undefined): string {
  const name = workflowInput?.name?.trim()
  if (!name) {
    throw new Error('A non-empty name is required for this structural workflow.')
  }
  return name
}

function createMarkdownArtifact(title: string, lines: readonly string[]): WorkbookAgentWorkflowArtifact {
  return {
    kind: 'markdown',
    title,
    text: lines.join('\n'),
  }
}

export function isStructuralWorkflowTemplate(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
): workflowTemplate is StructuralWorkflowTemplate {
  return (
    workflowTemplate === 'createSheet' ||
    workflowTemplate === 'renameCurrentSheet' ||
    workflowTemplate === 'hideCurrentRow' ||
    workflowTemplate === 'hideCurrentColumn' ||
    workflowTemplate === 'unhideCurrentRow' ||
    workflowTemplate === 'unhideCurrentColumn'
  )
}

export function getStructuralWorkflowTemplateMetadata(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  workflowInput?: StructuralWorkflowExecutionInput | null,
): StructuralWorkflowTemplateMetadata | null {
  if (!isStructuralWorkflowTemplate(workflowTemplate)) {
    return null
  }
  switch (workflowTemplate) {
    case 'createSheet': {
      const sheetName = workflowInput?.name?.trim() || 'new sheet'
      return {
        title: 'Create Sheet',
        runningSummary: `Preparing a structural workbook change set to create ${sheetName}.`,
        stepPlans: [
          {
            stepId: 'plan-sheet-create',
            label: 'Plan sheet creation',
            runningSummary: `Preparing the semantic sheet-creation change set for ${sheetName}.`,
            pendingSummary: 'Waiting to plan the semantic sheet-creation change set.',
          },
          {
            stepId: 'stage-structural-preview',
            label: 'Prepare structural change set',
            runningSummary: 'Preparing the structural workbook change set.',
            pendingSummary: 'Waiting to prepare the structural workbook change set.',
          },
        ],
      }
    }
    case 'renameCurrentSheet': {
      const nextName = workflowInput?.name?.trim() || 'renamed sheet'
      return {
        title: 'Rename Current Sheet',
        runningSummary: `Preparing a structural workbook change set to rename the active sheet to ${nextName}.`,
        stepPlans: [
          {
            stepId: 'inspect-current-sheet',
            label: 'Resolve current sheet',
            runningSummary: 'Resolving the active sheet from the current workbook context.',
            pendingSummary: 'Waiting to resolve the active sheet from the current workbook context.',
          },
          {
            stepId: 'stage-sheet-rename-preview',
            label: 'Prepare sheet rename change set',
            runningSummary: `Preparing the semantic change set that renames the active sheet to ${nextName}.`,
            pendingSummary: 'Waiting to prepare the semantic sheet-rename change set.',
          },
        ],
      }
    }
    case 'hideCurrentRow':
      return {
        title: 'Hide Current Row',
        runningSummary: 'Preparing a structural workbook change set to hide the current row.',
        stepPlans: [
          {
            stepId: 'resolve-current-row',
            label: 'Resolve current row',
            runningSummary: 'Resolving the selected row from the current workbook context.',
            pendingSummary: 'Waiting to resolve the selected row from the current workbook context.',
          },
          {
            stepId: 'stage-row-visibility-preview',
            label: 'Prepare row visibility change set',
            runningSummary: 'Staging the semantic change set that hides the current row.',
            pendingSummary: 'Waiting to prepare the semantic row-visibility change set.',
          },
        ],
      }
    case 'hideCurrentColumn':
      return {
        title: 'Hide Current Column',
        runningSummary: 'Preparing a structural workbook change set to hide the current column.',
        stepPlans: [
          {
            stepId: 'resolve-current-column',
            label: 'Resolve current column',
            runningSummary: 'Resolving the selected column from the current workbook context.',
            pendingSummary: 'Waiting to resolve the selected column from the current workbook context.',
          },
          {
            stepId: 'stage-column-visibility-preview',
            label: 'Prepare column visibility change set',
            runningSummary: 'Staging the semantic change set that hides the current column.',
            pendingSummary: 'Waiting to prepare the semantic column-visibility change set.',
          },
        ],
      }
    case 'unhideCurrentRow':
      return {
        title: 'Unhide Current Row',
        runningSummary: 'Preparing a structural workbook change set to unhide the current row.',
        stepPlans: [
          {
            stepId: 'resolve-current-row',
            label: 'Resolve current row',
            runningSummary: 'Resolving the selected row from the current workbook context.',
            pendingSummary: 'Waiting to resolve the selected row from the current workbook context.',
          },
          {
            stepId: 'stage-row-visibility-preview',
            label: 'Prepare row visibility change set',
            runningSummary: 'Staging the semantic change set that unhides the current row.',
            pendingSummary: 'Waiting to prepare the semantic row-visibility change set.',
          },
        ],
      }
    case 'unhideCurrentColumn':
      return {
        title: 'Unhide Current Column',
        runningSummary: 'Preparing a structural workbook change set to unhide the current column.',
        stepPlans: [
          {
            stepId: 'resolve-current-column',
            label: 'Resolve current column',
            runningSummary: 'Resolving the selected column from the current workbook context.',
            pendingSummary: 'Waiting to resolve the selected column from the current workbook context.',
          },
          {
            stepId: 'stage-column-visibility-preview',
            label: 'Prepare column visibility change set',
            runningSummary: 'Staging the semantic change set that unhides the current column.',
            pendingSummary: 'Waiting to prepare the semantic column-visibility change set.',
          },
        ],
      }
    default:
      return null
  }
}

export function executeStructuralWorkflow(input: {
  workflowTemplate: WorkbookAgentWorkflowTemplate
  context?: WorkbookAgentUiContext | null
  workflowInput?: StructuralWorkflowExecutionInput | null
  signal?: AbortSignal
}): StructuralWorkflowExecutionResult | null {
  if (!isStructuralWorkflowTemplate(input.workflowTemplate)) {
    return null
  }
  throwIfWorkflowCancelled(input.signal)
  switch (input.workflowTemplate) {
    case 'createSheet': {
      const sheetName = requireWorkflowName(input.workflowInput)
      return {
        title: 'Create Sheet',
        summary: `Prepared a structural workbook change set to create ${sheetName}.`,
        artifact: createMarkdownArtifact('Create Sheet Preview', [
          '## Create Sheet Preview',
          '',
          `- Create a new sheet named \`${sheetName}\`.`,
          '- Workbook panel review item is ready whenever the session policy routes this change set to review.',
        ]),
        citations: [],
        steps: [
          {
            stepId: 'plan-sheet-create',
            label: 'Plan sheet creation',
            summary: `Prepared the semantic sheet-creation command for ${sheetName}.`,
          },
          {
            stepId: 'stage-structural-preview',
            label: 'Prepare structural change set',
            summary: 'Prepared the structural workbook change set.',
          },
        ],
        commands: [{ kind: 'createSheet', name: sheetName }],
        goalText: `Create a new sheet named ${sheetName}`,
      }
    }
    case 'renameCurrentSheet': {
      const currentName = requireSheetName(input.context)
      const nextName = requireWorkflowName(input.workflowInput)
      if (currentName === nextName) {
        throw new Error('The new sheet name must be different from the current sheet name.')
      }
      return {
        title: 'Rename Current Sheet',
        summary: `Prepared a structural workbook change set to rename ${currentName} to ${nextName}.`,
        artifact: createMarkdownArtifact('Rename Sheet Preview', [
          '## Rename Sheet Preview',
          '',
          `- Rename the active sheet from \`${currentName}\` to \`${nextName}\`.`,
          '- Workbook panel review item is ready whenever the session policy routes this change set to review.',
        ]),
        citations: [
          {
            kind: 'range',
            sheetName: currentName,
            startAddress: input.context?.selection.address ?? 'A1',
            endAddress: input.context?.selection.address ?? 'A1',
            role: 'source',
          },
        ],
        steps: [
          {
            stepId: 'inspect-current-sheet',
            label: 'Resolve current sheet',
            summary: `Resolved the active sheet as ${currentName}.`,
          },
          {
            stepId: 'stage-sheet-rename-preview',
            label: 'Prepare sheet rename change set',
            summary: `Staged the semantic change set that renames ${currentName} to ${nextName}.`,
          },
        ],
        commands: [{ kind: 'renameSheet', currentName, nextName }],
        goalText: `Rename sheet ${currentName} to ${nextName}`,
      }
    }
    case 'hideCurrentRow': {
      const sheetName = requireSheetName(input.context)
      const address = requireSelectionAddress(input.context)
      const { row } = parseCellAddress(address)
      return {
        title: 'Hide Current Row',
        summary: `Prepared a structural workbook change set to hide row ${String(row)} on ${sheetName}.`,
        artifact: createMarkdownArtifact('Hide Current Row Preview', [
          '## Hide Current Row Preview',
          '',
          `- Hide row ${String(row)} on \`${sheetName}\`.`,
          '- Workbook panel review item is ready whenever the session policy routes this change set to review.',
        ]),
        citations: [
          {
            kind: 'range',
            sheetName,
            startAddress: formatAddress(row, 0),
            endAddress: formatAddress(row, 0),
            role: 'target',
          },
        ],
        steps: [
          {
            stepId: 'resolve-current-row',
            label: 'Resolve current row',
            summary: `Resolved the selected row as ${String(row)} on ${sheetName}.`,
          },
          {
            stepId: 'stage-row-visibility-preview',
            label: 'Prepare row visibility change set',
            summary: `Staged the semantic change set that hides row ${String(row)} on ${sheetName}.`,
          },
        ],
        commands: [
          {
            kind: 'updateRowMetadata',
            sheetName,
            startRow: row,
            count: 1,
            hidden: true,
          },
        ],
        goalText: `Hide row ${String(row)} on ${sheetName}`,
      }
    }
    case 'hideCurrentColumn': {
      const sheetName = requireSheetName(input.context)
      const address = requireSelectionAddress(input.context)
      const { col } = parseCellAddress(address)
      const columnLabel = formatAddress(0, col).replace(/\d+/gu, '')
      return {
        title: 'Hide Current Column',
        summary: `Prepared a structural workbook change set to hide column ${columnLabel} on ${sheetName}.`,
        artifact: createMarkdownArtifact('Hide Current Column Preview', [
          '## Hide Current Column Preview',
          '',
          `- Hide column ${columnLabel} on \`${sheetName}\`.`,
          '- Workbook panel review item is ready whenever the session policy routes this change set to review.',
        ]),
        citations: [
          {
            kind: 'range',
            sheetName,
            startAddress: formatAddress(0, col),
            endAddress: formatAddress(0, col),
            role: 'target',
          },
        ],
        steps: [
          {
            stepId: 'resolve-current-column',
            label: 'Resolve current column',
            summary: `Resolved the selected column as ${columnLabel} on ${sheetName}.`,
          },
          {
            stepId: 'stage-column-visibility-preview',
            label: 'Prepare column visibility change set',
            summary: `Staged the semantic change set that hides column ${columnLabel} on ${sheetName}.`,
          },
        ],
        commands: [
          {
            kind: 'updateColumnMetadata',
            sheetName,
            startCol: col,
            count: 1,
            hidden: true,
          },
        ],
        goalText: `Hide column ${columnLabel} on ${sheetName}`,
      }
    }
    case 'unhideCurrentRow': {
      const sheetName = requireSheetName(input.context)
      const address = requireSelectionAddress(input.context)
      const { row } = parseCellAddress(address)
      return {
        title: 'Unhide Current Row',
        summary: `Prepared a structural workbook change set to unhide row ${String(row)} on ${sheetName}.`,
        artifact: createMarkdownArtifact('Unhide Current Row Preview', [
          '## Unhide Current Row Preview',
          '',
          `- Unhide row ${String(row)} on \`${sheetName}\`.`,
          '- Workbook panel review item is ready whenever the session policy routes this change set to review.',
        ]),
        citations: [
          {
            kind: 'range',
            sheetName,
            startAddress: formatAddress(row, 0),
            endAddress: formatAddress(row, 0),
            role: 'target',
          },
        ],
        steps: [
          {
            stepId: 'resolve-current-row',
            label: 'Resolve current row',
            summary: `Resolved the selected row as ${String(row)} on ${sheetName}.`,
          },
          {
            stepId: 'stage-row-visibility-preview',
            label: 'Prepare row visibility change set',
            summary: `Staged the semantic change set that unhides row ${String(row)} on ${sheetName}.`,
          },
        ],
        commands: [
          {
            kind: 'updateRowMetadata',
            sheetName,
            startRow: row,
            count: 1,
            hidden: false,
          },
        ],
        goalText: `Unhide row ${String(row)} on ${sheetName}`,
      }
    }
    case 'unhideCurrentColumn': {
      const sheetName = requireSheetName(input.context)
      const address = requireSelectionAddress(input.context)
      const { col } = parseCellAddress(address)
      const columnLabel = formatAddress(0, col).replace(/\d+/gu, '')
      return {
        title: 'Unhide Current Column',
        summary: `Prepared a structural workbook change set to unhide column ${columnLabel} on ${sheetName}.`,
        artifact: createMarkdownArtifact('Unhide Current Column Preview', [
          '## Unhide Current Column Preview',
          '',
          `- Unhide column ${columnLabel} on \`${sheetName}\`.`,
          '- Workbook panel review item is ready whenever the session policy routes this change set to review.',
        ]),
        citations: [
          {
            kind: 'range',
            sheetName,
            startAddress: formatAddress(0, col),
            endAddress: formatAddress(0, col),
            role: 'target',
          },
        ],
        steps: [
          {
            stepId: 'resolve-current-column',
            label: 'Resolve current column',
            summary: `Resolved the selected column as ${columnLabel} on ${sheetName}.`,
          },
          {
            stepId: 'stage-column-visibility-preview',
            label: 'Prepare column visibility change set',
            summary: `Staged the semantic change set that unhides column ${columnLabel} on ${sheetName}.`,
          },
        ],
        commands: [
          {
            kind: 'updateColumnMetadata',
            sheetName,
            startCol: col,
            count: 1,
            hidden: false,
          },
        ],
        goalText: `Unhide column ${columnLabel} on ${sheetName}`,
      }
    }
    default:
      return null
  }
}
