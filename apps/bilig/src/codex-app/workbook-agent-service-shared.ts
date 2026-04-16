import type {
  CodexServerNotification,
  WorkbookAgentCommandBundle,
  WorkbookAgentContextRef,
  WorkbookAgentExecutionRecord,
  WorkbookAgentReviewQueueItem,
} from '@bilig/agent-api'
import type {
  WorkbookAgentExecutionPolicy,
  WorkbookAgentThreadSnapshot,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowRun,
} from '@bilig/contracts'
export interface WorkbookAgentThreadDurableState {
  context: WorkbookAgentUiContext | null
  entries: WorkbookAgentTimelineEntry[]
  reviewQueueItems: WorkbookAgentReviewQueueItem[]
  executionRecords: WorkbookAgentExecutionRecord[]
  workflowRuns: WorkbookAgentWorkflowRun[]
}

export interface WorkbookAgentThreadLiveState {
  activeTurnId: string | null
  status: WorkbookAgentThreadSnapshot['status']
  lastError: string | null
  stagedPrivateBundleByTurn: Map<string, WorkbookAgentCommandBundle>
  optimisticUserEntryIdByTurn: Map<string, string>
  promptByTurn: Map<string, string>
  turnActorUserIdByTurn: Map<string, string>
  turnContextByTurn: Map<string, WorkbookAgentUiContext | null>
  lastAccessedAt: number
}

export interface WorkbookAgentThreadState {
  readonly documentId: string
  readonly userId: string
  readonly storageActorUserId: string
  scope: 'private' | 'shared'
  executionPolicy: WorkbookAgentExecutionPolicy
  threadId: string
  durable: WorkbookAgentThreadDurableState
  live: WorkbookAgentThreadLiveState
}

export interface WorkbookAgentWorkflowInput {
  readonly query?: string
  readonly sheetName?: string
  readonly limit?: number
  readonly name?: string
}

export interface QueuedWorkbookAgentWorkflowRun {
  readonly sessionState: WorkbookAgentThreadState
  readonly documentId: string
  readonly runId: string
  readonly workflowTurnId: string
  readonly workflowTemplate: WorkbookAgentWorkflowRun['workflowTemplate']
  readonly workflowInput: WorkbookAgentWorkflowInput
  readonly startedByUserId: string
  readonly runningRun: WorkbookAgentWorkflowRun
}

export function upsertEntry(
  entries: readonly WorkbookAgentTimelineEntry[],
  nextEntry: WorkbookAgentTimelineEntry,
): WorkbookAgentTimelineEntry[] {
  const index = entries.findIndex((entry) => entry.id === nextEntry.id)
  if (index < 0) {
    return [...entries, nextEntry]
  }
  const nextEntries = [...entries]
  nextEntries[index] = nextEntry
  return nextEntries
}

export function removeEntry(entries: readonly WorkbookAgentTimelineEntry[], entryId: string): WorkbookAgentTimelineEntry[] {
  return entries.filter((entry) => entry.id !== entryId)
}

export function upsertWorkflowRun(
  runs: readonly WorkbookAgentWorkflowRun[],
  nextRun: WorkbookAgentWorkflowRun,
): WorkbookAgentWorkflowRun[] {
  const index = runs.findIndex((run) => run.runId === nextRun.runId)
  if (index < 0) {
    return [nextRun, ...runs]
  }
  const nextRuns = [...runs]
  nextRuns[index] = nextRun
  return nextRuns
}

export function mergeTimelineEntries(
  codexEntries: readonly WorkbookAgentTimelineEntry[],
  durableEntries: readonly WorkbookAgentTimelineEntry[],
): WorkbookAgentTimelineEntry[] {
  const merged = [...codexEntries]
  const indexById = new Map(merged.map((entry, index) => [entry.id, index]))
  for (const entry of durableEntries) {
    const existingIndex = indexById.get(entry.id)
    if (existingIndex === undefined) {
      indexById.set(entry.id, merged.length)
      merged.push(entry)
      continue
    }
    merged[existingIndex] = entry
  }
  return merged
}

export function buildSnapshot(sessionState: WorkbookAgentThreadState): WorkbookAgentThreadSnapshot {
  return {
    documentId: sessionState.documentId,
    threadId: sessionState.threadId,
    scope: sessionState.scope,
    executionPolicy: sessionState.executionPolicy,
    status: sessionState.live.status,
    activeTurnId: sessionState.live.activeTurnId,
    lastError: sessionState.live.lastError,
    context: sessionState.durable.context ? structuredClone(sessionState.durable.context) : null,
    entries: sessionState.durable.entries.map((entry) => ({ ...entry })),
    reviewQueueItems: sessionState.durable.reviewQueueItems.map((item) => structuredClone(item)),
    executionRecords: sessionState.durable.executionRecords.map((record) => structuredClone(record)),
    workflowRuns: sessionState.durable.workflowRuns.map((run) => structuredClone(run)),
  }
}

export function normalizeExecutionPolicy(input: {
  scope: 'private' | 'shared'
  requestedPolicy?: WorkbookAgentExecutionPolicy | null
}): WorkbookAgentExecutionPolicy {
  if (input.requestedPolicy) {
    return input.requestedPolicy
  }
  return input.scope === 'shared' ? 'ownerReview' : 'autoApplyAll'
}

export function toContextRef(context: WorkbookAgentUiContext | null): WorkbookAgentContextRef | null {
  return context
    ? {
        selection: {
          sheetName: context.selection.sheetName,
          address: context.selection.address,
          ...(context.selection.range
            ? {
                range: {
                  startAddress: context.selection.range.startAddress,
                  endAddress: context.selection.range.endAddress,
                },
              }
            : {}),
        },
        viewport: { ...context.viewport },
      }
    : null
}

export function cloneUiContext(context: WorkbookAgentUiContext | null): WorkbookAgentUiContext | null {
  return context
    ? {
        selection: {
          sheetName: context.selection.sheetName,
          address: context.selection.address,
          ...(context.selection.range
            ? {
                range: {
                  startAddress: context.selection.range.startAddress,
                  endAddress: context.selection.range.endAddress,
                },
              }
            : {}),
        },
        viewport: {
          ...context.viewport,
        },
      }
    : null
}

export function isMutatingWorkflowTemplate(workflowTemplate: string): boolean {
  return (
    workflowTemplate === 'highlightFormulaIssues' ||
    workflowTemplate === 'repairFormulaIssues' ||
    workflowTemplate === 'highlightCurrentSheetOutliers' ||
    workflowTemplate === 'styleCurrentSheetHeaders' ||
    workflowTemplate === 'normalizeCurrentSheetHeaders' ||
    workflowTemplate === 'normalizeCurrentSheetNumberFormats' ||
    workflowTemplate === 'normalizeCurrentSheetWhitespace' ||
    workflowTemplate === 'fillCurrentSheetFormulasDown' ||
    workflowTemplate === 'createCurrentSheetRollup' ||
    workflowTemplate === 'createSheet' ||
    workflowTemplate === 'renameCurrentSheet' ||
    workflowTemplate === 'hideCurrentRow' ||
    workflowTemplate === 'hideCurrentColumn' ||
    workflowTemplate === 'unhideCurrentRow' ||
    workflowTemplate === 'unhideCurrentColumn'
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readNonEmptyString(value: unknown): string | null {
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : null
}

function extractCodexNotificationErrorMessage(value: unknown): string | null {
  const direct = readNonEmptyString(value)
  if (direct) {
    return direct
  }
  if (!isRecord(value)) {
    return null
  }

  for (const key of ['message', 'detail', 'details', 'reason', 'hint', 'title', 'errorMessage']) {
    const nested = readNonEmptyString(value[key])
    if (nested) {
      return nested
    }
  }

  for (const key of ['error', 'cause', 'data']) {
    const nested = extractCodexNotificationErrorMessage(value[key])
    if (nested) {
      return nested
    }
  }

  const errors = value['errors']
  if (Array.isArray(errors)) {
    for (const entry of errors) {
      const nested = extractCodexNotificationErrorMessage(entry)
      if (nested) {
        return nested
      }
    }
  }

  return null
}

export function normalizeCodexNotificationErrorMessage(notification: CodexServerNotification): string {
  return extractCodexNotificationErrorMessage(notification.params) ?? 'Workbook assistant runtime failed. Retry in a moment.'
}
