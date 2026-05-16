import {
  isWorkbookAgentCommand,
  isWorkbookAgentReviewQueueItem,
  normalizeWorkbookAgentToolName,
  type WorkbookAgentContextRef,
  type WorkbookAgentReviewQueueItem,
} from '@bilig/agent-api'
import type {
  WorkbookAgentExecutionPolicy,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineCitation,
  WorkbookAgentTimelineEntry,
  WorkbookAgentToolStatus,
  WorkbookAgentUiContext,
} from '@bilig/contracts'
import type { QueryResultRow } from './store.js'
import { parseNullableInteger } from './store-support.js'

export type WorkbookChatThreadScope = 'private' | 'shared'

export interface WorkbookChatThreadRow extends QueryResultRow {
  readonly workbookId?: unknown
  readonly threadId?: unknown
  readonly actorUserId?: unknown
  readonly scope?: unknown
  readonly executionPolicy?: unknown
  readonly contextJson?: unknown
  readonly updatedAtUnixMs?: unknown
}

export interface WorkbookChatItemRow extends QueryResultRow {
  readonly entryId?: unknown
  readonly turnId?: unknown
  readonly kind?: unknown
  readonly text?: unknown
  readonly phase?: unknown
  readonly toolName?: unknown
  readonly toolStatus?: unknown
  readonly argumentsText?: unknown
  readonly outputText?: unknown
  readonly success?: unknown
  readonly citationsJson?: unknown
  readonly sortOrder?: unknown
}

export interface WorkbookChatToolCallRow extends QueryResultRow {
  readonly entryId?: unknown
  readonly turnId?: unknown
  readonly toolName?: unknown
  readonly toolStatus?: unknown
  readonly argumentsText?: unknown
  readonly outputText?: unknown
  readonly success?: unknown
  readonly sortOrder?: unknown
}

export interface WorkbookReviewQueueItemRow extends QueryResultRow {
  readonly reviewItemId?: unknown
  readonly workbookId?: unknown
  readonly threadId?: unknown
  readonly actorUserId?: unknown
  readonly turnId?: unknown
  readonly goalText?: unknown
  readonly summary?: unknown
  readonly scope?: unknown
  readonly riskClass?: unknown
  readonly reviewMode?: unknown
  readonly ownerUserId?: unknown
  readonly status?: unknown
  readonly decidedByUserId?: unknown
  readonly decidedAtUnixMs?: unknown
  readonly baseRevision?: unknown
  readonly createdAtUnixMs?: unknown
  readonly contextJson?: unknown
  readonly commandsJson?: unknown
  readonly affectedRangesJson?: unknown
  readonly estimatedAffectedCells?: unknown
  readonly recommendationsJson?: unknown
}

export interface WorkbookChatThreadSummaryRow extends QueryResultRow {
  readonly threadId?: unknown
  readonly scope?: unknown
  readonly ownerUserId?: unknown
  readonly updatedAtUnixMs?: unknown
  readonly entryCount?: unknown
  readonly reviewQueueItemCount?: unknown
  readonly latestEntryText?: unknown
}

export function isExecutionPolicy(value: unknown): value is WorkbookAgentExecutionPolicy {
  return value === 'autoApplySafe' || value === 'autoApplyAll' || value === 'ownerReview'
}

export function parseNumericValue(value: unknown): number | null {
  return parseNullableInteger(value)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function normalizeTimelineToolName(toolName: string | null): string | null {
  if (typeof toolName !== 'string') {
    return null
  }
  return normalizeWorkbookAgentToolName(toolName)
}

export function isWorkbookAgentUiContext(value: unknown): value is WorkbookAgentUiContext {
  return (
    isRecord(value) &&
    isRecord(value['selection']) &&
    typeof value['selection']['sheetName'] === 'string' &&
    typeof value['selection']['address'] === 'string' &&
    (value['selection']['range'] === undefined ||
      (isRecord(value['selection']['range']) &&
        typeof value['selection']['range']['startAddress'] === 'string' &&
        typeof value['selection']['range']['endAddress'] === 'string')) &&
    isRecord(value['viewport']) &&
    typeof value['viewport']['rowStart'] === 'number' &&
    typeof value['viewport']['rowEnd'] === 'number' &&
    typeof value['viewport']['colStart'] === 'number' &&
    typeof value['viewport']['colEnd'] === 'number'
  )
}

function toBundleContextRef(context: WorkbookAgentUiContext | null): WorkbookAgentContextRef | null {
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

function isToolStatus(value: unknown): value is WorkbookAgentToolStatus {
  return value === 'inProgress' || value === 'completed' || value === 'failed' || value === null
}

function isTimelineKind(value: unknown): value is WorkbookAgentTimelineEntry['kind'] {
  return value === 'user' || value === 'assistant' || value === 'plan' || value === 'reasoning' || value === 'tool' || value === 'system'
}

function isTimelineCitation(value: unknown): value is WorkbookAgentTimelineCitation {
  return (
    (isRecord(value) &&
      value['kind'] === 'range' &&
      typeof value['sheetName'] === 'string' &&
      typeof value['startAddress'] === 'string' &&
      typeof value['endAddress'] === 'string' &&
      (value['role'] === 'target' || value['role'] === 'source')) ||
    (isRecord(value) && value['kind'] === 'revision' && typeof value['revision'] === 'number')
  )
}

export function normalizeTimelineEntry(row: WorkbookChatItemRow): WorkbookAgentTimelineEntry | null {
  if (
    typeof row.entryId !== 'string' ||
    !isTimelineKind(row.kind) ||
    (row.turnId !== null && row.turnId !== undefined && typeof row.turnId !== 'string') ||
    (row.text !== null && row.text !== undefined && typeof row.text !== 'string') ||
    (row.phase !== null && row.phase !== undefined && typeof row.phase !== 'string') ||
    (row.toolName !== null && row.toolName !== undefined && typeof row.toolName !== 'string') ||
    !isToolStatus(row.toolStatus ?? null) ||
    (row.argumentsText !== null && row.argumentsText !== undefined && typeof row.argumentsText !== 'string') ||
    (row.outputText !== null && row.outputText !== undefined && typeof row.outputText !== 'string') ||
    (row.success !== null && row.success !== undefined && typeof row.success !== 'boolean') ||
    (row.citationsJson !== null &&
      row.citationsJson !== undefined &&
      (!Array.isArray(row.citationsJson) || !row.citationsJson.every((entry) => isTimelineCitation(entry))))
  ) {
    return null
  }
  const toolStatus: WorkbookAgentToolStatus =
    row.toolStatus === 'inProgress' || row.toolStatus === 'completed' || row.toolStatus === 'failed' ? row.toolStatus : null
  return {
    id: row.entryId,
    kind: row.kind,
    turnId: typeof row.turnId === 'string' ? row.turnId : null,
    text: typeof row.text === 'string' ? row.text : null,
    phase: typeof row.phase === 'string' ? row.phase : null,
    toolName: normalizeTimelineToolName(typeof row.toolName === 'string' ? row.toolName : null),
    toolStatus,
    argumentsText: typeof row.argumentsText === 'string' ? row.argumentsText : null,
    outputText: typeof row.outputText === 'string' ? row.outputText : null,
    success: typeof row.success === 'boolean' ? row.success : null,
    citations: Array.isArray(row.citationsJson) ? [...row.citationsJson] : [],
  }
}

export function normalizeToolCallRow(row: WorkbookChatToolCallRow): {
  readonly entryId: string
  readonly turnId: string | null
  readonly toolName: string | null
  readonly toolStatus: WorkbookAgentToolStatus
  readonly argumentsText: string | null
  readonly outputText: string | null
  readonly success: boolean | null
} | null {
  if (
    typeof row.entryId !== 'string' ||
    (row.turnId !== null && row.turnId !== undefined && typeof row.turnId !== 'string') ||
    (row.toolName !== null && row.toolName !== undefined && typeof row.toolName !== 'string') ||
    !isToolStatus(row.toolStatus ?? null) ||
    (row.argumentsText !== null && row.argumentsText !== undefined && typeof row.argumentsText !== 'string') ||
    (row.outputText !== null && row.outputText !== undefined && typeof row.outputText !== 'string') ||
    (row.success !== null && row.success !== undefined && typeof row.success !== 'boolean')
  ) {
    return null
  }
  const toolStatus: WorkbookAgentToolStatus =
    row.toolStatus === 'inProgress' || row.toolStatus === 'completed' || row.toolStatus === 'failed' ? row.toolStatus : null
  return {
    entryId: row.entryId,
    turnId: typeof row.turnId === 'string' ? row.turnId : null,
    toolName: normalizeTimelineToolName(typeof row.toolName === 'string' ? row.toolName : null),
    toolStatus,
    argumentsText: typeof row.argumentsText === 'string' ? row.argumentsText : null,
    outputText: typeof row.outputText === 'string' ? row.outputText : null,
    success: typeof row.success === 'boolean' ? row.success : null,
  }
}

export function hasToolCallState(entry: WorkbookAgentTimelineEntry): boolean {
  return (
    entry.kind === 'tool' ||
    entry.toolName !== null ||
    entry.toolStatus !== null ||
    entry.argumentsText !== null ||
    entry.outputText !== null ||
    entry.success !== null
  )
}

export function normalizeReviewQueueItem(row: WorkbookReviewQueueItemRow): WorkbookAgentReviewQueueItem | null {
  const baseRevision = parseNumericValue(row.baseRevision)
  const createdAtUnixMs = parseNumericValue(row.createdAtUnixMs)
  const decidedAtUnixMs = row.decidedAtUnixMs === null || row.decidedAtUnixMs === undefined ? null : parseNumericValue(row.decidedAtUnixMs)
  const estimatedAffectedCells =
    row.estimatedAffectedCells === null || row.estimatedAffectedCells === undefined ? null : parseNumericValue(row.estimatedAffectedCells)
  if (
    typeof row.reviewItemId !== 'string' ||
    typeof row.workbookId !== 'string' ||
    typeof row.threadId !== 'string' ||
    typeof row.turnId !== 'string' ||
    typeof row.goalText !== 'string' ||
    typeof row.summary !== 'string' ||
    (row.scope !== 'selection' && row.scope !== 'sheet' && row.scope !== 'workbook') ||
    (row.riskClass !== 'low' && row.riskClass !== 'medium' && row.riskClass !== 'high') ||
    (row.reviewMode !== 'manual' && row.reviewMode !== 'ownerReview') ||
    (row.ownerUserId !== null && row.ownerUserId !== undefined && typeof row.ownerUserId !== 'string') ||
    (row.status !== 'pending' && row.status !== 'approved' && row.status !== 'rejected') ||
    (row.decidedByUserId !== null && row.decidedByUserId !== undefined && typeof row.decidedByUserId !== 'string') ||
    baseRevision === null ||
    createdAtUnixMs === null ||
    !Array.isArray(row.commandsJson) ||
    !row.commandsJson.every((entry) => isWorkbookAgentCommand(entry)) ||
    !Array.isArray(row.affectedRangesJson) ||
    !Array.isArray(row.recommendationsJson)
  ) {
    return null
  }
  const reviewItem = {
    id: row.reviewItemId,
    documentId: row.workbookId,
    threadId: row.threadId,
    turnId: row.turnId,
    goalText: row.goalText,
    summary: row.summary,
    scope: row.scope,
    riskClass: row.riskClass,
    reviewMode: row.reviewMode,
    ownerUserId: typeof row.ownerUserId === 'string' ? row.ownerUserId : null,
    status: row.status,
    decidedByUserId: typeof row.decidedByUserId === 'string' ? row.decidedByUserId : null,
    decidedAtUnixMs,
    recommendations: [...row.recommendationsJson],
    baseRevision,
    createdAtUnixMs,
    context: toBundleContextRef(isWorkbookAgentUiContext(row.contextJson) ? row.contextJson : null),
    commands: [...row.commandsJson],
    affectedRanges: [...row.affectedRangesJson],
    estimatedAffectedCells,
  } satisfies WorkbookAgentReviewQueueItem
  return isWorkbookAgentReviewQueueItem(reviewItem) ? reviewItem : null
}

export function normalizeThreadSummary(row: WorkbookChatThreadSummaryRow): WorkbookAgentThreadSummary | null {
  const updatedAtUnixMs = parseNumericValue(row.updatedAtUnixMs)
  const entryCount = parseNumericValue(row.entryCount)
  const reviewQueueItemCount = parseNumericValue(row.reviewQueueItemCount)
  if (
    typeof row.threadId !== 'string' ||
    (row.scope !== 'private' && row.scope !== 'shared') ||
    typeof row.ownerUserId !== 'string' ||
    updatedAtUnixMs === null ||
    entryCount === null ||
    reviewQueueItemCount === null ||
    (row.latestEntryText !== null && row.latestEntryText !== undefined && typeof row.latestEntryText !== 'string')
  ) {
    return null
  }
  return {
    threadId: row.threadId,
    scope: row.scope,
    ownerUserId: row.ownerUserId,
    updatedAtUnixMs,
    entryCount,
    reviewQueueItemCount,
    latestEntryText: typeof row.latestEntryText === 'string' ? row.latestEntryText : null,
  }
}

export function defaultExecutionPolicyForScope(scope: WorkbookChatThreadScope): WorkbookAgentExecutionPolicy {
  return scope === 'shared' ? 'ownerReview' : 'autoApplyAll'
}
