import type { WorkbookAgentAppliedBy, WorkbookAgentCommandBundle, WorkbookAgentExecutionRecord } from '@bilig/agent-api'
import { createWorkbookAgentCommandBundle, decodeWorkbookAgentPreviewSummary, toWorkbookAgentCommandBundle } from '@bilig/agent-api'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import { createBundleRangeCitations } from './workbook-agent-bundle-state.js'
import {
  createWorkbookAgentReviewQueueItem,
  getCurrentWorkbookAgentReviewItem,
  replaceCurrentWorkbookAgentReviewItem,
  requireWorkbookAgentReviewItem,
} from './workbook-agent-review-transitions.js'
import { createSystemEntry } from './workbook-agent-session-model.js'
import { toContextRef, upsertEntry, type WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export interface WorkbookAgentReviewActionContext {
  readonly now: () => number
  readonly getWorkbookHeadRevision: (documentId: string) => Promise<number>
  readonly applyCommandBundleForSessionState: (input: {
    sessionState: WorkbookAgentThreadState
    commandBundle: WorkbookAgentCommandBundle
    actorUserId: string
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null | undefined
  }) => Promise<WorkbookAgentExecutionRecord>
  readonly shouldApplyToolBundleImmediately: (sessionState: WorkbookAgentThreadState, bundle: WorkbookAgentCommandBundle) => boolean
  readonly persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>
  readonly emitSnapshot: (threadId: string) => void
  readonly touchSession: (sessionState: WorkbookAgentThreadState) => void
}

export async function applyWorkbookAgentReviewItem(input: {
  readonly context: WorkbookAgentReviewActionContext
  readonly sessionState: WorkbookAgentThreadState
  readonly reviewItemId: string
  readonly actorUserId: string
  readonly appliedBy: WorkbookAgentAppliedBy
  readonly commandIndexes?: readonly number[] | null | undefined
  readonly preview: unknown
}): Promise<void> {
  const reviewItem = requireWorkbookAgentReviewItem({
    reviewItem: getCurrentWorkbookAgentReviewItem(input.sessionState),
    reviewItemId: input.reviewItemId,
    notFoundMessage: 'Workbook agent change set was not found.',
  })
  const callerPreview = decodeWorkbookAgentPreviewSummary(input.preview)
  if (!callerPreview) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_PREVIEW_REQUIRED',
      message: 'Workbook preview details are required before applying this change set.',
      statusCode: 400,
      retryable: false,
    })
  }
  const commandBundle = toWorkbookAgentCommandBundle(reviewItem)
  await input.context.applyCommandBundleForSessionState({
    sessionState: input.sessionState,
    commandBundle,
    actorUserId: input.actorUserId,
    appliedBy: input.appliedBy,
    commandIndexes: input.commandIndexes,
  })
  await input.context.persistSessionState(input.sessionState)
  input.context.emitSnapshot(input.sessionState.threadId)
}

export async function replayWorkbookAgentExecutionRecord(input: {
  readonly context: WorkbookAgentReviewActionContext
  readonly sessionState: WorkbookAgentThreadState
  readonly documentId: string
  readonly recordId: string
  readonly actorUserId: string
}): Promise<void> {
  const record = input.sessionState.durable.executionRecords.find((entry) => entry.id === input.recordId)
  if (!record) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_RUN_NOT_FOUND',
      message: 'Workbook agent execution record not found',
      statusCode: 404,
      retryable: false,
    })
  }
  const baseRevision = await input.context.getWorkbookHeadRevision(input.documentId)
  const replayedBundle = createWorkbookAgentCommandBundle({
    documentId: input.documentId,
    threadId: input.sessionState.threadId,
    turnId: `replay:${record.id}:${String(input.context.now())}`,
    goalText: record.goalText,
    baseRevision,
    context: toContextRef(input.sessionState.durable.context) ?? record.context,
    commands: record.commands,
    now: input.context.now(),
  })
  if (input.context.shouldApplyToolBundleImmediately(input.sessionState, replayedBundle)) {
    await input.context.applyCommandBundleForSessionState({
      sessionState: input.sessionState,
      commandBundle: replayedBundle,
      actorUserId: input.actorUserId,
      appliedBy: 'auto',
    })
    await input.context.persistSessionState(input.sessionState)
    input.context.emitSnapshot(input.sessionState.threadId)
    return
  }
  if (input.sessionState.scope === 'private') {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_PRIVATE_EXECUTION_BLOCKED',
      message:
        'Private workbook threads execute replayed changes directly and do not queue review items under the current execution policy.',
      statusCode: 409,
      retryable: false,
    })
  }
  replaceCurrentWorkbookAgentReviewItem(
    input.sessionState,
    createWorkbookAgentReviewQueueItem({ sessionState: input.sessionState, bundle: replayedBundle }),
  )
  input.sessionState.durable.entries = upsertEntry(
    input.sessionState.durable.entries,
    createSystemEntry(
      `system-replay:${record.id}:${String(input.context.now())}`,
      replayedBundle.turnId,
      `Prepared workbook review item from a prior execution: ${replayedBundle.summary}`,
      createBundleRangeCitations(replayedBundle),
    ),
  )
  input.context.touchSession(input.sessionState)
  await input.context.persistSessionState(input.sessionState)
  input.context.emitSnapshot(input.sessionState.threadId)
}
