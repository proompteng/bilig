import {
  createWorkbookAgentCommandBundle,
  toWorkbookAgentCommandBundle,
  toWorkbookAgentReviewQueueItem,
  type CodexThread,
  type WorkbookAgentReviewQueueItem,
} from '@bilig/agent-api'
import { createBundleRangeCitations } from './workbook-agent-bundle-state.js'
import { buildEntriesFromThread, createSystemEntry } from './workbook-agent-session-model.js'
import {
  mergeTimelineEntries,
  normalizeExecutionPolicy,
  upsertEntry,
  type WorkbookAgentThreadState,
} from './workbook-agent-service-shared.js'
import type { WorkbookAgentLoadedThreadState } from './workbook-agent-thread-repository.js'

export const LEGACY_PRIVATE_BOOTSTRAP_REVIEW_MESSAGE =
  'Private workbook threads no longer keep queued review items. Replay the request to apply it again.'

function resolveWorkbookAgentBootstrapStatus(thread: CodexThread | null): WorkbookAgentThreadState['live']['status'] {
  if (!thread) {
    return 'failed'
  }
  if (thread.turns.some((turn) => turn.status === 'failed')) {
    return 'failed'
  }
  if (thread.turns.some((turn) => turn.status === 'inProgress')) {
    return 'inProgress'
  }
  return 'idle'
}

function resolveWorkbookAgentBootstrapErrorMessage(thread: CodexThread | null, sessionBootstrapError: unknown): string | null {
  if (thread) {
    return thread.turns.findLast((turn) => turn.error?.message)?.error?.message ?? null
  }
  if (!sessionBootstrapError) {
    return null
  }
  return sessionBootstrapError instanceof Error
    ? sessionBootstrapError.message
    : 'Workbook assistant live session is unavailable. Loaded durable thread history only.'
}

export function createWorkbookAgentBootstrappedSessionState(input: {
  readonly documentId: string
  readonly userId: string
  readonly threadId: string
  readonly requestedScope?: WorkbookAgentThreadState['scope']
  readonly requestedExecutionPolicy?: WorkbookAgentThreadState['executionPolicy'] | null
  readonly requestedContext?: WorkbookAgentThreadState['durable']['context']
  readonly durableThreadSession: WorkbookAgentLoadedThreadState
  readonly liveThread: CodexThread | null
  readonly sessionBootstrapError: unknown
  readonly now: number
}): WorkbookAgentThreadState {
  const durableThreadState = input.durableThreadSession.threadState
  const resolvedScope = durableThreadState?.scope ?? input.requestedScope ?? 'private'
  const resolvedExecutionPolicy = normalizeExecutionPolicy({
    scope: resolvedScope,
    requestedPolicy: input.requestedExecutionPolicy ?? durableThreadState?.executionPolicy ?? null,
  })
  const codexEntries = input.liveThread ? buildEntriesFromThread(input.liveThread) : []

  return {
    documentId: input.documentId,
    userId: input.userId,
    storageActorUserId: durableThreadState?.actorUserId ?? input.userId,
    scope: resolvedScope,
    executionPolicy: resolvedExecutionPolicy,
    threadId: input.threadId,
    durable: {
      context: input.requestedContext ?? durableThreadState?.context ?? null,
      entries: mergeTimelineEntries(codexEntries, durableThreadState?.entries ?? []),
      reviewQueueItems: [...(durableThreadState?.reviewQueueItems ?? [])],
      executionRecords: input.durableThreadSession.executionRecords,
      workflowRuns: input.durableThreadSession.workflowRuns,
    },
    live: {
      activeTurnId: input.liveThread?.turns.findLast((turn) => turn.status === 'inProgress')?.id ?? null,
      status: resolveWorkbookAgentBootstrapStatus(input.liveThread),
      lastError: resolveWorkbookAgentBootstrapErrorMessage(input.liveThread, input.sessionBootstrapError),
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn: new Map(),
      turnContextByTurn: new Map(),
      lastAccessedAt: input.now,
    },
  }
}

export function planWorkbookAgentBootstrapReviewRecovery(input: {
  readonly sessionState: WorkbookAgentThreadState
  readonly rolloutAllowed: boolean
  readonly autoApplyLowRiskEnabled: boolean
}):
  | { readonly kind: 'none' }
  | { readonly kind: 'autoApply'; readonly reviewItem: WorkbookAgentReviewQueueItem }
  | { readonly kind: 'clearLegacy'; readonly reviewItem: WorkbookAgentReviewQueueItem } {
  const reviewItem = input.sessionState.durable.reviewQueueItems[0] ?? null
  if (!reviewItem || input.sessionState.scope !== 'private') {
    return { kind: 'none' }
  }
  if (
    input.rolloutAllowed &&
    (input.sessionState.executionPolicy === 'autoApplyAll' ||
      (input.sessionState.executionPolicy === 'autoApplySafe' &&
        input.autoApplyLowRiskEnabled &&
        (reviewItem.riskClass ?? 'high') === 'low'))
  ) {
    return {
      kind: 'autoApply',
      reviewItem,
    }
  }
  return {
    kind: 'clearLegacy',
    reviewItem,
  }
}

export function rebaseWorkbookAgentBootstrapReviewItem(input: {
  readonly reviewItem: WorkbookAgentReviewQueueItem
  readonly currentRevision: number
  readonly fallbackOwnerUserId: string
}): WorkbookAgentReviewQueueItem {
  const queuedBundle = toWorkbookAgentCommandBundle(input.reviewItem)
  const migratedBundle =
    queuedBundle.baseRevision === input.currentRevision
      ? queuedBundle
      : createWorkbookAgentCommandBundle({
          bundleId: queuedBundle.id,
          documentId: queuedBundle.documentId,
          threadId: queuedBundle.threadId,
          turnId: queuedBundle.turnId,
          goalText: queuedBundle.goalText,
          baseRevision: input.currentRevision,
          context: queuedBundle.context,
          commands: queuedBundle.commands,
          now: queuedBundle.createdAtUnixMs,
          sharedReview: queuedBundle.sharedReview ?? null,
        })
  return toWorkbookAgentReviewQueueItem({
    bundle: migratedBundle,
    reviewMode: input.reviewItem.reviewMode,
    sharedReview:
      input.reviewItem.reviewMode === 'ownerReview'
        ? {
            ownerUserId: input.reviewItem.ownerUserId ?? input.fallbackOwnerUserId,
            status: input.reviewItem.status,
            decidedByUserId: input.reviewItem.decidedByUserId,
            decidedAtUnixMs: input.reviewItem.decidedAtUnixMs,
            recommendations: [...input.reviewItem.recommendations],
          }
        : null,
  })
}

export function clearLegacyPrivateBootstrapReviewItem(input: {
  readonly sessionState: WorkbookAgentThreadState
  readonly reviewItem: WorkbookAgentReviewQueueItem
  readonly now: number
}): void {
  input.sessionState.durable.reviewQueueItems = []
  input.sessionState.live.lastError = LEGACY_PRIVATE_BOOTSTRAP_REVIEW_MESSAGE
  input.sessionState.durable.entries = upsertEntry(
    input.sessionState.durable.entries,
    createSystemEntry(
      `system-private-review-item-cleared:${input.reviewItem.id}:${input.now}`,
      input.reviewItem.turnId,
      `Cleared a legacy private review item that could not be resumed automatically: ${input.reviewItem.summary}`,
      createBundleRangeCitations(toWorkbookAgentCommandBundle(input.reviewItem)),
    ),
  )
}
