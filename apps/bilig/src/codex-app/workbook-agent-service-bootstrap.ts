import {
  createWorkbookAgentCommandBundle,
  toWorkbookAgentCommandBundle,
  toWorkbookAgentReviewQueueItem,
  type CodexThread,
  type WorkbookAgentReviewQueueItem,
} from '@bilig/agent-api'
import type { WorkbookAgentTimelineEntry, WorkbookAgentWorkflowRun } from '@bilig/contracts'
import { createBundleRangeCitations } from './workbook-agent-bundle-state.js'
import { buildEntriesFromThread, createSystemEntry } from './workbook-agent-session-model.js'
import {
  mergeTimelineEntries,
  normalizeExecutionPolicy,
  upsertEntry,
  type WorkbookAgentThreadState,
} from './workbook-agent-service-shared.js'
import type { WorkbookAgentLoadedThreadState } from './workbook-agent-thread-repository.js'
import { failWorkflowSteps } from './workbook-agent-workflows.js'

export const LEGACY_PRIVATE_BOOTSTRAP_REVIEW_MESSAGE =
  'Private workbook threads no longer keep queued review items. Replay the request to apply it again.'
export const STALE_BOOTSTRAP_WORKFLOW_MESSAGE = 'Workflow interrupted because the workbook assistant restarted before it could finish.'

export interface WorkbookAgentBootstrapWorkflowRecovery {
  readonly workflowRuns: WorkbookAgentWorkflowRun[]
  readonly recoveredRuns: WorkbookAgentWorkflowRun[]
  readonly entries: WorkbookAgentTimelineEntry[]
}

function findBootstrapActiveTurn(thread: CodexThread | null): CodexThread['turns'][number] | null {
  if (!thread) {
    return null
  }
  const activeTurn = thread.turns.findLast((turn) => turn.status === 'inProgress') ?? null
  if (!activeTurn) {
    return null
  }
  if (thread.status === undefined) {
    return activeTurn
  }
  return thread.status.type === 'active' ? activeTurn : null
}

function findStaleBootstrapInProgressTurns(thread: CodexThread | null): CodexThread['turns'] {
  if (!thread || thread.status === undefined || thread.status.type === 'active') {
    return []
  }
  return thread.turns.filter((turn) => turn.status === 'inProgress')
}

function resolveWorkbookAgentBootstrapStatus(thread: CodexThread | null): WorkbookAgentThreadState['live']['status'] {
  if (!thread) {
    return 'failed'
  }
  if (thread.status?.type === 'systemError' || thread.status?.type === 'notLoaded') {
    return 'failed'
  }
  if (thread.turns.some((turn) => turn.status === 'failed')) {
    return 'failed'
  }
  if (findBootstrapActiveTurn(thread)) {
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
  const workflowRecovery = recoverStaleBootstrapWorkflowRuns({
    workflowRuns: input.durableThreadSession.workflowRuns,
    now: input.now,
  })
  const staleInProgressTurns = findStaleBootstrapInProgressTurns(input.liveThread)
  const resolvedScope = durableThreadState?.scope ?? input.requestedScope ?? 'private'
  const resolvedExecutionPolicy = normalizeExecutionPolicy({
    scope: resolvedScope,
    requestedPolicy: input.requestedExecutionPolicy ?? durableThreadState?.executionPolicy ?? null,
  })
  const codexEntries = input.liveThread ? buildEntriesFromThread(input.liveThread) : []
  const staleTurnEntries = staleInProgressTurns.map((turn) =>
    createSystemEntry(
      `system-turn-bootstrap-recovered:${turn.id}:${String(input.now)}`,
      turn.id,
      `Cleared stale in-progress turn ${turn.id} after Codex resumed the thread as ${input.liveThread?.status?.type}.`,
    ),
  )

  return {
    documentId: input.documentId,
    userId: input.userId,
    storageActorUserId: durableThreadState?.actorUserId ?? input.userId,
    scope: resolvedScope,
    executionPolicy: resolvedExecutionPolicy,
    threadId: input.threadId,
    durable: {
      context: input.requestedContext ?? durableThreadState?.context ?? null,
      entries: mergeTimelineEntries([...codexEntries, ...staleTurnEntries, ...workflowRecovery.entries], durableThreadState?.entries ?? []),
      reviewQueueItems: [...(durableThreadState?.reviewQueueItems ?? [])],
      executionRecords: input.durableThreadSession.executionRecords,
      workflowRuns: workflowRecovery.workflowRuns,
    },
    live: {
      activeTurnId: findBootstrapActiveTurn(input.liveThread)?.id ?? null,
      status: resolveWorkbookAgentBootstrapStatus(input.liveThread),
      lastError: resolveWorkbookAgentBootstrapErrorMessage(input.liveThread, input.sessionBootstrapError),
      authorizedUserIds: new Set(
        [input.userId, durableThreadState?.actorUserId].filter((userId): userId is string => typeof userId === 'string'),
      ),
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn: new Map(),
      turnContextByTurn: new Map(),
      lastAccessedAt: input.now,
    },
  }
}

function recoverStaleBootstrapWorkflowRuns(input: {
  readonly workflowRuns: readonly WorkbookAgentWorkflowRun[]
  readonly now: number
}): WorkbookAgentBootstrapWorkflowRecovery {
  const recoveredRuns: WorkbookAgentWorkflowRun[] = []
  const entries: WorkbookAgentTimelineEntry[] = []
  const workflowRuns = input.workflowRuns.map((run) => {
    if (run.status !== 'running') {
      return run
    }
    const recoveredRun: WorkbookAgentWorkflowRun = {
      ...run,
      status: 'failed',
      summary: `Workflow interrupted: ${run.title}`,
      updatedAtUnixMs: input.now,
      completedAtUnixMs: input.now,
      errorMessage: STALE_BOOTSTRAP_WORKFLOW_MESSAGE,
      steps: failWorkflowSteps(run.workflowTemplate, run.steps, STALE_BOOTSTRAP_WORKFLOW_MESSAGE, input.now),
      artifact: null,
    }
    recoveredRuns.push(recoveredRun)
    entries.push(
      createSystemEntry(
        `system-workflow-bootstrap-recovered:${run.runId}:${String(input.now)}`,
        null,
        `Marked stale running workflow as failed after assistant restart: ${run.title}`,
      ),
    )
    return recoveredRun
  })
  return {
    workflowRuns,
    recoveredRuns,
    entries,
  }
}

export function findRecoveredStaleBootstrapWorkflowRuns(input: {
  readonly previousWorkflowRuns: readonly WorkbookAgentWorkflowRun[]
  readonly nextWorkflowRuns: readonly WorkbookAgentWorkflowRun[]
}): WorkbookAgentWorkflowRun[] {
  const previousRunningIds = new Set(input.previousWorkflowRuns.filter((run) => run.status === 'running').map((run) => run.runId))
  return input.nextWorkflowRuns.filter(
    (run) => previousRunningIds.has(run.runId) && run.status === 'failed' && run.errorMessage === STALE_BOOTSTRAP_WORKFLOW_MESSAGE,
  )
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
  const baseRevisionChanged = queuedBundle.baseRevision !== input.currentRevision
  const migratedBundle = !baseRevisionChanged
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
            status: baseRevisionChanged ? 'pending' : input.reviewItem.status,
            decidedByUserId: baseRevisionChanged ? null : input.reviewItem.decidedByUserId,
            decidedAtUnixMs: baseRevisionChanged ? null : input.reviewItem.decidedAtUnixMs,
            recommendations: baseRevisionChanged ? [] : [...input.reviewItem.recommendations],
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
