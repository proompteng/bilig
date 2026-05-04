import type { WorkbookAgentReviewQueueItem } from '@bilig/agent-api'
import type { WorkbookAgentStreamEvent } from '@bilig/contracts'
import type { CodexAppServerClientPoolStats } from './codex-app-server-pool.js'
import type { WorkbookAgentFeatureFlags } from './workbook-agent-feature-flags.js'
import { chooseWorkbookAgentEvictionCandidates } from './workbook-agent-service-session-policy.js'
import { buildSnapshot, type WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export interface WorkbookAgentObservabilityCounters {
  readonly turnBackpressureCount: number
  readonly workflowStartedCount: number
  readonly workflowCompletedCount: number
  readonly workflowFailedCount: number
  readonly workflowCancelledCount: number
  readonly sharedReviewApprovedCount: number
  readonly sharedReviewRejectedCount: number
  readonly sharedRecommendationApprovedCount: number
  readonly sharedRecommendationRejectedCount: number
}

export type WorkbookAgentObservabilityCounterName = keyof WorkbookAgentObservabilityCounters

export interface WorkbookAgentObservabilitySnapshot {
  readonly enabled: boolean
  readonly generatedAtUnixMs: number
  readonly featureFlags: {
    readonly sharedThreadsEnabled: boolean
    readonly workflowRunnerEnabled: boolean
    readonly autoApplyLowRiskEnabled: boolean
    readonly formulaWorkflowFamilyEnabled: boolean
    readonly formattingWorkflowFamilyEnabled: boolean
    readonly importWorkflowFamilyEnabled: boolean
    readonly rollupWorkflowFamilyEnabled: boolean
    readonly structuralWorkflowFamilyEnabled: boolean
    readonly allowlistedUserCount: number
    readonly allowlistedDocumentCount: number
  }
  readonly sessions: {
    readonly sessionCount: number
    readonly subscriberThreadCount: number
    readonly subscriberCount: number
    readonly activeTurnCount: number
    readonly runningWorkflowCount: number
    readonly reviewQueueSessionCount: number
    readonly sharedPendingReviewCount: number
  }
  readonly pool: CodexAppServerClientPoolStats
  readonly counters: WorkbookAgentObservabilityCounters
}

interface SessionRegistryOptions {
  readonly maxSessions: number
  readonly now: () => number
}

const EMPTY_POOL_STATS: CodexAppServerClientPoolStats = {
  slotCount: 0,
  boundThreadCount: 0,
  activeTurnCount: 0,
  queuedTurnCount: 0,
  maxClients: 0,
  maxConcurrentTurnsPerClient: 0,
  maxQueuedTurnsPerClient: 0,
}

function createEmptyCounters(): { -readonly [K in WorkbookAgentObservabilityCounterName]: number } {
  return {
    turnBackpressureCount: 0,
    workflowStartedCount: 0,
    workflowCompletedCount: 0,
    workflowFailedCount: 0,
    workflowCancelledCount: 0,
    sharedReviewApprovedCount: 0,
    sharedReviewRejectedCount: 0,
    sharedRecommendationApprovedCount: 0,
    sharedRecommendationRejectedCount: 0,
  }
}

function firstPendingReviewItem(sessionState: WorkbookAgentThreadState): WorkbookAgentReviewQueueItem | null {
  return sessionState.durable.reviewQueueItems[0] ?? null
}

function hasPendingSharedOwnerReview(reviewItem: WorkbookAgentReviewQueueItem | null): boolean {
  return reviewItem?.reviewMode === 'ownerReview' && reviewItem.status === 'pending'
}

function countSharedPendingReviews(sessions: readonly WorkbookAgentThreadState[]): number {
  return sessions.filter((sessionState) => hasPendingSharedOwnerReview(firstPendingReviewItem(sessionState))).length
}

export function createDisabledWorkbookAgentObservabilitySnapshot(now: number): WorkbookAgentObservabilitySnapshot {
  return {
    enabled: false,
    generatedAtUnixMs: now,
    featureFlags: {
      sharedThreadsEnabled: false,
      workflowRunnerEnabled: false,
      autoApplyLowRiskEnabled: false,
      formulaWorkflowFamilyEnabled: false,
      formattingWorkflowFamilyEnabled: false,
      importWorkflowFamilyEnabled: false,
      rollupWorkflowFamilyEnabled: false,
      structuralWorkflowFamilyEnabled: false,
      allowlistedUserCount: 0,
      allowlistedDocumentCount: 0,
    },
    sessions: {
      sessionCount: 0,
      subscriberThreadCount: 0,
      subscriberCount: 0,
      activeTurnCount: 0,
      runningWorkflowCount: 0,
      reviewQueueSessionCount: 0,
      sharedPendingReviewCount: 0,
    },
    pool: { ...EMPTY_POOL_STATS },
    counters: createEmptyCounters(),
  }
}

export class WorkbookAgentSessionRegistry {
  private readonly sessions = new Map<string, WorkbookAgentThreadState>()
  private readonly subscribers = new Map<string, Set<(event: WorkbookAgentStreamEvent) => void>>()
  private readonly counters = createEmptyCounters()

  constructor(private readonly options: SessionRegistryOptions) {}

  storeSession(sessionState: WorkbookAgentThreadState, onEvict?: (threadId: string) => void): void {
    this.sessions.set(sessionState.threadId, sessionState)
    this.evictIfNeeded(onEvict)
  }

  tryGetSession(threadId: string): WorkbookAgentThreadState | null {
    return this.sessions.get(threadId) ?? null
  }

  listSessions(): WorkbookAgentThreadState[] {
    return [...this.sessions.values()]
  }

  touch(sessionState: WorkbookAgentThreadState): void {
    sessionState.live.lastAccessedAt = this.options.now()
  }

  incrementCounter(counter: WorkbookAgentObservabilityCounterName): void {
    this.counters[counter] += 1
  }

  subscribe(threadId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void {
    const listeners = this.subscribers.get(threadId) ?? new Set()
    listeners.add(listener)
    this.subscribers.set(threadId, listeners)
    return () => {
      const current = this.subscribers.get(threadId)
      if (!current) {
        return
      }
      current.delete(listener)
      if (current.size === 0) {
        this.subscribers.delete(threadId)
      }
    }
  }

  emit(threadId: string, event: WorkbookAgentStreamEvent): void {
    const listeners = this.subscribers.get(threadId)
    if (!listeners) {
      return
    }
    listeners.forEach((listener) => {
      listener(event)
    })
  }

  emitSnapshot(threadId: string): void {
    const sessionState = this.tryGetSession(threadId)
    if (!sessionState) {
      return
    }
    this.emit(threadId, {
      type: 'snapshot',
      snapshot: buildSnapshot(sessionState),
    })
  }

  getObservabilitySnapshot(input: {
    readonly featureFlags: WorkbookAgentFeatureFlags
    readonly poolStats: CodexAppServerClientPoolStats
  }): WorkbookAgentObservabilitySnapshot {
    const sessions = this.listSessions()
    const subscriberSets = [...this.subscribers.values()]
    return {
      enabled: true,
      generatedAtUnixMs: this.options.now(),
      featureFlags: {
        sharedThreadsEnabled: input.featureFlags.sharedThreadsEnabled,
        workflowRunnerEnabled: input.featureFlags.workflowRunnerEnabled,
        autoApplyLowRiskEnabled: input.featureFlags.autoApplyLowRiskEnabled,
        formulaWorkflowFamilyEnabled: input.featureFlags.formulaWorkflowFamilyEnabled,
        formattingWorkflowFamilyEnabled: input.featureFlags.formattingWorkflowFamilyEnabled,
        importWorkflowFamilyEnabled: input.featureFlags.importWorkflowFamilyEnabled,
        rollupWorkflowFamilyEnabled: input.featureFlags.rollupWorkflowFamilyEnabled,
        structuralWorkflowFamilyEnabled: input.featureFlags.structuralWorkflowFamilyEnabled,
        allowlistedUserCount: input.featureFlags.allowlistedUserIds.length,
        allowlistedDocumentCount: input.featureFlags.allowlistedDocumentIds.length,
      },
      sessions: {
        sessionCount: sessions.length,
        subscriberThreadCount: this.subscribers.size,
        subscriberCount: subscriberSets.reduce((sum, listeners) => sum + listeners.size, 0),
        activeTurnCount: sessions.filter((sessionState) => sessionState.live.activeTurnId !== null).length,
        runningWorkflowCount: sessions.reduce(
          (sum, sessionState) => sum + sessionState.durable.workflowRuns.filter((run) => run.status === 'running').length,
          0,
        ),
        reviewQueueSessionCount: sessions.filter((sessionState) => sessionState.durable.reviewQueueItems.length > 0).length,
        sharedPendingReviewCount: countSharedPendingReviews(sessions),
      },
      pool: input.poolStats,
      counters: { ...this.counters },
    }
  }

  clear(): void {
    this.sessions.clear()
    this.subscribers.clear()
  }

  private evictIfNeeded(onEvict?: (threadId: string) => void): void {
    if (this.sessions.size <= this.options.maxSessions) {
      return
    }
    const candidates = chooseWorkbookAgentEvictionCandidates(this.listSessions(), this.subscribers)
    while (this.sessions.size > this.options.maxSessions && candidates.length > 0) {
      const evicted = candidates.shift()
      if (!evicted) {
        return
      }
      this.sessions.delete(evicted.threadId)
      this.subscribers.delete(evicted.threadId)
      onEvict?.(evicted.threadId)
    }
  }
}
