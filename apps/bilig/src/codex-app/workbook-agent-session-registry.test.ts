import type { WorkbookAgentReviewQueueItem } from '@bilig/agent-api'
import { describe, expect, it, vi } from 'vitest'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import { buildSnapshot } from './workbook-agent-service-shared.js'
import { WorkbookAgentSessionRegistry, createDisabledWorkbookAgentObservabilitySnapshot } from './workbook-agent-session-registry.js'

function createReviewItem(overrides: Partial<WorkbookAgentReviewQueueItem> = {}): WorkbookAgentReviewQueueItem {
  return {
    id: 'review-1',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'Review workbook changes',
    summary: 'Update Sheet1!A1',
    scope: 'sheet',
    riskClass: 'medium',
    reviewMode: 'ownerReview',
    ownerUserId: 'alex@example.com',
    status: 'pending',
    decidedByUserId: null,
    decidedAtUnixMs: null,
    recommendations: [],
    baseRevision: 1,
    createdAtUnixMs: 10,
    context: null,
    commands: [],
    affectedRanges: [],
    estimatedAffectedCells: 1,
    ...overrides,
  }
}

function createSessionState(overrides: Partial<WorkbookAgentThreadState> = {}): WorkbookAgentThreadState {
  return {
    documentId: 'doc-1',
    userId: 'alex@example.com',
    storageActorUserId: 'alex@example.com',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    threadId: 'thr-1',
    durable: {
      context: null,
      entries: [],
      reviewQueueItems: [],
      executionRecords: [],
      workflowRuns: [],
    },
    live: {
      activeTurnId: null,
      status: 'idle',
      lastError: null,
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn: new Map(),
      turnContextByTurn: new Map(),
      lastAccessedAt: 1,
    },
    ...overrides,
  }
}

describe('workbook agent session registry', () => {
  it('emits snapshots to thread subscribers and cleans up empty subscriptions', () => {
    const registry = new WorkbookAgentSessionRegistry({
      maxSessions: 4,
      now: () => 100,
    })
    const sessionState = createSessionState()
    registry.storeSession(sessionState)

    const listener = vi.fn()
    const unsubscribe = registry.subscribe(sessionState.threadId, listener)
    registry.emitSnapshot(sessionState.threadId)

    expect(listener).toHaveBeenCalledWith({
      type: 'snapshot',
      snapshot: buildSnapshot(sessionState),
    })

    unsubscribe()
    registry.emitSnapshot(sessionState.threadId)
    expect(listener).toHaveBeenCalledTimes(1)
  })

  it('reports observability counts from active sessions, subscribers, and review queues', () => {
    const registry = new WorkbookAgentSessionRegistry({
      maxSessions: 4,
      now: () => 200,
    })
    const activeSession = createSessionState({
      threadId: 'thr-active',
      live: {
        activeTurnId: 'turn-1',
        status: 'inProgress',
        lastError: null,
        stagedPrivateBundleByTurn: new Map(),
        optimisticUserEntryIdByTurn: new Map(),
        promptByTurn: new Map(),
        turnActorUserIdByTurn: new Map(),
        turnContextByTurn: new Map(),
        lastAccessedAt: 5,
      },
      durable: {
        context: null,
        entries: [],
        reviewQueueItems: [],
        executionRecords: [],
        workflowRuns: [
          {
            runId: 'run-1',
            threadId: 'thr-active',
            startedByUserId: 'alex@example.com',
            workflowTemplate: 'summarizeWorkbook',
            title: 'Summarize Workbook',
            summary: 'Running workflow',
            status: 'running',
            errorMessage: null,
            createdAtUnixMs: 1,
            updatedAtUnixMs: 1,
            completedAtUnixMs: null,
            artifact: null,
            steps: [],
          },
        ],
      },
    })
    const reviewSession = createSessionState({
      threadId: 'thr-review',
      scope: 'shared',
      executionPolicy: 'ownerReview',
      durable: {
        context: null,
        entries: [],
        reviewQueueItems: [createReviewItem({ threadId: 'thr-review' })],
        executionRecords: [],
        workflowRuns: [],
      },
    })

    registry.storeSession(activeSession)
    registry.storeSession(reviewSession)
    registry.subscribe(activeSession.threadId, () => {})
    registry.subscribe(reviewSession.threadId, () => {})
    registry.subscribe(reviewSession.threadId, () => {})
    registry.incrementCounter('workflowStartedCount')
    registry.incrementCounter('sharedReviewApprovedCount')

    const snapshot = registry.getObservabilitySnapshot({
      featureFlags: {
        sharedThreadsEnabled: true,
        workflowRunnerEnabled: true,
        autoApplyLowRiskEnabled: false,
        formulaWorkflowFamilyEnabled: true,
        formattingWorkflowFamilyEnabled: true,
        importWorkflowFamilyEnabled: false,
        rollupWorkflowFamilyEnabled: true,
        structuralWorkflowFamilyEnabled: false,
        allowlistedUserIds: ['alex@example.com', 'pat@example.com'],
        allowlistedDocumentIds: ['doc-1'],
      },
      poolStats: {
        slotCount: 1,
        boundThreadCount: 2,
        activeTurnCount: 1,
        queuedTurnCount: 0,
        maxClients: 4,
        maxConcurrentTurnsPerClient: 1,
        maxQueuedTurnsPerClient: 8,
      },
    })

    expect(snapshot.enabled).toBe(true)
    expect(snapshot.generatedAtUnixMs).toBe(200)
    expect(snapshot.featureFlags.allowlistedUserCount).toBe(2)
    expect(snapshot.featureFlags.allowlistedDocumentCount).toBe(1)
    expect(snapshot.sessions).toEqual({
      sessionCount: 2,
      subscriberThreadCount: 2,
      subscriberCount: 3,
      activeTurnCount: 1,
      runningWorkflowCount: 1,
      reviewQueueSessionCount: 1,
      sharedPendingReviewCount: 1,
    })
    expect(snapshot.counters.workflowStartedCount).toBe(1)
    expect(snapshot.counters.sharedReviewApprovedCount).toBe(1)
  })

  it('evicts the oldest idle unsubscribed session when capacity is exceeded', () => {
    const evicted: string[] = []
    const registry = new WorkbookAgentSessionRegistry({
      maxSessions: 1,
      now: () => 50,
    })
    registry.storeSession(createSessionState({ threadId: 'thr-old', live: { ...createSessionState().live, lastAccessedAt: 1 } }))
    registry.storeSession(
      createSessionState({ threadId: 'thr-new', live: { ...createSessionState().live, lastAccessedAt: 2 } }),
      (threadId) => {
        evicted.push(threadId)
      },
    )

    expect(evicted).toEqual(['thr-old'])
    expect(registry.tryGetSession('thr-old')).toBeNull()
    expect(registry.tryGetSession('thr-new')).not.toBeNull()
  })

  it('keeps subscribed sessions resident and evicts a newer unsubscribed idle session first', () => {
    const evicted: string[] = []
    const registry = new WorkbookAgentSessionRegistry({
      maxSessions: 1,
      now: () => 60,
    })
    registry.storeSession(createSessionState({ threadId: 'thr-kept', live: { ...createSessionState().live, lastAccessedAt: 1 } }))
    registry.subscribe('thr-kept', () => {})
    registry.storeSession(
      createSessionState({ threadId: 'thr-evicted', live: { ...createSessionState().live, lastAccessedAt: 2 } }),
      (threadId) => {
        evicted.push(threadId)
      },
    )

    expect(evicted).toEqual(['thr-evicted'])
    expect(registry.tryGetSession('thr-kept')).not.toBeNull()
    expect(registry.tryGetSession('thr-evicted')).toBeNull()
  })

  it('builds a disabled observability snapshot with zeroed counts', () => {
    expect(createDisabledWorkbookAgentObservabilitySnapshot(42)).toEqual({
      enabled: false,
      generatedAtUnixMs: 42,
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
      pool: {
        slotCount: 0,
        boundThreadCount: 0,
        activeTurnCount: 0,
        queuedTurnCount: 0,
        maxClients: 0,
        maxConcurrentTurnsPerClient: 0,
        maxQueuedTurnsPerClient: 0,
      },
      counters: {
        turnBackpressureCount: 0,
        workflowStartedCount: 0,
        workflowCompletedCount: 0,
        workflowFailedCount: 0,
        workflowCancelledCount: 0,
        sharedReviewApprovedCount: 0,
        sharedReviewRejectedCount: 0,
        sharedRecommendationApprovedCount: 0,
        sharedRecommendationRejectedCount: 0,
      },
    })
  })
})
