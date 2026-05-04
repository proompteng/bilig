import {
  createWorkbookAgentCommandBundle,
  toWorkbookAgentCommandBundle,
  toWorkbookAgentReviewQueueItem,
  type CodexThread,
} from '@bilig/agent-api'
import { describe, expect, it } from 'vitest'
import {
  clearLegacyPrivateBootstrapReviewItem,
  createWorkbookAgentBootstrappedSessionState,
  LEGACY_PRIVATE_BOOTSTRAP_REVIEW_MESSAGE,
  planWorkbookAgentBootstrapReviewRecovery,
  rebaseWorkbookAgentBootstrapReviewItem,
} from './workbook-agent-service-bootstrap.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import type { WorkbookAgentLoadedThreadState } from './workbook-agent-thread-repository.js'

function createLoadedThreadState(): WorkbookAgentLoadedThreadState {
  return {
    threadState: {
      documentId: 'doc-1',
      threadId: 'thr-1',
      actorUserId: 'alex@example.com',
      scope: 'private',
      executionPolicy: 'autoApplyAll',
      context: {
        selection: {
          sheetName: 'Sheet1',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      entries: [
        {
          id: 'system-existing',
          kind: 'system',
          turnId: null,
          text: 'Recovered durable history.',
          phase: null,
          toolName: null,
          toolStatus: null,
          argumentsText: null,
          outputText: null,
          success: null,
          citations: [],
        },
      ],
      reviewQueueItems: [],
      updatedAtUnixMs: 100,
    },
    executionRecords: [],
    workflowRuns: [],
  }
}

function createReviewItem(overrides: Partial<ReturnType<typeof createWorkbookAgentCommandBundle>> = {}) {
  const bundle = createWorkbookAgentCommandBundle({
    bundleId: 'bundle-1',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'Normalize workbook',
    baseRevision: 4,
    context: null,
    commands: [{ kind: 'writeRange', sheetName: 'Sheet1', startAddress: 'B2', values: [[42]] }],
    now: 100,
    ...overrides,
  })
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: bundle.sharedReview ? 'ownerReview' : 'manual',
    ...(bundle.sharedReview ? { sharedReview: bundle.sharedReview } : {}),
  })
}

function createSessionState(reviewQueueItems = [createReviewItem()]): WorkbookAgentThreadState {
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
      reviewQueueItems,
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
      lastAccessedAt: 0,
    },
  }
}

describe('workbook agent service bootstrap helpers', () => {
  it('builds bootstrapped session state from durable history and live thread state', () => {
    const liveThread: CodexThread = {
      id: 'thr-1',
      preview: '',
      turns: [
        {
          id: 'turn-1',
          status: 'inProgress',
          items: [],
          error: null,
        },
      ],
    }

    const sessionState = createWorkbookAgentBootstrappedSessionState({
      documentId: 'doc-1',
      userId: 'alex@example.com',
      threadId: 'thr-1',
      requestedContext: {
        selection: {
          sheetName: 'Sheet2',
          address: 'C7',
        },
        viewport: {
          rowStart: 2,
          rowEnd: 22,
          colStart: 1,
          colEnd: 11,
        },
      },
      durableThreadSession: createLoadedThreadState(),
      liveThread,
      sessionBootstrapError: null,
      now: 200,
    })

    expect(sessionState.threadId).toBe('thr-1')
    expect(sessionState.scope).toBe('private')
    expect(sessionState.executionPolicy).toBe('autoApplyAll')
    expect(sessionState.durable.context).toEqual(
      expect.objectContaining({
        selection: expect.objectContaining({
          sheetName: 'Sheet2',
          address: 'C7',
        }),
      }),
    )
    expect(sessionState.durable.entries).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: 'system-existing',
          text: 'Recovered durable history.',
        }),
      ]),
    )
    expect(sessionState.live.activeTurnId).toBe('turn-1')
    expect(sessionState.live.status).toBe('inProgress')
    expect(sessionState.live.lastError).toBeNull()
    expect(sessionState.live.lastAccessedAt).toBe(200)
  })

  it('plans bootstrap review recovery for auto-apply, clear-legacy, and no-op cases', () => {
    const sessionState = createSessionState([createReviewItem()])
    expect(
      planWorkbookAgentBootstrapReviewRecovery({
        sessionState,
        rolloutAllowed: true,
        autoApplyLowRiskEnabled: true,
      }),
    ).toMatchObject({ kind: 'autoApply' })

    expect(
      planWorkbookAgentBootstrapReviewRecovery({
        sessionState,
        rolloutAllowed: false,
        autoApplyLowRiskEnabled: true,
      }),
    ).toMatchObject({ kind: 'clearLegacy' })

    expect(
      planWorkbookAgentBootstrapReviewRecovery({
        sessionState: {
          ...sessionState,
          scope: 'shared',
        },
        rolloutAllowed: true,
        autoApplyLowRiskEnabled: true,
      }),
    ).toEqual({ kind: 'none' })
  })

  it('rebases bootstrap review items to the current revision while preserving owner review state', () => {
    const rebased = rebaseWorkbookAgentBootstrapReviewItem({
      reviewItem: createReviewItem({
        sharedReview: {
          ownerUserId: 'alex@example.com',
          status: 'pending',
          decidedByUserId: null,
          decidedAtUnixMs: null,
          recommendations: [],
        },
      }),
      currentRevision: 9,
      fallbackOwnerUserId: 'alex@example.com',
    })

    expect(rebased.reviewMode).toBe('ownerReview')
    expect(rebased.ownerUserId).toBe('alex@example.com')
    expect(toWorkbookAgentCommandBundle(rebased).baseRevision).toBe(9)
  })

  it('clears legacy private review items and records the recovery entry', () => {
    const sessionState = createSessionState([createReviewItem()])
    const reviewItem = sessionState.durable.reviewQueueItems[0]
    if (!reviewItem) {
      throw new Error('Expected review item')
    }

    clearLegacyPrivateBootstrapReviewItem({
      sessionState,
      reviewItem,
      now: 300,
    })

    expect(sessionState.durable.reviewQueueItems).toEqual([])
    expect(sessionState.live.lastError).toBe(LEGACY_PRIVATE_BOOTSTRAP_REVIEW_MESSAGE)
    expect(sessionState.durable.entries).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: 'system-private-review-item-cleared:bundle-1:300',
          text: 'Cleared a legacy private review item that could not be resumed automatically: Write cells in Sheet1!B2',
        }),
      ]),
    )
  })
})
