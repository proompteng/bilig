import { describe, expect, it } from 'vitest'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import { createWorkbookAgentCommandBundle, toWorkbookAgentCommandBundle, toWorkbookAgentReviewQueueItem } from '@bilig/agent-api'
import {
  createWorkbookAgentReviewQueueItem,
  createWorkbookAgentDismissReviewEntry,
  getCurrentWorkbookAgentReviewItem,
  replaceCurrentWorkbookAgentReviewItem,
  requireWorkbookAgentReviewItem,
  stageWorkbookAgentReviewBundle,
  transitionWorkbookAgentSharedReview,
} from './workbook-agent-review-transitions.js'

function createSessionState(): WorkbookAgentThreadState {
  return {
    documentId: 'doc-1',
    userId: 'alex@example.com',
    storageActorUserId: 'alex@example.com',
    scope: 'shared',
    executionPolicy: 'ownerReview',
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
      lastAccessedAt: 0,
    },
  }
}

function createReviewItem() {
  const bundle = createWorkbookAgentCommandBundle({
    bundleId: 'bundle-1',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'Normalize the workbook',
    baseRevision: 4,
    context: null,
    commands: [{ kind: 'createSheet', name: 'Summary' }],
    now: 100,
    sharedReview: {
      ownerUserId: 'alex@example.com',
      status: 'pending',
      decidedByUserId: null,
      decidedAtUnixMs: null,
      recommendations: [],
    },
  })
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: 'ownerReview',
    sharedReview: bundle.sharedReview ?? null,
  })
}

describe('workbook agent review transitions', () => {
  it('requires the expected review item id', () => {
    expect(() =>
      requireWorkbookAgentReviewItem({
        reviewItem: null,
        reviewItemId: 'bundle-1',
        notFoundMessage: 'Workbook review item was not found.',
      }),
    ).toThrow('Workbook review item was not found.')
  })

  it('transitions owner review decisions and increments the owner counter kind', () => {
    const result = transitionWorkbookAgentSharedReview({
      sessionState: createSessionState(),
      reviewItem: createReviewItem(),
      decision: 'approved',
      reviewerUserId: 'alex@example.com',
      now: 200,
    })

    expect(result.counter).toBe('sharedReviewApprovedCount')
    expect(result.nextReviewItem.status).toBe('approved')
    expect(result.nextReviewItem.decidedByUserId).toBe('alex@example.com')
    expect(result.entryText).toContain('Approved shared review item')
  })

  it('records collaborator recommendations without deciding the owner review', () => {
    const result = transitionWorkbookAgentSharedReview({
      sessionState: createSessionState(),
      reviewItem: createReviewItem(),
      decision: 'approved',
      reviewerUserId: 'pat@example.com',
      now: 200,
    })

    expect(result.counter).toBe('sharedRecommendationApprovedCount')
    expect(result.nextReviewItem.status).toBe('pending')
    expect(result.nextReviewItem.decidedByUserId).toBeNull()
    expect(result.nextReviewItem.recommendations).toEqual([
      expect.objectContaining({
        userId: 'pat@example.com',
        decision: 'approved',
      }),
    ])
    expect(result.entryText).toContain('pat@example.com shared a ready-to-apply review recommendation')
  })

  it('reads and replaces the current review item', () => {
    const sessionState = createSessionState()
    const reviewItem = createReviewItem()

    expect(getCurrentWorkbookAgentReviewItem(sessionState)).toBeNull()

    replaceCurrentWorkbookAgentReviewItem(sessionState, reviewItem)
    expect(getCurrentWorkbookAgentReviewItem(sessionState)).toBe(reviewItem)

    replaceCurrentWorkbookAgentReviewItem(sessionState, null)
    expect(getCurrentWorkbookAgentReviewItem(sessionState)).toBeNull()
  })

  it('stages review bundles as preview entries', () => {
    const sessionState = createSessionState()
    const bundle = toWorkbookAgentCommandBundle(createReviewItem())

    stageWorkbookAgentReviewBundle({
      sessionState,
      turnId: bundle.turnId,
      bundle,
    })

    expect(getCurrentWorkbookAgentReviewItem(sessionState)).toEqual(
      expect.objectContaining({
        id: 'bundle-1',
        reviewMode: 'ownerReview',
      }),
    )
    expect(sessionState.durable.entries).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: 'system-preview:bundle-1',
          kind: 'system',
          turnId: 'turn-1',
        }),
      ]),
    )
  })

  it('creates review queue items that reflect shared-owner review requirements', () => {
    const manual = createWorkbookAgentReviewQueueItem({
      sessionState: {
        ...createSessionState(),
        scope: 'private',
        executionPolicy: 'autoApplyAll',
      },
      bundle: toWorkbookAgentCommandBundle(createReviewItem()),
    })
    expect(manual.reviewMode).toBe('manual')

    const ownerReview = createWorkbookAgentReviewQueueItem({
      sessionState: createSessionState(),
      bundle: toWorkbookAgentCommandBundle(createReviewItem()),
    })
    expect(ownerReview.reviewMode).toBe('ownerReview')
  })

  it('creates deterministic dismiss entries', () => {
    expect(
      createWorkbookAgentDismissReviewEntry({
        reviewItem: createReviewItem(),
        now: 300,
      }),
    ).toEqual(
      expect.objectContaining({
        id: 'system-dismiss:bundle-1:300',
        kind: 'system',
        text: 'Cleared workbook review item: Create sheet Summary',
      }),
    )
  })
})
