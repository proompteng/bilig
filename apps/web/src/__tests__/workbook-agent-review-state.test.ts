import { describe, expect, it } from 'vitest'
import {
  decodeWorkbookAgentReviewItems,
  resolvePrimaryWorkbookAgentReviewItem,
  resolveWorkbookAgentReviewOwnerUserId,
} from '../workbook-agent-review-state.js'
import { toWorkbookAgentReviewQueueItem, type WorkbookAgentCommandBundle } from '@bilig/agent-api'

function createSnapshot(overrides: Record<string, unknown> = {}) {
  return {
    sessionId: 'agent-session-1',
    documentId: 'doc-1',
    threadId: 'thr-1',
    executionPolicy: 'autoApplyAll',
    scope: 'private',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
    context: null,
    entries: [],
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
  }
}

function createReviewQueueItem(bundle: WorkbookAgentCommandBundle) {
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: bundle.sharedReview ? 'ownerReview' : 'manual',
    ...(bundle.sharedReview ? { sharedReview: bundle.sharedReview } : {}),
  })
}

describe('workbook agent review state', () => {
  it('decodes the current review item from the session snapshot', () => {
    const snapshot = createSnapshot({
      reviewQueueItems: [
        createReviewQueueItem({
          id: 'bundle-1',
          documentId: 'doc-1',
          threadId: 'thr-1',
          turnId: 'turn-1',
          goalText: 'Bold the current cell',
          summary: 'Format Sheet1!A1',
          scope: 'selection',
          riskClass: 'low',
          baseRevision: 1,
          createdAtUnixMs: 1,
          context: null,
          commands: [],
          affectedRanges: [],
          estimatedAffectedCells: 1,
          sharedReview: null,
        }),
      ],
    })

    expect(resolvePrimaryWorkbookAgentReviewItem(decodeWorkbookAgentReviewItems(snapshot.reviewQueueItems))?.id).toBe('bundle-1')
  })

  it('derives owner review state for shared high-risk work', () => {
    const reviewItem = createReviewQueueItem({
      id: 'bundle-1',
      documentId: 'doc-1',
      threadId: 'thr-1',
      turnId: 'turn-1',
      goalText: 'Normalize the workbook',
      summary: 'Normalize shared workbook structure',
      scope: 'workbook',
      riskClass: 'high',
      baseRevision: 1,
      createdAtUnixMs: 1,
      context: null,
      commands: [],
      affectedRanges: [],
      estimatedAffectedCells: 0,
      sharedReview: {
        ownerUserId: 'alex@example.com',
        status: 'pending',
        decidedByUserId: null,
        decidedAtUnixMs: null,
        recommendations: [],
      },
    })

    expect(
      resolveWorkbookAgentReviewOwnerUserId({
        reviewItem,
        sessionScope: 'shared',
        activeThreadOwnerUserId: 'casey@example.com',
      }),
    ).toBe('alex@example.com')
  })
})
