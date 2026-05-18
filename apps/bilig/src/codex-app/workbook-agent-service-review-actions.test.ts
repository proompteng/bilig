import {
  createWorkbookAgentCommandBundle,
  toWorkbookAgentReviewQueueItem,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import { describe, expect, it, vi } from 'vitest'
import type { WorkbookAgentReviewActionContext } from './workbook-agent-service-review-actions.js'
import { replayWorkbookAgentExecutionRecord } from './workbook-agent-service-review-actions.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

function createSessionState(
  input: {
    readonly reviewBundle?: WorkbookAgentCommandBundle
    readonly executionRecords?: readonly WorkbookAgentExecutionRecord[]
  } = {},
): WorkbookAgentThreadState {
  return {
    documentId: 'doc-1',
    userId: 'alex@example.com',
    storageActorUserId: 'alex@example.com',
    scope: 'shared',
    executionPolicy: 'ownerReview',
    threadId: 'thr-shared',
    durable: {
      context: null,
      entries: [],
      reviewQueueItems: input.reviewBundle
        ? [
            toWorkbookAgentReviewQueueItem({
              bundle: input.reviewBundle,
              reviewMode: 'ownerReview',
              sharedReview: input.reviewBundle.sharedReview ?? null,
            }),
          ]
        : [],
      executionRecords: [...(input.executionRecords ?? [])],
      workflowRuns: [],
    },
    live: {
      activeTurnId: null,
      status: 'idle',
      lastError: null,
      authorizedUserIds: new Set(['alex@example.com']),
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn: new Map(),
      turnContextByTurn: new Map(),
      lastAccessedAt: 0,
    },
  }
}

function createBundle(bundleId: string): WorkbookAgentCommandBundle {
  return createWorkbookAgentCommandBundle({
    bundleId,
    documentId: 'doc-1',
    threadId: 'thr-shared',
    turnId: 'turn-1',
    goalText: 'Update the workbook',
    baseRevision: 4,
    context: null,
    commands: [
      {
        kind: 'createSheet',
        name: 'Summary',
      },
    ],
    now: 100,
    sharedReview: {
      ownerUserId: 'alex@example.com',
      status: 'pending',
      decidedByUserId: null,
      decidedAtUnixMs: null,
      recommendations: [],
    },
  })
}

function createExecutionRecord(): WorkbookAgentExecutionRecord {
  return {
    id: 'run-1',
    bundleId: 'bundle-replay',
    documentId: 'doc-1',
    threadId: 'thr-shared',
    turnId: 'turn-1',
    actorUserId: 'alex@example.com',
    goalText: 'Replay prior workbook setup',
    planText: null,
    summary: 'Create summary sheet',
    scope: 'workbook',
    riskClass: 'high',
    acceptedScope: 'full',
    appliedBy: 'user',
    baseRevision: 4,
    appliedRevision: 5,
    createdAtUnixMs: 100,
    appliedAtUnixMs: 120,
    context: null,
    commands: [
      {
        kind: 'createSheet',
        name: 'Replay Summary',
      },
    ],
    preview: null,
  }
}

function createContext(): WorkbookAgentReviewActionContext {
  return {
    now: () => 200,
    getWorkbookHeadRevision: vi.fn(async () => 9),
    applyCommandBundleForSessionState: vi.fn(async () => createExecutionRecord()),
    shouldApplyToolBundleImmediately: vi.fn(() => false),
    persistSessionState: vi.fn(async () => undefined),
    emitSnapshot: vi.fn(),
    touchSession: vi.fn(),
  }
}

describe('workbook agent service review actions', () => {
  it('does not overwrite a pending review item when replaying execution history', async () => {
    const existingReviewBundle = createBundle('bundle-existing-review')
    const executionRecord = createExecutionRecord()
    const sessionState = createSessionState({
      reviewBundle: existingReviewBundle,
      executionRecords: [executionRecord],
    })
    const context = createContext()

    await expect(
      replayWorkbookAgentExecutionRecord({
        context,
        sessionState,
        documentId: 'doc-1',
        recordId: executionRecord.id,
        actorUserId: 'alex@example.com',
      }),
    ).rejects.toThrow('Finish the current workbook review item before replaying another workbook change.')

    expect(sessionState.durable.reviewQueueItems).toHaveLength(1)
    expect(sessionState.durable.reviewQueueItems[0]?.id).toBe(existingReviewBundle.id)
    expect(context.getWorkbookHeadRevision).not.toHaveBeenCalled()
    expect(context.applyCommandBundleForSessionState).not.toHaveBeenCalled()
    expect(context.persistSessionState).not.toHaveBeenCalled()
    expect(context.emitSnapshot).not.toHaveBeenCalled()
  })

  it('does not overwrite a review item staged while replay is resolving workbook revision', async () => {
    const existingReviewBundle = createBundle('bundle-existing-review')
    const executionRecord = createExecutionRecord()
    const sessionState = createSessionState({
      executionRecords: [executionRecord],
    })
    let releaseHeadRevision!: () => void
    const headRevisionBlocked = new Promise<void>((resolve) => {
      releaseHeadRevision = resolve
    })
    let resolveHeadRevisionRequested!: () => void
    const headRevisionRequested = new Promise<void>((resolve) => {
      resolveHeadRevisionRequested = resolve
    })
    const context: WorkbookAgentReviewActionContext = {
      ...createContext(),
      getWorkbookHeadRevision: vi.fn(async () => {
        resolveHeadRevisionRequested()
        await headRevisionBlocked
        return 9
      }),
    }

    const replayPromise = replayWorkbookAgentExecutionRecord({
      context,
      sessionState,
      documentId: 'doc-1',
      recordId: executionRecord.id,
      actorUserId: 'alex@example.com',
    })
    await headRevisionRequested
    sessionState.durable.reviewQueueItems = [
      toWorkbookAgentReviewQueueItem({
        bundle: existingReviewBundle,
        reviewMode: 'ownerReview',
        sharedReview: existingReviewBundle.sharedReview ?? null,
      }),
    ]
    releaseHeadRevision()

    await expect(replayPromise).rejects.toThrow('Finish the current workbook review item before replaying another workbook change.')

    expect(sessionState.durable.reviewQueueItems).toHaveLength(1)
    expect(sessionState.durable.reviewQueueItems[0]?.id).toBe(existingReviewBundle.id)
    expect(context.applyCommandBundleForSessionState).not.toHaveBeenCalled()
    expect(context.persistSessionState).not.toHaveBeenCalled()
    expect(context.emitSnapshot).not.toHaveBeenCalled()
  })
})
