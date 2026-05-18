import { createWorkbookAgentCommandBundle } from '@bilig/agent-api'
import { describe, expect, it, vi } from 'vitest'
import type { ZeroSyncService } from '../zero/service.js'
import { finalizeWorkbookAgentPrivateTurnBundle } from './workbook-agent-service-application.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

function createSessionState(): WorkbookAgentThreadState {
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
      activeTurnId: 'turn-2',
      status: 'inProgress',
      lastError: null,
      authorizedUserIds: new Set(['alex@example.com']),
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map([
        ['turn-1', 'Old turn'],
        ['turn-2', 'New turn'],
      ]),
      turnActorUserIdByTurn: new Map([
        ['turn-1', 'alex@example.com'],
        ['turn-2', 'alex@example.com'],
      ]),
      turnContextByTurn: new Map(),
      lastAccessedAt: 0,
    },
  }
}

function createZeroSyncServiceStub(input: {
  readonly inspectWorkbook: ZeroSyncService['inspectWorkbook']
  readonly applyAgentCommandBundle: ZeroSyncService['applyAgentCommandBundle']
}): ZeroSyncService {
  return {
    enabled: true,
    async initialize() {},
    async close() {},
    async handleQuery() {
      throw new Error('not used')
    },
    async handleMutate() {
      throw new Error('not used')
    },
    inspectWorkbook: input.inspectWorkbook,
    async applyServerMutator() {
      throw new Error('not used')
    },
    applyAgentCommandBundle: input.applyAgentCommandBundle,
    async listWorkbookChanges() {
      return []
    },
    async listWorkbookAgentRuns() {
      return []
    },
    async listWorkbookAgentThreadRuns() {
      return []
    },
    async appendWorkbookAgentRun() {},
    async listWorkbookAgentThreadSummaries() {
      return []
    },
    async loadWorkbookAgentThreadState() {
      return null
    },
    async saveWorkbookAgentThreadState() {},
    async listWorkbookThreadWorkflowRuns() {
      return []
    },
    async upsertWorkbookWorkflowRun() {},
    async getWorkbookHeadRevision() {
      return 1
    },
    async loadAuthoritativeEvents() {
      throw new Error('not used')
    },
  }
}

describe('workbook agent service application', () => {
  it('does not auto-apply a queued private bundle from a stale completed turn', async () => {
    const sessionState = createSessionState()
    const staleBundle = createWorkbookAgentCommandBundle({
      documentId: 'doc-1',
      threadId: 'thr-1',
      turnId: 'turn-1',
      goalText: 'Old turn should not mutate workbook after turn-2 starts',
      baseRevision: 1,
      context: null,
      commands: [
        {
          kind: 'writeRange',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          values: [['stale']],
        },
      ],
      now: 100,
    })
    sessionState.live.stagedPrivateBundleByTurn.set('turn-1', staleBundle)
    const inspectWorkbook = vi.fn(async () => {
      throw new Error('stale bundle should not build preview')
    })
    const applyAgentCommandBundle = vi.fn(async () => {
      throw new Error('stale bundle should not apply')
    })

    await finalizeWorkbookAgentPrivateTurnBundle(
      {
        zeroSyncService: createZeroSyncServiceStub({
          inspectWorkbook,
          applyAgentCommandBundle,
        }),
        now: () => 200,
        autoApplyLowRiskEnabled: true,
        isRolloutAllowed: () => true,
        touchSession: vi.fn(),
        resolveTurnActorUserId: () => 'alex@example.com',
      },
      {
        sessionState,
        turnId: 'turn-1',
        turnStatus: 'completed',
      },
    )

    expect(inspectWorkbook).not.toHaveBeenCalled()
    expect(applyAgentCommandBundle).not.toHaveBeenCalled()
    expect(sessionState.live.stagedPrivateBundleByTurn.has('turn-1')).toBe(false)
    expect(sessionState.live.activeTurnId).toBe('turn-2')
    expect(sessionState.live.status).toBe('inProgress')
    expect(sessionState.live.lastError).toBeNull()
    expect(sessionState.durable.entries).toEqual([])
  })
})
