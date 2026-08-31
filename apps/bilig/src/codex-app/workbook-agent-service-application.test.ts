import { createWorkbookAgentCommandBundle } from '@bilig/agent-api'
import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it, vi } from 'vitest'
import type { ZeroSyncService } from '../zero/service.js'
import { buildWorkbookSourceProjection } from '../zero/projection.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import {
  applyWorkbookAgentCommandBundleForSessionState,
  finalizeWorkbookAgentPrivateTurnBundle,
} from './workbook-agent-service-application.js'
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
  readonly appendWorkbookAgentRun?: ZeroSyncService['appendWorkbookAgentRun']
}): ZeroSyncService {
  return {
    enabled: true,
    isReady: () => true,
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
    async applyWorkbookPlanData() {
      throw new Error('not used')
    },
    async listWorkbookChanges() {
      return []
    },
    async listWorkbookAgentRuns() {
      return []
    },
    async listWorkbookAgentThreadRuns() {
      return []
    },
    appendWorkbookAgentRun: input.appendWorkbookAgentRun ?? (async () => {}),
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
  it('attaches runtime command result proof to execution records', async () => {
    const sessionState = createSessionState()
    const bundle = createWorkbookAgentCommandBundle({
      documentId: 'doc-1',
      threadId: 'thr-1',
      turnId: 'turn-2',
      goalText: 'Apply a workbook command with proof',
      baseRevision: 1,
      context: null,
      commands: [
        {
          kind: 'writeRange',
          sheetName: 'Sheet1',
          startAddress: 'C3',
          values: [[2]],
        },
      ],
      now: 100,
    })
    const engine = new SpreadsheetEngine({
      workbookName: 'Proof Workbook',
      replicaId: 'test:command-result-proof',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const runtime = {
      documentId: 'doc-1',
      engine,
      projection: buildWorkbookSourceProjection('doc-1', engine.exportSnapshot(), {
        revision: 1,
        calculatedRevision: 1,
        ownerUserId: 'alex@example.com',
        updatedBy: 'alex@example.com',
        updatedAt: '2026-05-23T00:00:00.000Z',
      }),
      headRevision: 1,
      calculatedRevision: 1,
      ownerUserId: 'alex@example.com',
    } satisfies WorkbookRuntime
    const commandResult = {
      status: 'applied' as const,
      bundleId: bundle.id,
      targetRevision: 1,
      idempotencyKey: bundle.id,
      commandCount: 1,
      touchedRanges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'C3',
        },
      ],
      touchedCellCount: 1,
      receipts: [
        {
          status: 'applied' as const,
          featureId: 'workbook-agent',
          commandId: 'workbookAgent.writeRange',
          category: 'mutation' as const,
          changedRanges: [
            {
              sheetName: 'Sheet1',
              startAddress: 'C3',
              endAddress: 'C3',
            },
          ],
        },
      ],
      matched: null,
      changedRanges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'C3',
        },
      ],
      revision: 2,
    }
    const inspectWorkbook: ZeroSyncService['inspectWorkbook'] = vi.fn(async (_documentId, task) => task(runtime))
    const applyAgentCommandBundle: ZeroSyncService['applyAgentCommandBundle'] = vi.fn(async (_documentId, _bundle, preview) => ({
      revision: 2,
      preview,
      commandResult,
    }))
    const appendWorkbookAgentRun = vi.fn(async () => {})

    const record = await applyWorkbookAgentCommandBundleForSessionState(
      {
        zeroSyncService: createZeroSyncServiceStub({
          inspectWorkbook,
          applyAgentCommandBundle,
          appendWorkbookAgentRun,
        }),
        now: () => 200,
        autoApplyLowRiskEnabled: true,
        isRolloutAllowed: () => true,
        touchSession: vi.fn(),
      },
      {
        sessionState,
        commandBundle: bundle,
        actorUserId: 'alex@example.com',
        appliedBy: 'auto',
      },
    )

    expect(record.commandResult).toEqual(commandResult)
    expect(appendWorkbookAgentRun).toHaveBeenCalledWith(expect.objectContaining({ commandResult }))
    expect(sessionState.durable.executionRecords[0]).toEqual(expect.objectContaining({ commandResult }))
  })

  it('fails closed when runtime apply omits generic command result proof', async () => {
    const sessionState = createSessionState()
    const bundle = createWorkbookAgentCommandBundle({
      documentId: 'doc-1',
      threadId: 'thr-1',
      turnId: 'turn-2',
      goalText: 'Apply a workbook command with proof',
      baseRevision: 1,
      context: null,
      commands: [
        {
          kind: 'writeRange',
          sheetName: 'Sheet1',
          startAddress: 'C3',
          values: [[2]],
        },
      ],
      now: 100,
    })
    const engine = new SpreadsheetEngine({
      workbookName: 'Proof Workbook',
      replicaId: 'test:missing-command-result-proof',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const runtime = {
      documentId: 'doc-1',
      engine,
      projection: buildWorkbookSourceProjection('doc-1', engine.exportSnapshot(), {
        revision: 1,
        calculatedRevision: 1,
        ownerUserId: 'alex@example.com',
        updatedBy: 'alex@example.com',
        updatedAt: '2026-05-23T00:00:00.000Z',
      }),
      headRevision: 1,
      calculatedRevision: 1,
      ownerUserId: 'alex@example.com',
    } satisfies WorkbookRuntime
    const inspectWorkbook: ZeroSyncService['inspectWorkbook'] = vi.fn(async (_documentId, task) => task(runtime))
    const applyAgentCommandBundle: ZeroSyncService['applyAgentCommandBundle'] = vi.fn(async (_documentId, _bundle, preview) => ({
      revision: 2,
      preview,
    }))
    const appendWorkbookAgentRun = vi.fn(async () => {})

    await expect(
      applyWorkbookAgentCommandBundleForSessionState(
        {
          zeroSyncService: createZeroSyncServiceStub({
            inspectWorkbook,
            applyAgentCommandBundle,
            appendWorkbookAgentRun,
          }),
          now: () => 200,
          autoApplyLowRiskEnabled: true,
          isRolloutAllowed: () => true,
          touchSession: vi.fn(),
        },
        {
          sessionState,
          commandBundle: bundle,
          actorUserId: 'alex@example.com',
          appliedBy: 'auto',
        },
      ),
    ).rejects.toMatchObject({
      code: 'WORKBOOK_AGENT_COMMAND_RESULT_REQUIRED',
      retryable: false,
      statusCode: 500,
    })
    expect(appendWorkbookAgentRun).not.toHaveBeenCalled()
    expect(sessionState.durable.executionRecords).toEqual([])
  })

  it('fails closed when runtime apply returns command result proof for a different bundle', async () => {
    const sessionState = createSessionState()
    const bundle = createWorkbookAgentCommandBundle({
      documentId: 'doc-1',
      threadId: 'thr-1',
      turnId: 'turn-2',
      goalText: 'Apply a workbook command with proof',
      baseRevision: 1,
      context: null,
      commands: [
        {
          kind: 'writeRange',
          sheetName: 'Sheet1',
          startAddress: 'C3',
          values: [[2]],
        },
      ],
      now: 100,
    })
    const engine = new SpreadsheetEngine({
      workbookName: 'Proof Workbook',
      replicaId: 'test:mismatched-command-result-proof',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const runtime = {
      documentId: 'doc-1',
      engine,
      projection: buildWorkbookSourceProjection('doc-1', engine.exportSnapshot(), {
        revision: 1,
        calculatedRevision: 1,
        ownerUserId: 'alex@example.com',
        updatedBy: 'alex@example.com',
        updatedAt: '2026-05-23T00:00:00.000Z',
      }),
      headRevision: 1,
      calculatedRevision: 1,
      ownerUserId: 'alex@example.com',
    } satisfies WorkbookRuntime
    const commandResult = {
      status: 'applied' as const,
      bundleId: `${bundle.id}:other`,
      targetRevision: 1,
      idempotencyKey: bundle.id,
      commandCount: 1,
      touchedRanges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'C3',
        },
      ],
      touchedCellCount: 1,
      receipts: [
        {
          status: 'applied' as const,
          featureId: 'workbook-agent',
          commandId: 'workbookAgent.writeRange',
          category: 'mutation' as const,
          changedRanges: [
            {
              sheetName: 'Sheet1',
              startAddress: 'C3',
              endAddress: 'C3',
            },
          ],
        },
      ],
      matched: null,
      changedRanges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'C3',
        },
      ],
      revision: 2,
    }
    const inspectWorkbook: ZeroSyncService['inspectWorkbook'] = vi.fn(async (_documentId, task) => task(runtime))
    const applyAgentCommandBundle: ZeroSyncService['applyAgentCommandBundle'] = vi.fn(async (_documentId, _bundle, preview) => ({
      revision: 2,
      preview,
      commandResult,
    }))
    const appendWorkbookAgentRun = vi.fn(async () => {})

    await expect(
      applyWorkbookAgentCommandBundleForSessionState(
        {
          zeroSyncService: createZeroSyncServiceStub({
            inspectWorkbook,
            applyAgentCommandBundle,
            appendWorkbookAgentRun,
          }),
          now: () => 200,
          autoApplyLowRiskEnabled: true,
          isRolloutAllowed: () => true,
          touchSession: vi.fn(),
        },
        {
          sessionState,
          commandBundle: bundle,
          actorUserId: 'alex@example.com',
          appliedBy: 'auto',
        },
      ),
    ).rejects.toMatchObject({
      code: 'WORKBOOK_AGENT_INVALID_COMMAND_RESULT',
      retryable: false,
      statusCode: 500,
    })
    expect(appendWorkbookAgentRun).not.toHaveBeenCalled()
    expect(sessionState.durable.executionRecords).toEqual([])
  })

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
