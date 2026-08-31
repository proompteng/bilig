import { isWorkbookAgentCommandBundle } from '@bilig/agent-api'
import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it, vi } from 'vitest'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { WorkbookAgentThreadStateRecord } from '../zero/workbook-chat-thread-store.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { createWorkbookAgentService } from './workbook-agent-service.js'

import {
  FakeCodexTransport,
  createPreviewSummary,
  createReviewQueueItem,
  createZeroSyncStub,
  getPrimaryReviewBundle,
  startWorkbookAgentTestTurn,
} from './workbook-agent-service.test-helpers.js'

describe('workbook agent service review apply and recovery', () => {
  it('keeps the review item when authoritative apply rejects a stale preview', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle: vi.fn(async () => {
          throw createWorkbookAgentServiceError({
            code: 'WORKBOOK_AGENT_PREVIEW_STALE',
            message: 'Workbook changed after preview. Replay the plan to stage a fresh change set.',
            statusCode: 409,
            retryable: true,
          })
        }),
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options
          return fakeCodex
        },
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          scope: 'shared',
          executionPolicy: 'ownerReview',
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B2',
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      })
      await startWorkbookAgentTestTurn(service, {
        threadId: snapshot.threadId,
      })

      await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-test',
        turnId: 'turn-1',
        callId: 'call-1',
        tool: 'bilig_write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          values: [[42]],
        },
      })

      const pending = getPrimaryReviewBundle(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }),
      )
      if (!isWorkbookAgentCommandBundle(pending)) {
        throw new Error('Expected a staged review item')
      }

      await service.reviewReviewItem({
        documentId: 'doc-1',
        threadId: 'thr-test',
        reviewItemId: pending.id,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          decision: 'approved',
        },
      })

      await expect(
        service.applyReviewItem({
          documentId: 'doc-1',
          threadId: 'thr-test',
          reviewItemId: pending.id,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          appliedBy: 'user',
        }),
      ).rejects.toThrow('Replay the plan to stage a fresh change set.')

      const afterFailure = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      const afterFailureBundle = getPrimaryReviewBundle(afterFailure)
      expect(isWorkbookAgentCommandBundle(afterFailureBundle)).toBe(true)
      if (!isWorkbookAgentCommandBundle(afterFailureBundle)) {
        throw new Error('Expected the review item to remain staged')
      }
      expect(afterFailureBundle.id).toBe(pending.id)
      expect(afterFailure.executionRecords).toEqual([])
    } finally {
      await service.close()
    }
  })

  it('applies a selected command subset and re-stages the remaining plan', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 7,
      preview: createPreviewSummary({
        cellDiffs: [
          {
            sheetName: 'Sheet1',
            address: 'C3',
            beforeInput: null,
            beforeFormula: null,
            afterInput: 2,
            afterFormula: null,
            changeKinds: ['input'],
          },
        ],
        effectSummary: {
          displayedCellDiffCount: 1,
          truncatedCellDiffs: false,
          inputChangeCount: 1,
          formulaChangeCount: 0,
          styleChangeCount: 0,
          numberFormatChangeCount: 0,
          structuralChangeCount: 0,
        },
      }),
    }))
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options
          return fakeCodex
        },
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          scope: 'shared',
          executionPolicy: 'ownerReview',
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B2',
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      })
      await startWorkbookAgentTestTurn(service, {
        threadId: snapshot.threadId,
      })

      await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-test',
        turnId: 'turn-1',
        callId: 'call-1',
        tool: 'bilig_write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          values: [[1]],
        },
      })
      await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-test',
        turnId: 'turn-1',
        callId: 'call-2',
        tool: 'bilig_write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          values: [[2]],
        },
      })

      const pending = getPrimaryReviewBundle(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }),
      )
      if (!isWorkbookAgentCommandBundle(pending)) {
        throw new Error('Expected a staged review item')
      }

      await service.reviewReviewItem({
        documentId: 'doc-1',
        threadId: 'thr-test',
        reviewItemId: pending.id,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          decision: 'approved',
        },
      })

      const applied = await service.applyReviewItem({
        documentId: 'doc-1',
        threadId: 'thr-test',
        reviewItemId: pending.id,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        appliedBy: 'user',
        commandIndexes: [1],
      })

      expect(applyAgentCommandBundle).toHaveBeenCalledWith(
        'doc-1',
        expect.objectContaining({
          id: pending.id,
          summary: 'Write cells in Sheet1!C3',
          commands: [
            {
              kind: 'writeRange',
              sheetName: 'Sheet1',
              startAddress: 'C3',
              values: [[2]],
            },
          ],
        }),
        expect.objectContaining({
          ranges: [
            {
              sheetName: 'Sheet1',
              startAddress: 'C3',
              endAddress: 'C3',
              role: 'target',
            },
          ],
        }),
        expect.objectContaining({
          userID: 'alex@example.com',
        }),
      )
      expect(applied.executionRecords[0]).toEqual(
        expect.objectContaining({
          bundleId: pending.id,
          acceptedScope: 'partial',
          summary: 'Write cells in Sheet1!C3',
          commands: [
            {
              kind: 'writeRange',
              sheetName: 'Sheet1',
              startAddress: 'C3',
              values: [[2]],
            },
          ],
        }),
      )
      expect(getPrimaryReviewBundle(applied)).toEqual(
        expect.objectContaining({
          baseRevision: 7,
          summary: 'Write cells in Sheet1!B2',
          commands: [
            {
              kind: 'writeRange',
              sheetName: 'Sheet1',
              startAddress: 'B2',
              values: [[1]],
            },
          ],
        }),
      )
      const remainingBundle = getPrimaryReviewBundle(applied)
      if (!isWorkbookAgentCommandBundle(remainingBundle)) {
        throw new Error('Expected the remaining staged bundle to stay pending')
      }
      expect(remainingBundle.id).not.toBe(pending.id)
      expect(appendWorkbookAgentRun).toHaveBeenCalledWith(
        expect.objectContaining({
          acceptedScope: 'partial',
          summary: 'Write cells in Sheet1!C3',
        }),
      )
    } finally {
      await service.close()
    }
  })

  it('does not resurrect a dismissed review item after a concurrent partial apply finishes', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    let releaseApply!: () => void
    const applyBlocked = new Promise<void>((resolve) => {
      releaseApply = resolve
    })
    let resolveApplyStarted!: () => void
    const applyStarted = new Promise<void>((resolve) => {
      resolveApplyStarted = resolve
    })
    const applyAgentCommandBundle = vi.fn(async () => {
      resolveApplyStarted()
      await applyBlocked
      return {
        revision: 7,
        preview: createPreviewSummary({
          cellDiffs: [
            {
              sheetName: 'Sheet1',
              address: 'C3',
              beforeInput: null,
              beforeFormula: null,
              afterInput: 2,
              afterFormula: null,
              changeKinds: ['input'],
            },
          ],
          effectSummary: {
            displayedCellDiffCount: 1,
            truncatedCellDiffs: false,
            inputChangeCount: 1,
            formulaChangeCount: 0,
            styleChangeCount: 0,
            numberFormatChangeCount: 0,
            structuralChangeCount: 0,
          },
        }),
      }
    })
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options
          return fakeCodex
        },
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          scope: 'shared',
          executionPolicy: 'ownerReview',
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B2',
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      })
      await startWorkbookAgentTestTurn(service, {
        threadId: snapshot.threadId,
      })

      await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-test',
        turnId: 'turn-1',
        callId: 'call-1',
        tool: 'bilig_write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          values: [[1]],
        },
      })
      await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-test',
        turnId: 'turn-1',
        callId: 'call-2',
        tool: 'bilig_write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          values: [[2]],
        },
      })

      const pending = getPrimaryReviewBundle(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }),
      )
      if (!isWorkbookAgentCommandBundle(pending)) {
        throw new Error('Expected a staged review item')
      }

      await service.reviewReviewItem({
        documentId: 'doc-1',
        threadId: 'thr-test',
        reviewItemId: pending.id,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          decision: 'approved',
        },
      })

      const applyPromise = service.applyReviewItem({
        documentId: 'doc-1',
        threadId: 'thr-test',
        reviewItemId: pending.id,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        appliedBy: 'user',
        commandIndexes: [1],
      })
      await applyStarted

      const dismissed = await service.dismissReviewItem({
        documentId: 'doc-1',
        threadId: 'thr-test',
        reviewItemId: pending.id,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      expect(dismissed.reviewQueueItems).toEqual([])

      releaseApply()
      const applied = await applyPromise

      expect(applied.executionRecords[0]).toEqual(
        expect.objectContaining({
          bundleId: pending.id,
          acceptedScope: 'partial',
          summary: 'Write cells in Sheet1!C3',
        }),
      )
      expect(applied.reviewQueueItems).toEqual([])
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }).reviewQueueItems,
      ).toEqual([])
    } finally {
      releaseApply()
      await service.close()
    }
  })

  it('recovers durable review item state after the service restarts', async () => {
    let durableThreadState: WorkbookAgentThreadStateRecord | null = {
      documentId: 'doc-1',
      threadId: 'thr-test',
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
      entries: [],
      reviewQueueItems: [
        createReviewQueueItem({
          id: 'bundle-legacy-1',
          documentId: 'doc-1',
          threadId: 'thr-test',
          turnId: 'turn-1',
          goalText: 'Update workbook from assistant request',
          summary: 'Write cells in Sheet1!B2',
          scope: 'sheet',
          riskClass: 'medium',
          baseRevision: 1,
          createdAtUnixMs: 100,
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
          commands: [
            {
              kind: 'writeRange',
              sheetName: 'Sheet1',
              startAddress: 'B2',
              values: [[42]],
            },
          ],
          affectedRanges: [
            {
              sheetName: 'Sheet1',
              startAddress: 'B2',
              endAddress: 'B2',
              role: 'target',
            },
          ],
          estimatedAffectedCells: 1,
          sharedReview: null,
        }),
      ],
      updatedAtUnixMs: 100,
    }
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const inspectWorkbook = async <T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>): Promise<T> => {
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: 1,
          calculatedRevision: 1,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-10T00:00:00.000Z',
        }),
        headRevision: 1,
        calculatedRevision: 1,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    }
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 7,
      preview: createPreviewSummary({
        cellDiffs: [
          {
            sheetName: 'Sheet1',
            address: 'B2',
            beforeInput: null,
            beforeFormula: null,
            afterInput: 42,
            afterFormula: null,
            changeKinds: ['input'],
          },
        ],
      }),
    }))
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const zeroSync = createZeroSyncStub({
      applyAgentCommandBundle,
      appendWorkbookAgentRun,
      inspectWorkbook,
      async loadWorkbookAgentThreadState() {
        return durableThreadState ? structuredClone(durableThreadState) : null
      },
      async saveWorkbookAgentThreadState(record) {
        durableThreadState = structuredClone(record)
      },
    })
    const fakeCodexB = new FakeCodexTransport()
    const serviceB = createWorkbookAgentService(zeroSync, {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodexB,
    })

    try {
      const resumed = await serviceB.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-test',
        },
      })

      expect(resumed.context).toEqual(
        expect.objectContaining({
          selection: expect.objectContaining({
            sheetName: 'Sheet1',
            address: 'A1',
          }),
        }),
      )
      expect(resumed.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: expect.stringContaining('Write cells in Sheet1!B2'),
          }),
        ]),
      )
      expect(resumed.reviewQueueItems).toEqual([])
      expect(resumed.executionRecords).toEqual([
        expect.objectContaining({
          summary: 'Write cells in Sheet1!B2',
          appliedRevision: 7,
          appliedBy: 'auto',
        }),
      ])
      expect(applyAgentCommandBundle).toHaveBeenCalledWith(
        'doc-1',
        expect.objectContaining({
          summary: 'Write cells in Sheet1!B2',
        }),
        expect.objectContaining({
          ranges: [
            expect.objectContaining({
              sheetName: 'Sheet1',
              startAddress: 'B2',
            }),
          ],
        }),
        expect.objectContaining({
          userID: 'alex@example.com',
        }),
      )
      expect(appendWorkbookAgentRun).toHaveBeenCalledTimes(1)
      const savedThreadState = durableThreadState as WorkbookAgentThreadStateRecord | null
      if (!savedThreadState) {
        throw new Error('Expected durable thread state to be saved')
      }
      expect(savedThreadState.reviewQueueItems).toEqual([])
    } finally {
      await serviceB.close()
    }
  })

  it('replays private-thread execution history directly into the workbook', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const inspectWorkbook = async <T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>): Promise<T> => {
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: 4,
          calculatedRevision: 4,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-10T00:00:00.000Z',
        }),
        headRevision: 4,
        calculatedRevision: 4,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    }
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 11,
      preview: createPreviewSummary({
        cellDiffs: [
          {
            sheetName: 'Sheet1',
            address: 'B2',
            beforeInput: null,
            beforeFormula: null,
            afterInput: 42,
            afterFormula: null,
            changeKinds: ['input'],
          },
        ],
      }),
    }))
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const historicalRecord = {
      id: 'run-1',
      bundleId: 'bundle-1',
      documentId: 'doc-1',
      threadId: 'thr-test',
      turnId: 'turn-1',
      actorUserId: 'alex@example.com',
      goalText: 'Write 42',
      planText: 'Write 42 into the active cell',
      summary: 'Write cells in Sheet1!B2',
      scope: 'selection' as const,
      riskClass: 'low' as const,
      acceptedScope: 'full' as const,
      appliedBy: 'user' as const,
      baseRevision: 3,
      appliedRevision: 4,
      createdAtUnixMs: 100,
      appliedAtUnixMs: 101,
      context: {
        selection: {
          sheetName: 'Sheet1',
          address: 'B2',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      commands: [
        {
          kind: 'writeRange' as const,
          sheetName: 'Sheet1',
          startAddress: 'B2',
          values: [[42]],
        },
      ],
      preview: null,
    }
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        inspectWorkbook,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-test',
            actorUserId: 'alex@example.com',
            scope: 'private',
            executionPolicy: 'autoApplyAll',
            context: null,
            entries: [],
            reviewQueueItems: [],
            updatedAtUnixMs: 100,
          }
        },
        async listWorkbookAgentThreadRuns() {
          return [historicalRecord]
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-test',
        },
      })

      const replayed = await service.replayExecutionRecord({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        recordId: 'run-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })

      expect(replayed.reviewQueueItems).toEqual([])
      expect(replayed.executionRecords[0]).toEqual(
        expect.objectContaining({
          summary: 'Write cells in Sheet1!B2',
          appliedBy: 'auto',
          appliedRevision: 11,
        }),
      )
      expect(replayed.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Applied automatically workbook change set at revision r11: Write cells in Sheet1!B2',
          }),
        ]),
      )
      expect(applyAgentCommandBundle).toHaveBeenCalledTimes(1)
      expect(appendWorkbookAgentRun).toHaveBeenCalledTimes(1)
    } finally {
      await service.close()
    }
  })

  it('falls back to durable thread state when live thread resume is unavailable', async () => {
    const fakeCodex = new FakeCodexTransport()
    fakeCodex.resumeError = new Error('codex resume unavailable')
    const durableThreadState: WorkbookAgentThreadStateRecord = {
      documentId: 'doc-1',
      threadId: 'thr-durable-only',
      actorUserId: 'alex@example.com',
      scope: 'shared',
      executionPolicy: 'ownerReview',
      context: {
        selection: {
          sheetName: 'Sheet2',
          address: 'C7',
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
          id: 'system-1',
          kind: 'system',
          turnId: null,
          text: 'Recovered durable shared thread history.',
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
    }
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return structuredClone(durableThreadState)
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-durable-only',
        },
      })

      expect(fakeCodex.lastThreadResumeInput).toEqual(
        expect.objectContaining({
          threadId: 'thr-durable-only',
        }),
      )
      expect(snapshot.threadId).toBe('thr-durable-only')
      expect(snapshot.scope).toBe('shared')
      expect(snapshot.status).toBe('failed')
      expect(snapshot.lastError).toContain('codex resume unavailable')
      expect(snapshot.context).toEqual(
        expect.objectContaining({
          selection: expect.objectContaining({
            sheetName: 'Sheet2',
            address: 'C7',
          }),
        }),
      )
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Recovered durable shared thread history.',
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('recovers a resumed idle thread with stale in-progress turn history so new work is not blocked', async () => {
    const fakeCodex = new FakeCodexTransport()
    fakeCodex.resumedThread = {
      id: 'thr-stale-turn',
      preview: '',
      status: { type: 'idle' },
      turns: [
        {
          id: 'turn-stale',
          status: 'inProgress',
          items: [],
          error: null,
        },
      ],
    }
    const savedThreadStates: WorkbookAgentThreadStateRecord[] = []
    const zeroSync = createZeroSyncStub({
      async loadWorkbookAgentThreadState() {
        return {
          documentId: 'doc-1',
          threadId: 'thr-stale-turn',
          actorUserId: 'alex@example.com',
          scope: 'private',
          executionPolicy: 'autoApplyAll',
          context: null,
          entries: [],
          reviewQueueItems: [],
          updatedAtUnixMs: 100,
        }
      },
      async saveWorkbookAgentThreadState(record) {
        savedThreadStates.push(structuredClone(record))
      },
    })
    const service = createWorkbookAgentService(zeroSync, {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
    })

    try {
      const recovered = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-stale-turn',
        },
      })

      expect(recovered.status).toBe('idle')
      expect(recovered.lastError).toBeNull()
      expect(recovered.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Cleared stale in-progress turn turn-stale after Codex resumed the thread as idle.',
          }),
        ]),
      )
      expect(savedThreadStates.at(-1)?.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            text: 'Cleared stale in-progress turn turn-stale after Codex resumed the thread as idle.',
          }),
        ]),
      )

      const next = await service.startTurn({
        documentId: 'doc-1',
        threadId: recovered.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Continue after restart',
        },
      })

      expect(next.status).toBe('inProgress')
      expect(next.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'user',
            text: 'Continue after restart',
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })
})
