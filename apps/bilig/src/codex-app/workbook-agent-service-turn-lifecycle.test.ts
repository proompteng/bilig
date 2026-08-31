import { isWorkbookAgentCommandBundle, isWorkbookAgentExecutionRecord, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it, vi } from 'vitest'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { createWorkbookAgentService } from './workbook-agent-service.js'

import {
  FakeCodexTransport,
  createDurableRunningWorkflowRun,
  createPreviewSummary,
  createZeroSyncStub,
  getPrimaryReviewBundle,
  startWorkbookAgentTestTurn,
  waitForWorkflowStatus,
} from './workbook-agent-service.test-helpers.js'

describe('workbook agent service turn lifecycle and stale ownership', () => {
  it('recovers stale durable running workflows on bootstrap and allows new workflow starts', async () => {
    const fakeCodex = new FakeCodexTransport()
    const staleWorkflowRun = createDurableRunningWorkflowRun()
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-existing',
            actorUserId: 'alex@example.com',
            scope: 'private',
            executionPolicy: 'autoApplyAll',
            context: null,
            entries: [],
            reviewQueueItems: [],
            updatedAtUnixMs: 100,
          }
        },
        async listWorkbookThreadWorkflowRuns() {
          return [staleWorkflowRun]
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
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
          threadId: 'thr-existing',
        },
      })

      expect(snapshot.workflowRuns).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            runId: 'workflow-existing',
            status: 'failed',
            errorMessage: 'Workflow interrupted because the workbook assistant restarted before it could finish.',
          }),
        ]),
      )
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            text: 'Marked stale running workflow as failed after assistant restart: Summarize Workbook',
          }),
        ]),
      )
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledWith(
        'doc-1',
        expect.objectContaining({
          runId: 'workflow-existing',
          status: 'failed',
        }),
      )

      const running = await service.startWorkflow({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          workflowTemplate: 'createSheet',
          name: 'Summary',
        },
      })

      expect(running.workflowRuns).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            workflowTemplate: 'createSheet',
            status: 'running',
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('cancels a running durable workflow without letting late completion overwrite it', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)

    let releaseInspection!: () => void
    const inspectBarrier = new Promise<void>((resolve) => {
      releaseInspection = () => {
        resolve()
      }
    })
    let resolveRunningPersisted!: () => void
    const runningPersisted = new Promise<void>((resolve) => {
      resolveRunningPersisted = () => {
        resolve()
      }
    })
    const upsertWorkbookWorkflowRun = vi.fn(async (_documentId: string, run) => {
      if (run.status === 'running') {
        resolveRunningPersisted()
      }
    })

    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>) {
          await inspectBarrier
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
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      },
    )

    try {
      await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      const workflowPromise = service.startWorkflow({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          workflowTemplate: 'summarizeWorkbook',
        },
      })

      await runningPersisted
      const queuedSnapshot = await workflowPromise
      expect(queuedSnapshot.workflowRuns[0]?.status).toBe('running')
      const runningSnapshot = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      const runningRunId = runningSnapshot.workflowRuns[0]?.runId
      if (!runningRunId) {
        throw new Error('Expected running workflow run id')
      }

      const cancelledSnapshot = await service.cancelWorkflow({
        documentId: 'doc-1',
        threadId: 'thr-test',
        runId: runningRunId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })

      expect(cancelledSnapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'summarizeWorkbook',
          status: 'cancelled',
          summary: 'Cancelled workflow: Summarize Workbook',
          errorMessage: 'Cancelled by alex@example.com.',
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'inspect-workbook',
              status: 'cancelled',
            }),
          ]),
        }),
      )

      releaseInspection()
      const finalSnapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'cancelled')

      expect(finalSnapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'summarizeWorkbook',
          status: 'cancelled',
          summary: 'Cancelled workflow: Summarize Workbook',
          artifact: null,
        }),
      )
      expect(finalSnapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Started workflow: Summarize Workbook',
          }),
          expect.objectContaining({
            kind: 'system',
            text: 'Cancelled workflow: Summarize Workbook',
          }),
        ]),
      )
      expect(finalSnapshot.entries).not.toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Summarize Workbook',
          }),
        ]),
      )
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(upsertWorkbookWorkflowRun.mock.calls.map(([, run]) => run.status)).toEqual(['running', 'cancelled'])
    } finally {
      releaseInspection()
      await service.close()
    }
  })

  it('persists terminal workflow cancellation during service shutdown', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)

    let releaseInspection!: () => void
    const inspectBarrier = new Promise<void>((resolve) => {
      releaseInspection = () => {
        resolve()
      }
    })
    let resolveRunningPersisted!: () => void
    const runningPersisted = new Promise<void>((resolve) => {
      resolveRunningPersisted = () => {
        resolve()
      }
    })
    let resolveShutdownCancelPersisted!: () => void
    const shutdownCancelPersisted = new Promise<void>((resolve) => {
      resolveShutdownCancelPersisted = () => {
        resolve()
      }
    })
    const upsertWorkbookWorkflowRun = vi.fn(async (_documentId: string, run) => {
      if (run.status === 'running') {
        resolveRunningPersisted()
      }
      if (run.status === 'cancelled') {
        resolveShutdownCancelPersisted()
      }
    })
    const saveWorkbookAgentThreadState = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>) {
          await inspectBarrier
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
        },
        upsertWorkbookWorkflowRun,
        saveWorkbookAgentThreadState,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      },
    )

    await service.createSession({
      documentId: 'doc-1',
      session: {
        userID: 'alex@example.com',
        roles: ['editor'],
      },
      body: {},
    })

    await service.startWorkflow({
      documentId: 'doc-1',
      threadId: 'thr-test',
      session: {
        userID: 'alex@example.com',
        roles: ['editor'],
      },
      body: {
        workflowTemplate: 'summarizeWorkbook',
      },
    })

    await runningPersisted
    const closePromise = service.close()
    await shutdownCancelPersisted

    expect(upsertWorkbookWorkflowRun.mock.calls.map(([, run]) => run.status)).toEqual(['running', 'cancelled'])
    const cancelledRun = upsertWorkbookWorkflowRun.mock.calls.at(-1)?.[1]
    expect(cancelledRun).toEqual(
      expect.objectContaining({
        status: 'cancelled',
        summary: 'Cancelled workflow: Summarize Workbook',
        errorMessage: 'Cancelled because the workbook assistant service shut down.',
        artifact: null,
      }),
    )
    expect(cancelledRun?.steps).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          stepId: 'inspect-workbook',
          status: 'cancelled',
        }),
      ]),
    )
    expect(saveWorkbookAgentThreadState).toHaveBeenLastCalledWith(
      expect.objectContaining({
        entries: expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Cancelled workflow during service shutdown: Summarize Workbook',
          }),
        ]),
      }),
    )

    releaseInspection()
    await closePromise
  })

  it('bounds workflow shutdown drain when a workflow ignores abort', async () => {
    const fakeCodex = new FakeCodexTransport()
    let resolveRunningPersisted!: () => void
    const runningPersisted = new Promise<void>((resolve) => {
      resolveRunningPersisted = () => {
        resolve()
      }
    })
    let resolveShutdownCancelPersisted!: () => void
    const shutdownCancelPersisted = new Promise<void>((resolve) => {
      resolveShutdownCancelPersisted = () => {
        resolve()
      }
    })
    const upsertWorkbookWorkflowRun = vi.fn(async (_documentId: string, run) => {
      if (run.status === 'running') {
        resolveRunningPersisted()
      }
      if (run.status === 'cancelled') {
        resolveShutdownCancelPersisted()
      }
    })
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(): Promise<T> {
          return await new Promise<T>(() => {})
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
        workflowShutdownDrainTimeoutMs: 1,
      },
    )

    await service.createSession({
      documentId: 'doc-1',
      session: {
        userID: 'alex@example.com',
        roles: ['editor'],
      },
      body: {},
    })

    await service.startWorkflow({
      documentId: 'doc-1',
      threadId: 'thr-test',
      session: {
        userID: 'alex@example.com',
        roles: ['editor'],
      },
      body: {
        workflowTemplate: 'summarizeWorkbook',
      },
    })

    await runningPersisted
    const closePromise = service.close()
    await shutdownCancelPersisted
    await closePromise

    expect(upsertWorkbookWorkflowRun.mock.calls.map(([, run]) => run.status)).toEqual(['running', 'cancelled'])
    expect(fakeCodex.closeCount).toBeGreaterThan(0)
  })

  it('uses nested app-server error messages instead of the generic fallback', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
    })

    try {
      await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      fakeCodex.emit({
        method: 'error',
        params: {
          error: {
            code: -32602,
            message: 'thread/start.dynamicTools requires experimentalApi capability',
          },
        },
      })

      const snapshot = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })

      expect(snapshot.status).toBe('failed')
      expect(snapshot.lastError).toBe('thread/start.dynamicTools requires experimentalApi capability')
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'thread/start.dynamicTools requires experimentalApi capability',
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('releases active turn state after a runtime error so the user can recover with a new turn', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxActiveTurnsPerUser: 1,
      maxActiveTurnsPerDocument: 1,
    })

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Inspect Sheet1 before the runtime fails',
        },
      })

      fakeCodex.emit({
        method: 'error',
        params: {
          error: {
            code: -32000,
            message: 'Codex app-server exited unexpectedly',
          },
        },
      })

      const failed = service.getSnapshot({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      expect(failed.status).toBe('failed')
      expect(failed.lastError).toBe('Codex app-server exited unexpectedly')
      expect(failed.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            turnId: 'turn-1',
            text: 'Codex app-server exited unexpectedly',
          }),
        ]),
      )

      const recovered = await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Recover after runtime failure',
        },
      })

      expect(recovered.status).toBe('inProgress')
      expect(recovered.lastError).toBeNull()
      expect(recovered.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'user',
            text: 'Recover after runtime failure',
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('does not let a stale completed-turn notification clear a newer active turn', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxActiveTurnsPerUser: 1,
      maxActiveTurnsPerDocument: 1,
    })

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'First turn',
        },
      })
      fakeCodex.emit({
        method: 'turn/started',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-2',
            status: 'inProgress',
            items: [],
            error: null,
          },
        },
      })

      fakeCodex.emit({
        method: 'turn/completed',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-1',
            status: 'completed',
            items: [],
            error: null,
          },
        },
      })

      await vi.waitFor(() => {
        const stillActive = service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        })
        expect(stillActive.status).toBe('inProgress')
        expect(stillActive.activeTurnId).toBe('turn-2')
      })

      await expect(
        service.startTurn({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            prompt: 'Should still be blocked by turn-2',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_TURN_ALREADY_RUNNING',
      })
    } finally {
      await service.close()
    }
  })

  it('does not let a stale started-turn notification reclaim a newer active turn', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxActiveTurnsPerUser: 1,
      maxActiveTurnsPerDocument: 1,
    })

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'First turn',
        },
      })
      fakeCodex.emit({
        method: 'turn/started',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-2',
            status: 'inProgress',
            items: [],
            error: null,
          },
        },
      })
      await vi.waitFor(() => {
        expect(
          service.getSnapshot({
            documentId: 'doc-1',
            threadId: snapshot.threadId,
            session: {
              userID: 'alex@example.com',
              roles: ['editor'],
            },
          }).activeTurnId,
        ).toBe('turn-2')
      })
      fakeCodex.emit({
        method: 'turn/started',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-1',
            status: 'inProgress',
            items: [],
            error: null,
          },
        },
      })

      await vi.waitFor(() => {
        const stillActive = service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        })
        expect(stillActive.status).toBe('inProgress')
        expect(stillActive.activeTurnId).toBe('turn-2')
      })
    } finally {
      await service.close()
    }
  })

  it('does not let a stale failed turn poison a newer active turn result', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxActiveTurnsPerUser: 1,
      maxActiveTurnsPerDocument: 1,
    })

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'First turn',
        },
      })
      fakeCodex.emit({
        method: 'turn/started',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-2',
            status: 'inProgress',
            items: [],
            error: null,
          },
        },
      })
      fakeCodex.emit({
        method: 'turn/completed',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-1',
            status: 'failed',
            items: [],
            error: {
              message: 'Old turn failed after retry already started',
            },
          },
        },
      })

      const stillActive = service.getSnapshot({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      expect(stillActive.status).toBe('inProgress')
      expect(stillActive.activeTurnId).toBe('turn-2')
      expect(stillActive.lastError).toBeNull()

      fakeCodex.emit({
        method: 'turn/completed',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-2',
            status: 'completed',
            items: [],
            error: null,
          },
        },
      })

      await vi.waitFor(() => {
        const recovered = service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        })
        expect(recovered.status).toBe('idle')
        expect(recovered.activeTurnId).toBeNull()
        expect(recovered.lastError).toBeNull()
      })
    } finally {
      await service.close()
    }
  })

  it('rejects dynamic tool calls when no turn currently owns the live session', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 7,
      preview: createPreviewSummary(),
    }))
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
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
        body: {},
      })
      const handler = capturedOptions.current?.handleDynamicToolCall
      if (!handler) {
        throw new Error('Expected dynamic tool handler to be captured')
      }

      await expect(
        handler({
          threadId: snapshot.threadId,
          turnId: 'turn-1',
          callId: 'call-idle-tool',
          tool: 'bilig_write_range',
          arguments: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            values: [[42]],
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_STALE_TOOL_CALL',
        statusCode: 409,
        retryable: false,
      })
      expect(applyAgentCommandBundle).not.toHaveBeenCalled()
    } finally {
      await service.close()
    }
  })

  it('rejects stale dynamic tool calls for turns that no longer own the live session', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
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
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
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
        body: {},
      })
      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'First turn',
        },
      })
      fakeCodex.emit({
        method: 'turn/started',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-2',
            status: 'inProgress',
            items: [],
            error: null,
          },
        },
      })
      const handler = capturedOptions.current?.handleDynamicToolCall
      if (!handler) {
        throw new Error('Expected dynamic tool handler to be captured')
      }

      await expect(
        handler({
          threadId: snapshot.threadId,
          turnId: 'turn-1',
          callId: 'call-stale-tool',
          tool: 'bilig_write_range',
          arguments: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            values: [[42]],
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_STALE_TOOL_CALL',
        statusCode: 409,
        retryable: false,
      })
      expect(applyAgentCommandBundle).not.toHaveBeenCalled()
    } finally {
      await service.close()
    }
  })

  it('rejects dynamic tool commands when the turn loses ownership while resolving workbook authority', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    let releaseHeadRevision!: () => void
    const headRevisionBlocked = new Promise<void>((resolve) => {
      releaseHeadRevision = resolve
    })
    let resolveHeadRevisionRequested!: () => void
    const headRevisionRequested = new Promise<void>((resolve) => {
      resolveHeadRevisionRequested = resolve
    })
    const applyAgentCommandBundle = vi.fn(async (_documentId, _bundle, preview) => ({
      revision: 7,
      preview,
    }))
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>) {
          const runtime: WorkbookRuntime = {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-11T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
          return await task(runtime)
        },
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        async getWorkbookHeadRevision() {
          resolveHeadRevisionRequested()
          await headRevisionBlocked
          return 7
        },
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
        body: {},
      })
      await startWorkbookAgentTestTurn(service, {
        threadId: snapshot.threadId,
        prompt: 'Create a stale sheet',
      })
      const handler = capturedOptions.current?.handleDynamicToolCall
      if (!handler) {
        throw new Error('Expected dynamic tool handler to be captured')
      }

      const toolCallPromise = handler({
        threadId: snapshot.threadId,
        turnId: 'turn-1',
        callId: 'call-stale-create-sheet',
        tool: 'bilig_create_sheet',
        arguments: {
          name: 'Stale Sheet',
        },
      })
      await headRevisionRequested
      fakeCodex.emit({
        method: 'turn/completed',
        params: {
          threadId: snapshot.threadId,
          turn: {
            id: 'turn-1',
            status: 'completed',
            items: [],
            error: null,
          },
        },
      })
      await vi.waitFor(() => {
        expect(
          service.getSnapshot({
            documentId: 'doc-1',
            threadId: snapshot.threadId,
            session: {
              userID: 'alex@example.com',
              roles: ['editor'],
            },
          }).activeTurnId,
        ).toBeNull()
      })
      releaseHeadRevision()

      const result = await toolCallPromise
      expect(result.success).toBe(false)
      expect(result.contentItems).toEqual([
        expect.objectContaining({
          type: 'inputText',
          text: expect.stringContaining('Rejecting workbook tool call because the assistant turn is no longer active.'),
        }),
      ])
      expect(applyAgentCommandBundle).not.toHaveBeenCalled()
      expect(appendWorkbookAgentRun).not.toHaveBeenCalled()
      const finalSnapshot = service.getSnapshot({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      expect(finalSnapshot.executionRecords).toEqual([])
      expect(finalSnapshot.reviewQueueItems).toEqual([])
    } finally {
      releaseHeadRevision()
      await service.close()
    }
  })

  it('falls back to a stable runtime message when the app-server emits an empty error', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
    })

    try {
      await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      fakeCodex.emit({
        method: 'error',
        params: {},
      })

      const snapshot = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })

      expect(snapshot.status).toBe('failed')
      expect(snapshot.lastError).toBe('Workbook assistant runtime failed. Retry in a moment.')
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Workbook assistant runtime failed. Retry in a moment.',
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('persists the authoritative preview returned by apply and not the caller payload', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const applyAgentCommandBundle = vi.fn(async (_documentId: string, _bundle: WorkbookAgentCommandBundle, _preview: unknown) => ({
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

      const applied = await service.applyReviewItem({
        documentId: 'doc-1',
        threadId: 'thr-test',
        reviewItemId: pending.id,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        appliedBy: 'user',
      })

      const record = applied.executionRecords[0]
      if (!isWorkbookAgentExecutionRecord(record)) {
        throw new Error('Expected an execution record after apply')
      }
      expect(applyAgentCommandBundle).toHaveBeenCalled()
      expect(applyAgentCommandBundle).toHaveBeenCalledWith(
        'doc-1',
        expect.objectContaining({
          id: pending.id,
        }),
        expect.objectContaining({
          ranges: [
            {
              sheetName: 'Sheet1',
              startAddress: 'B2',
              endAddress: 'B2',
              role: 'target',
            },
          ],
        }),
        expect.objectContaining({
          userID: 'alex@example.com',
        }),
      )
      expect(record.preview).toEqual(
        expect.objectContaining({
          cellDiffs: [
            expect.objectContaining({
              sheetName: 'Sheet1',
              address: 'B2',
            }),
          ],
        }),
      )
      expect(appendWorkbookAgentRun).toHaveBeenCalledWith(
        expect.objectContaining({
          preview: expect.objectContaining({
            effectSummary: expect.objectContaining({
              displayedCellDiffCount: 1,
              inputChangeCount: 1,
            }),
          }),
        }),
      )
      expect(applied.entries).toContainEqual(
        expect.objectContaining({
          kind: 'system',
          text: 'Applied workbook change set at revision r7: Write cells in Sheet1!B2',
          citations: [
            expect.objectContaining({
              kind: 'range',
              sheetName: 'Sheet1',
              startAddress: 'B2',
              endAddress: 'B2',
            }),
            expect.objectContaining({
              kind: 'revision',
              revision: 7,
            }),
          ],
        }),
      )
    } finally {
      await service.close()
    }
  })
})
