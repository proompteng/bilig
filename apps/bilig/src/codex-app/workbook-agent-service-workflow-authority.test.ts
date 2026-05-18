import {
  createWorkbookAgentCommandBundle,
  toWorkbookAgentReviewQueueItem,
  type CodexServerNotification,
  type CodexThread,
  type CodexTurn,
} from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it, vi } from 'vitest'
import type { ZeroSyncService } from '../zero/service.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { WorkbookAgentThreadStateRecord } from '../zero/workbook-chat-thread-store.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { createWorkbookAgentService } from './workbook-agent-service.js'

class FakeCodexTransport implements CodexAppServerTransport {
  private threadCounter = 0
  private turnCounter = 0
  readonly interruptedThreadIds: string[] = []

  async ensureReady() {
    return {
      userAgent: 'fake',
      codexHome: '/tmp/fake-codex',
      platformFamily: 'unix',
      platformOs: 'macos',
    }
  }

  subscribe(_listener: (notification: CodexServerNotification) => void): () => void {
    return () => {}
  }

  async threadStart(): Promise<CodexThread> {
    this.threadCounter += 1
    return {
      id: this.threadCounter === 1 ? 'thr-test' : `thr-test-${String(this.threadCounter)}`,
      preview: '',
      turns: [],
    }
  }

  async threadResume(input: { threadId: string }): Promise<CodexThread> {
    return {
      id: input.threadId,
      preview: '',
      turns: [],
    }
  }

  async turnStart(): Promise<CodexTurn> {
    this.turnCounter += 1
    return {
      id: `turn-${String(this.turnCounter)}`,
      status: 'inProgress',
      items: [],
      error: null,
    }
  }

  async turnInterrupt(threadId: string): Promise<void> {
    this.interruptedThreadIds.push(threadId)
  }

  async close(): Promise<void> {}
}

function createContext(sheetName: string, address: string): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName,
      address,
    },
    viewport: {
      rowStart: 0,
      rowEnd: 20,
      colStart: 0,
      colEnd: 10,
    },
  }
}

function createThreadState(overrides: Partial<WorkbookAgentThreadStateRecord> = {}): WorkbookAgentThreadStateRecord {
  return {
    documentId: 'doc-1',
    threadId: 'thr-shared',
    actorUserId: 'alex@example.com',
    scope: 'shared',
    executionPolicy: 'ownerReview',
    context: createContext('Sheet1', 'A1'),
    entries: [],
    reviewQueueItems: [],
    updatedAtUnixMs: 100,
    ...overrides,
  }
}

function createReviewQueueItem() {
  return toWorkbookAgentReviewQueueItem({
    bundle: createWorkbookAgentCommandBundle({
      documentId: 'doc-1',
      threadId: 'thr-shared',
      turnId: 'turn-review',
      goalText: 'Review pending edits',
      baseRevision: 4,
      context: null,
      commands: [
        {
          kind: 'writeRange',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          values: [[42]],
        },
      ],
      now: 100,
    }),
    reviewMode: 'manual',
  })
}

function createZeroSyncStub(overrides: Partial<ZeroSyncService> = {}): ZeroSyncService {
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
    async inspectWorkbook<T>(_documentId: string, _task: (runtime: WorkbookRuntime) => T | Promise<T>): Promise<T> {
      throw new Error('not used')
    },
    async applyServerMutator() {},
    async applyAgentCommandBundle() {
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
    ...overrides,
  }
}

describe('workbook agent workflow authority', () => {
  it('does not mutate session context when workflow start is rejected by preflight', async () => {
    const saveWorkbookAgentThreadState = vi.fn(async (_record: WorkbookAgentThreadStateRecord) => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return createThreadState({
            reviewQueueItems: [createReviewQueueItem()],
          })
        },
        async saveWorkbookAgentThreadState(record) {
          await saveWorkbookAgentThreadState(record)
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
          threadId: 'thr-shared',
        },
      })
      saveWorkbookAgentThreadState.mockClear()

      await expect(
        service.startWorkflow({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            workflowTemplate: 'createSheet',
            name: 'Summary',
            context: createContext('Sheet2', 'B2'),
          },
        }),
      ).rejects.toThrow('Finish the current workbook review item before starting another mutating workflow.')

      expect(saveWorkbookAgentThreadState).not.toHaveBeenCalled()
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }).context,
      ).toEqual(createContext('Sheet1', 'A1'))
    } finally {
      await service.close()
    }
  })

  it('prevents shared collaborators from cancelling workflows they did not start or own', async () => {
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
    const saveWorkbookAgentThreadState = vi.fn(async (_record: WorkbookAgentThreadStateRecord) => undefined)
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>): Promise<T> {
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
        async loadWorkbookAgentThreadState() {
          return createThreadState()
        },
        async saveWorkbookAgentThreadState(record) {
          await saveWorkbookAgentThreadState(record)
        },
        upsertWorkbookWorkflowRun,
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
          threadId: 'thr-shared',
        },
      })

      const running = await service.startWorkflow({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          workflowTemplate: 'summarizeWorkbook',
        },
      })
      const runId = running.workflowRuns[0]?.runId
      if (!runId) {
        throw new Error('Expected running workflow id')
      }
      saveWorkbookAgentThreadState.mockClear()
      upsertWorkbookWorkflowRun.mockClear()

      await expect(
        service.cancelWorkflow({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          runId,
          session: {
            userID: 'casey@example.com',
            roles: ['editor'],
          },
        }),
      ).rejects.toThrow('Only the workflow starter or shared thread owner can cancel this workflow.')

      expect(upsertWorkbookWorkflowRun).not.toHaveBeenCalled()
      expect(saveWorkbookAgentThreadState).not.toHaveBeenCalled()
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'casey@example.com',
            roles: ['editor'],
          },
        }).workflowRuns[0],
      ).toEqual(expect.objectContaining({ status: 'running' }))
    } finally {
      releaseInspection()
      await service.close()
    }
  })

  it('prevents shared collaborators from interrupting active turns they did not start or own', async () => {
    const codexTransport = new FakeCodexTransport()
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return createThreadState()
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => codexTransport,
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
          threadId: 'thr-shared',
        },
      })
      const running = await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Audit the revenue sheet',
          context: createContext('Sheet1', 'A1'),
        },
      })

      await expect(
        service.interruptTurn({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'casey@example.com',
            roles: ['editor'],
          },
        }),
      ).rejects.toThrow('Only the active turn author or shared thread owner can stop this turn.')

      expect(codexTransport.interruptedThreadIds).toEqual([])
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'casey@example.com',
            roles: ['editor'],
          },
        }),
      ).toEqual(
        expect.objectContaining({
          activeTurnActorUserId: 'alex@example.com',
          activeTurnId: running.activeTurnId,
          status: 'inProgress',
        }),
      )
    } finally {
      await service.close()
    }
  })

  it('lets a shared turn author interrupt their own active turn', async () => {
    const codexTransport = new FakeCodexTransport()
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return createThreadState({
            actorUserId: 'alex@example.com',
          })
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => codexTransport,
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
          threadId: 'thr-shared',
        },
      })
      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Audit the revenue sheet',
          context: createContext('Sheet1', 'A1'),
        },
      })

      const interrupted = await service.interruptTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
      })

      expect(codexTransport.interruptedThreadIds).toEqual([snapshot.threadId])
      expect(interrupted.activeTurnActorUserId).toBe('casey@example.com')
    } finally {
      await service.close()
    }
  })
})
