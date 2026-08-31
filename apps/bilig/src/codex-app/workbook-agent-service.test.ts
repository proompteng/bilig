import type { CodexTurn } from '@bilig/agent-api'
import type { WorkbookAgentThreadSummary } from '@bilig/contracts'
import { describe, expect, it, vi } from 'vitest'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import type { WorkbookAgentFeatureFlags } from './workbook-agent-feature-flags.js'
import { createWorkbookAgentService } from './workbook-agent-service.js'

import {
  FakeCodexTransport,
  createPreviewSummary,
  createRenderedContextForServiceTest,
  createReviewQueueItem,
  createZeroSyncStub,
  startWorkbookAgentTestTurn,
} from './workbook-agent-service.test-helpers.js'

describe('workbook agent service transport timeline and access policy', () => {
  it('stays disabled when the top-level workbook-agent gate is closed', async () => {
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      featureFlags: {
        enabled: false,
      },
    })

    expect(service.enabled).toBe(false)
    await service.close()
  })

  it('boots the Codex app-server transport with local workbook skills', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: {
      current: CodexAppServerClientOptions | null
    } = { current: null }
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
        capturedOptions.current = options
        return fakeCodex
      },
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

      expect(capturedOptions.current?.args).toEqual(
        expect.arrayContaining([
          'app-server',
          '--strict-config',
          'approval_policy="on-request"',
          'sandbox_mode="read-only"',
          'sandbox_workspace_write.network_access=false',
          'web_search="disabled"',
          'mcp_servers={}',
        ]),
      )
      expect(fakeCodex.lastThreadStartInput).toMatchObject({
        model: 'gpt-5.5',
        approvalPolicy: 'on-request',
        sandbox: 'read-only',
        config: {
          approval_policy: 'on-request',
          sandbox_mode: 'read-only',
          sandbox_workspace_write: {
            network_access: false,
          },
          web_search: 'disabled',
          mcp_servers: {},
        },
      })
      expect(fakeCodex.lastThreadStartInput?.dynamicTools.map((tool) => tool.name)).toEqual(
        expect.arrayContaining([
          'read_selection',
          'read_visible_range',
          'start_workflow',
          'inspect_cell',
          'find_formula_issues',
          'search_workbook',
          'trace_dependencies',
          'read_range',
          'write_range',
        ]),
      )
      expect(fakeCodex.lastThreadStartInput?.dynamicTools.every((tool) => /^[a-zA-Z0-9_-]+$/.test(tool.name))).toBe(true)
      expect(fakeCodex.lastThreadStartInput?.baseInstructions).toContain('Help with the active workbook only.')
      expect(fakeCodex.lastThreadStartInput?.baseInstructions).toContain(
        'Use only the available workbook tools for workbook context; say when external context is unavailable.',
      )
      expect(fakeCodex.lastThreadStartInput?.baseInstructions).not.toContain('Tools:')
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).toContain(
        'Use the workflow tool only for built-in multi-step or durable tasks.',
      )
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).toContain(
        'Use direct structural sheet tools for one-step sheet edits that should happen immediately.',
      )
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).toContain(
        'Apply workbook changes directly when the session policy allows it.',
      )
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).toContain(
        'Workbook state must come from workbook tools; do not assume external search or network context is available.',
      )
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).not.toContain('review and apply it from the panel')
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).not.toContain('stage one coherent change set per turn')
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).not.toContain('summarizeWorkbook')
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).not.toContain('Do not use non-workbook tools')
    } finally {
      await service.close()
    }
  })

  it('hides a non-visible thread before attempting to resume it', async () => {
    const fakeCodex = new FakeCodexTransport()
    const loadWorkbookAgentThreadState = vi.fn(async () => null)
    const codexClientFactory = vi.fn((): CodexAppServerTransport => fakeCodex)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        loadWorkbookAgentThreadState,
      }),
      {
        codexClientFactory,
      },
    )

    try {
      await expect(
        service.createSession({
          documentId: 'doc-1',
          session: {
            userID: 'mallory@example.com',
            roles: ['editor'],
          },
          body: {
            threadId: 'thr-private',
            scope: 'private',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_THREAD_NOT_FOUND',
        statusCode: 404,
        retryable: false,
      })

      expect(loadWorkbookAgentThreadState).toHaveBeenCalledWith('doc-1', 'mallory@example.com', 'thr-private')
      expect(codexClientFactory).not.toHaveBeenCalled()
      expect(fakeCodex.lastThreadResumeInput).toBeNull()
    } finally {
      await service.close()
    }
  })

  it('loads durable authorization before resume and falls back when authorized resume fails', async () => {
    const fakeCodex = new FakeCodexTransport()
    const events: string[] = []
    const loadWorkbookAgentThreadState = vi.fn(async () => {
      events.push('load')
      return {
        documentId: 'doc-1',
        threadId: 'thr-durable',
        actorUserId: 'alex@example.com',
        scope: 'private' as const,
        executionPolicy: 'ownerReview' as const,
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
        entries: [],
        reviewQueueItems: [],
        updatedAtUnixMs: 100,
      }
    })
    vi.spyOn(fakeCodex, 'threadResume').mockImplementation(async () => {
      events.push('resume')
      throw new Error('codex resume unavailable')
    })
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        loadWorkbookAgentThreadState,
      }),
      {
        codexClientFactory: (): CodexAppServerTransport => fakeCodex,
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
          threadId: 'thr-durable',
        },
      })

      expect(events).toEqual(['load', 'resume'])
      expect(loadWorkbookAgentThreadState).toHaveBeenCalledTimes(1)
      expect(snapshot.threadId).toBe('thr-durable')
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
    } finally {
      await service.close()
    }
  })

  it('streams assistant updates into the session timeline', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxCodexClients: 1,
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

      expect(snapshot.threadId).toBe('thr-test')
      expect(snapshot.reviewQueueItems).toEqual([])
      expect(snapshot.executionRecords).toEqual([])

      const events: unknown[] = []
      const unsubscribe = service.subscribe(snapshot.threadId, (event) => {
        events.push(event)
      })

      const inProgress = await service.startTurn({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Summarize Sheet1',
        },
      })

      expect(inProgress.status).toBe('inProgress')
      expect(inProgress.entries.some((entry) => entry.kind === 'user')).toBe(true)

      fakeCodex.emit({
        method: 'item/agentMessage/delta',
        params: {
          threadId: 'thr-test',
          turnId: 'turn-1',
          itemId: 'msg-1',
          delta: 'Checking Sheet1',
        },
      })
      fakeCodex.emit({
        method: 'item/completed',
        params: {
          threadId: 'thr-test',
          turnId: 'turn-1',
          item: {
            type: 'agentMessage',
            id: 'msg-1',
            text: 'Checking Sheet1',
            phase: null,
            memoryCitation: null,
          },
        },
      })
      fakeCodex.emit({
        method: 'turn/completed',
        params: {
          threadId: 'thr-test',
          turn: {
            id: 'turn-1',
            status: 'completed',
            items: [],
            error: null,
          },
        },
      })

      await vi.waitFor(() => {
        const finalSnapshot = service.getSnapshot({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        })
        expect(finalSnapshot.status).toBe('idle')
        expect(finalSnapshot.entries).toEqual(
          expect.arrayContaining([
            expect.objectContaining({
              id: 'msg-1',
              kind: 'assistant',
              text: 'Checking Sheet1',
            }),
          ]),
        )
      })
      await vi.waitFor(() => {
        expect(events).toContainEqual({
          type: 'entryTextDelta',
          turnId: 'turn-1',
          entryKind: 'assistant',
          itemId: 'msg-1',
          delta: 'Checking Sheet1',
        })
      })

      unsubscribe()
    } finally {
      await service.close()
    }
  })

  it('does not broadcast duplicate workbook context snapshots for rendered proof metadata churn', async () => {
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (): CodexAppServerTransport => new FakeCodexTransport(),
      maxCodexClients: 1,
    })

    try {
      const baseContext = createRenderedContextForServiceTest({
        capturedRevision: 1,
        stringId: 1,
        value: 'same visible text',
      })
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          context: baseContext,
        },
      })
      const events: unknown[] = []
      const unsubscribe = service.subscribe(snapshot.threadId, (event) => {
        events.push(event)
      })

      await service.updateContext({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          context: createRenderedContextForServiceTest({
            capturedRevision: 9,
            stringId: 99,
            value: 'same visible text',
          }),
        },
      })

      expect(events).toEqual([])
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }).context,
      ).toEqual(baseContext)

      await service.updateContext({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          context: createRenderedContextForServiceTest({
            capturedRevision: 10,
            stringId: 100,
            value: 'changed visible text',
          }),
        },
      })

      expect(events).toHaveLength(1)
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }).context?.rendered?.visibleRange?.rows[0]?.[0]?.input,
      ).toBe('changed visible text')
      unsubscribe()
    } finally {
      await service.close()
    }
  })

  it('streams reasoning updates into first-class reasoning timeline entries', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxCodexClients: 1,
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

      const events: unknown[] = []
      const unsubscribe = service.subscribe(snapshot.threadId, (event) => {
        events.push(event)
      })

      await service.startTurn({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Check the version issues',
        },
      })

      fakeCodex.emit({
        method: 'item/reasoning/textDelta',
        params: {
          threadId: 'thr-test',
          turnId: 'turn-1',
          itemId: 'reasoning-1',
          delta: 'Examining version issues',
        },
      })

      const streamingSnapshot = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })

      expect(streamingSnapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            id: 'reasoning-1',
            kind: 'reasoning',
            text: 'Examining version issues',
          }),
        ]),
      )

      fakeCodex.emit({
        method: 'item/completed',
        params: {
          threadId: 'thr-test',
          turnId: 'turn-1',
          item: {
            type: 'reasoning',
            id: 'reasoning-1',
            summary: [
              {
                type: 'summary_text',
                text: 'Examining version issues',
              },
              {
                type: 'summary_text',
                text: 'Confirming whether staged changes must be cleared first.',
              },
            ],
          },
        },
      })
      fakeCodex.emit({
        method: 'turn/completed',
        params: {
          threadId: 'thr-test',
          turn: {
            id: 'turn-1',
            status: 'completed',
            items: [],
            error: null,
          },
        },
      })

      await vi.waitFor(() => {
        const finalSnapshot = service.getSnapshot({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        })
        expect(finalSnapshot.entries).toEqual(
          expect.arrayContaining([
            expect.objectContaining({
              id: 'reasoning-1',
              kind: 'reasoning',
              text: 'Examining version issues\nConfirming whether staged changes must be cleared first.',
            }),
          ]),
        )
      })
      await vi.waitFor(() => {
        expect(events).toContainEqual({
          type: 'entryTextDelta',
          turnId: 'turn-1',
          entryKind: 'reasoning',
          itemId: 'reasoning-1',
          delta: 'Examining version issues',
        })
      })

      unsubscribe()
    } finally {
      await service.close()
    }
  })

  it('streams command execution output into first-class command timeline entries', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxCodexClients: 1,
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

      const events: unknown[] = []
      const unsubscribe = service.subscribe(snapshot.threadId, (event) => {
        events.push(event)
      })

      await service.startTurn({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Run a command',
        },
      })

      fakeCodex.emit({
        method: 'item/started',
        params: {
          threadId: 'thr-test',
          turnId: 'turn-1',
          item: {
            type: 'commandExecution',
            id: 'cmd-1',
            command: 'printf hi',
            cwd: '/Users/gregkonush/github.com/bilig',
            processId: null,
            status: 'inProgress',
            commandActions: [],
            aggregatedOutput: null,
            exitCode: null,
            durationMs: null,
          },
        },
      })
      fakeCodex.emit({
        method: 'item/commandExecution/outputDelta',
        params: {
          threadId: 'thr-test',
          turnId: 'turn-1',
          itemId: 'cmd-1',
          delta: 'aGkNCg==',
        },
      })
      fakeCodex.emit({
        method: 'item/completed',
        params: {
          threadId: 'thr-test',
          turnId: 'turn-1',
          item: {
            type: 'commandExecution',
            id: 'cmd-1',
            command: 'printf hi',
            cwd: '/Users/gregkonush/github.com/bilig',
            processId: null,
            status: 'completed',
            commandActions: [],
            aggregatedOutput: 'aGkNCg==',
            exitCode: 0,
            durationMs: 12,
          },
        },
      })

      await vi.waitFor(() => {
        const finalSnapshot = service.getSnapshot({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        })
        expect(finalSnapshot.entries).toEqual(
          expect.arrayContaining([
            expect.objectContaining({
              id: 'cmd-1',
              kind: 'tool',
              toolName: 'command_execution',
              toolStatus: 'completed',
              outputText: 'hi\r\n',
              success: true,
            }),
          ]),
        )
        expect(finalSnapshot.entries.some((entry) => entry.text === 'Codex emitted commandExecution.')).toBe(false)
      })
      await vi.waitFor(() => {
        expect(events).toContainEqual({
          type: 'entryToolOutputDelta',
          turnId: 'turn-1',
          itemId: 'cmd-1',
          delta: 'hi\r\n',
        })
      })

      unsubscribe()
    } finally {
      await service.close()
    }
  })

  it('enforces per-user active turn quotas across sessions', async () => {
    const fakeCodex = new FakeCodexTransport()
    fakeCodex.uniqueThreadStart = true
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxActiveTurnsPerUser: 1,
      maxActiveTurnsPerDocument: 8,
    })

    try {
      const sessionA = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })
      const sessionB = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      await service.startTurn({
        documentId: 'doc-1',
        threadId: sessionA.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Inspect Sheet1',
        },
      })

      await expect(
        service.startTurn({
          documentId: 'doc-1',
          threadId: sessionB.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            prompt: 'Inspect Sheet2',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_USER_TURN_QUOTA_EXCEEDED',
        statusCode: 429,
        retryable: true,
      })
    } finally {
      await service.close()
    }
  })

  it('translates Codex pool backpressure into retryable service errors', async () => {
    const firstTurnResolver: { current: ((value: CodexTurn) => void) | null } = { current: null }
    const firstTurn = new Promise<CodexTurn>((resolve) => {
      firstTurnResolver.current = resolve
    })
    const fakeCodex = new FakeCodexTransport()
    fakeCodex.uniqueThreadStart = true
    fakeCodex.nextTurn = firstTurn
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      maxCodexClients: 1,
      maxConcurrentTurnsPerCodexClient: 1,
      maxQueuedTurnsPerCodexClient: 0,
      maxActiveTurnsPerUser: 8,
      maxActiveTurnsPerDocument: 8,
    })

    try {
      const sessionA = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })
      const sessionB = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {},
      })

      const firstStartPromise = service.startTurn({
        documentId: 'doc-1',
        threadId: sessionA.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Run first turn',
        },
      })

      await Promise.resolve()

      await expect(
        service.startTurn({
          documentId: 'doc-1',
          threadId: sessionB.threadId,
          session: {
            userID: 'casey@example.com',
            roles: ['editor'],
          },
          body: {
            prompt: 'Run second turn',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_TURN_BACKPRESSURE',
        statusCode: 429,
        retryable: true,
      })

      if (firstTurnResolver.current) {
        firstTurnResolver.current({
          id: 'turn-1',
          status: 'inProgress',
          items: [],
          error: null,
        })
      }
      fakeCodex.nextTurn = null
      await firstStartPromise
    } finally {
      await service.close()
    }
  })

  it('disables shared threads behind a feature flag', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      featureFlags: {
        sharedThreadsEnabled: false,
      },
    })

    try {
      await expect(
        service.createSession({
          documentId: 'doc-1',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            threadId: 'thr-shared',
            scope: 'shared',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_SHARED_THREADS_DISABLED',
        statusCode: 409,
        retryable: false,
      })
    } finally {
      await service.close()
    }
  })

  it('applies shared-thread feature gates to already-live sessions', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      featureFlags: {
        sharedThreadsEnabled: true,
      },
    })

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          scope: 'shared',
        },
      })
      Object.defineProperty(service, 'featureFlags', {
        value: {
          enabled: true,
          sharedThreadsEnabled: false,
          workflowRunnerEnabled: true,
          autoApplyLowRiskEnabled: true,
          formulaWorkflowFamilyEnabled: true,
          formattingWorkflowFamilyEnabled: true,
          importWorkflowFamilyEnabled: true,
          rollupWorkflowFamilyEnabled: true,
          structuralWorkflowFamilyEnabled: true,
          allowlistedUserIds: [],
          allowlistedDocumentIds: [],
        } satisfies WorkbookAgentFeatureFlags,
      })

      await expect(
        service.createSession({
          documentId: 'doc-1',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            threadId: snapshot.threadId,
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_SHARED_THREADS_DISABLED',
        statusCode: 409,
        retryable: false,
      })
    } finally {
      await service.close()
    }
  })

  it('revalidates live shared thread visibility before granting a new user access', async () => {
    const fakeCodex = new FakeCodexTransport()
    const loadWorkbookAgentThreadState = vi.fn(async (_documentId: string, actorUserId: string, _threadId: string) =>
      actorUserId === 'alex@example.com'
        ? {
            documentId: 'doc-1',
            threadId: 'thr-test',
            actorUserId: 'alex@example.com',
            scope: 'shared' as const,
            executionPolicy: 'ownerReview' as const,
            context: null,
            entries: [],
            reviewQueueItems: [],
            updatedAtUnixMs: 100,
          }
        : null,
    )
    const saveWorkbookAgentThreadState = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        loadWorkbookAgentThreadState,
        saveWorkbookAgentThreadState,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
        featureFlags: {
          sharedThreadsEnabled: true,
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
        },
      })

      await expect(
        service.createSession({
          documentId: 'doc-1',
          session: {
            userID: 'mallory@example.com',
            roles: ['editor'],
          },
          body: {
            threadId: snapshot.threadId,
            context: {
              selection: {
                sheetName: 'Sheet1',
                address: 'Z99',
              },
              viewport: {
                rowStart: 0,
                rowEnd: 20,
                colStart: 0,
                colEnd: 10,
              },
            },
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_THREAD_NOT_FOUND',
        statusCode: 404,
      })

      expect(loadWorkbookAgentThreadState).toHaveBeenCalledWith('doc-1', 'mallory@example.com', snapshot.threadId)
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: { userID: 'alex@example.com', roles: ['editor'] },
        }).context,
      ).not.toEqual(
        expect.objectContaining({
          selection: expect.objectContaining({
            address: 'Z99',
          }),
        }),
      )
      expect(saveWorkbookAgentThreadState).not.toHaveBeenCalledWith(
        expect.objectContaining({
          context: expect.objectContaining({
            selection: expect.objectContaining({
              address: 'Z99',
            }),
          }),
        }),
      )
    } finally {
      await service.close()
    }
  })

  it('prevents collaborators from changing execution policy on live shared threads', async () => {
    const fakeCodex = new FakeCodexTransport()
    const saveWorkbookAgentThreadState = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        saveWorkbookAgentThreadState,
        async loadWorkbookAgentThreadState(_documentId, actorUserId, threadId) {
          return actorUserId === 'casey@example.com'
            ? {
                documentId: 'doc-1',
                threadId,
                actorUserId: 'alex@example.com',
                scope: 'shared' as const,
                executionPolicy: 'ownerReview' as const,
                context: null,
                entries: [],
                reviewQueueItems: [],
                updatedAtUnixMs: 100,
              }
            : null
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
        featureFlags: {
          sharedThreadsEnabled: true,
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
        },
      })
      saveWorkbookAgentThreadState.mockClear()

      await expect(
        service.createSession({
          documentId: 'doc-1',
          session: {
            userID: 'casey@example.com',
            roles: ['editor'],
          },
          body: {
            threadId: snapshot.threadId,
            executionPolicy: 'autoApplySafe',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_SHARED_POLICY_CHANGE_FORBIDDEN',
        statusCode: 409,
        retryable: false,
      })

      expect(saveWorkbookAgentThreadState).not.toHaveBeenCalled()
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }).executionPolicy,
      ).toBe('ownerReview')
    } finally {
      await service.close()
    }
  })

  it('prevents collaborators from changing execution policy while bootstrapping durable shared threads', async () => {
    const fakeCodex = new FakeCodexTransport()
    const saveWorkbookAgentThreadState = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        saveWorkbookAgentThreadState,
        async loadWorkbookAgentThreadState(_documentId, _actorUserId, threadId) {
          return {
            documentId: 'doc-1',
            threadId,
            actorUserId: 'alex@example.com',
            scope: 'shared' as const,
            executionPolicy: 'ownerReview' as const,
            context: null,
            entries: [],
            reviewQueueItems: [],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
        featureFlags: {
          sharedThreadsEnabled: true,
        },
      },
    )

    try {
      await expect(
        service.createSession({
          documentId: 'doc-1',
          session: {
            userID: 'casey@example.com',
            roles: ['editor'],
          },
          body: {
            threadId: 'thr-shared',
            executionPolicy: 'autoApplySafe',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_SHARED_POLICY_CHANGE_FORBIDDEN',
        statusCode: 409,
        retryable: false,
      })

      expect(saveWorkbookAgentThreadState).not.toHaveBeenCalled()
    } finally {
      await service.close()
    }
  })

  it('limits shared threads to the rollout allowlist', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      featureFlags: {
        allowlistedUserIds: ['pat@example.com'],
      },
    })

    try {
      await expect(
        service.createSession({
          documentId: 'doc-1',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            threadId: 'thr-shared',
            scope: 'shared',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_SHARED_THREADS_ROLLOUT_BLOCKED',
        statusCode: 409,
        retryable: false,
      })
    } finally {
      await service.close()
    }
  })

  it('filters inaccessible shared threads from durable thread listings', async () => {
    const summaries: WorkbookAgentThreadSummary[] = [
      {
        threadId: 'thr-private',
        scope: 'private',
        ownerUserId: 'alex@example.com',
        updatedAtUnixMs: 100,
        entryCount: 1,
        reviewQueueItemCount: 0,
        latestEntryText: 'Private thread',
      },
      {
        threadId: 'thr-shared',
        scope: 'shared',
        ownerUserId: 'pat@example.com',
        updatedAtUnixMs: 200,
        entryCount: 2,
        reviewQueueItemCount: 1,
        latestEntryText: 'Shared thread',
      },
    ]
    const zeroSync = createZeroSyncStub({
      async listWorkbookAgentThreadSummaries() {
        return summaries
      },
    })
    const disabledService = createWorkbookAgentService(zeroSync, {
      featureFlags: {
        sharedThreadsEnabled: false,
      },
    })
    const rolloutBlockedService = createWorkbookAgentService(zeroSync, {
      featureFlags: {
        sharedThreadsEnabled: true,
        allowlistedUserIds: ['pat@example.com'],
      },
    })

    try {
      await expect(
        disabledService.listThreads({
          documentId: 'doc-1',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }),
      ).resolves.toEqual([summaries[0]])
      await expect(
        rolloutBlockedService.listThreads({
          documentId: 'doc-1',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }),
      ).resolves.toEqual([summaries[0]])
    } finally {
      await disabledService.close()
      await rolloutBlockedService.close()
    }
  })

  it('preserves shared auto-apply-safe policy and auto-applies low-risk work', async () => {
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
            afterInput: null,
            afterFormula: null,
            changeKinds: ['style'],
          },
        ],
        effectSummary: {
          displayedCellDiffCount: 1,
          truncatedCellDiffs: false,
          inputChangeCount: 0,
          formulaChangeCount: 0,
          styleChangeCount: 1,
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
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared' as const,
            executionPolicy: 'autoApplySafe' as const,
            context: null,
            entries: [],
            reviewQueueItems: [],
            updatedAtUnixMs: 100,
          }
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
        body: {
          threadId: 'thr-shared',
          scope: 'shared',
          executionPolicy: 'autoApplySafe',
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

      expect(snapshot.scope).toBe('shared')
      expect(snapshot.executionPolicy).toBe('autoApplySafe')
      await startWorkbookAgentTestTurn(service, {
        threadId: snapshot.threadId,
      })

      await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-shared',
        turnId: 'turn-1',
        callId: 'call-1',
        tool: 'bilig_format_range',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
          patch: {
            font: {
              bold: true,
            },
          },
        },
      })

      const after = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-shared',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })

      expect(after.reviewQueueItems).toEqual([])
      expect(after.executionRecords).toEqual([
        expect.objectContaining({
          appliedBy: 'auto',
          summary: 'Format Sheet1!B2',
          appliedRevision: 7,
        }),
      ])
      expect(applyAgentCommandBundle).toHaveBeenCalledTimes(1)
      expect(appendWorkbookAgentRun).toHaveBeenCalledTimes(1)
    } finally {
      await service.close()
    }
  })

  it('disables workflow families behind feature flags', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      featureFlags: {
        formulaWorkflowFamilyEnabled: false,
      },
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

      await expect(
        service.startWorkflow({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            workflowTemplate: 'highlightFormulaIssues',
            sheetName: 'Sheet1',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_WORKFLOW_FAMILY_DISABLED',
        statusCode: 409,
        retryable: false,
      })
    } finally {
      await service.close()
    }
  })

  it('disables auto-apply behind a feature flag', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-test',
            actorUserId: 'alex@example.com',
            scope: 'private',
            executionPolicy: 'autoApplySafe',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-auto-1',
                documentId: 'doc-1',
                threadId: 'thr-test',
                turnId: 'turn-1',
                goalText: 'Apply low-risk cleanup',
                summary: 'Write cells in Sheet1!B2',
                scope: 'selection',
                riskClass: 'low',
                baseRevision: 1,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'writeRange',
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    values: [[42]],
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 1,
                sharedReview: null,
              }),
            ],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
        featureFlags: {
          autoApplyLowRiskEnabled: false,
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
          threadId: 'thr-test',
        },
      })

      expect(snapshot.reviewQueueItems).toEqual([])
      expect(snapshot.lastError).toBe('Private workbook threads no longer keep queued review items. Replay the request to apply it again.')
    } finally {
      await service.close()
    }
  })
})
