import type { CodexDynamicToolCallResult, CodexServerNotification, CodexTurn, WorkbookAgentCommandBundle } from '@bilig/agent-api'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { describe, expect, it, vi } from 'vitest'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import { applyWorkbookAgentCommandBundleWithUndoCapture } from '../zero/workbook-agent-apply.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { createWorkbookAgentService } from './workbook-agent-service.js'

class FakeCodexTransport implements CodexAppServerTransport {
  private readonly listeners = new Set<(notification: CodexServerNotification) => void>()
  private turnCounter = 0
  currentOptions: CodexAppServerClientOptions | null = null

  async ensureReady() {
    return {
      userAgent: 'fake',
      codexHome: '/tmp/fake-codex',
      platformFamily: 'unix' as const,
      platformOs: 'macos' as const,
    }
  }

  subscribe(listener: (notification: CodexServerNotification) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  async threadStart() {
    return {
      id: 'thr-rendered-freshness',
      preview: '',
      turns: [],
    }
  }

  async threadResume(input: { threadId: string }) {
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

  async turnInterrupt() {}

  async close() {}
}

function isUnknownRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readDynamicToolJson(result: CodexDynamicToolCallResult | undefined): Record<string, unknown> {
  const output = result?.contentItems.find((item) => item.type === 'inputText')
  if (!output || !('text' in output)) {
    throw new Error('Expected dynamic tool inputText output')
  }
  const parsed = JSON.parse(output.text) as unknown
  if (!isUnknownRecord(parsed)) {
    throw new Error('Expected dynamic tool JSON object output')
  }
  return parsed
}

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-rendered-freshness',
    replicaId: 'server:test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  return engine
}

function createPreviewSummary(input: { readonly sheetName: string; readonly address: string; readonly afterInput: string }) {
  return {
    ranges: [
      {
        sheetName: input.sheetName,
        startAddress: input.address,
        endAddress: input.address,
        role: 'target' as const,
      },
    ],
    structuralChanges: [],
    cellDiffs: [
      {
        sheetName: input.sheetName,
        address: input.address,
        beforeInput: null,
        beforeFormula: null,
        afterInput: input.afterInput,
        afterFormula: null,
        changeKinds: ['input' as const],
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
  }
}

function createZeroSyncService(engine: SpreadsheetEngine, input: { readonly revisionRef: { current: number } }): ZeroSyncService {
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
    async inspectWorkbook<T>(documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>) {
      const runtime: WorkbookRuntime = {
        documentId,
        engine,
        projection: buildWorkbookSourceProjectionFromEngine(documentId, engine, {
          revision: input.revisionRef.current,
          calculatedRevision: input.revisionRef.current,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-30T12:00:00.000Z',
        }),
        headRevision: input.revisionRef.current,
        calculatedRevision: input.revisionRef.current,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    },
    async applyServerMutator() {
      throw new Error('not used')
    },
    async applyAgentCommandBundle(_documentId: string, bundle: WorkbookAgentCommandBundle) {
      applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
      input.revisionRef.current = 3
      return {
        revision: 3,
        preview: createPreviewSummary({
          sheetName: 'Sheet1',
          address: 'K14',
          afterInput: 'tool-check-freshness',
        }),
      }
    },
    async listWorkbookChanges() {
      return [
        {
          revision: 3,
          actorUserId: 'alex@example.com',
          clientMutationId: null,
          eventKind: 'applyAgentCommandBundle' as const,
          summary: 'Write cells in Sheet1!K14',
          sheetId: null,
          sheetName: 'Sheet1',
          anchorAddress: 'K14',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'K14',
            endAddress: 'K14',
          },
          rangeInvalid: false,
          undoBundle: {
            kind: 'engineOps' as const,
            ops: [],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAtUnixMs: 3,
        },
      ]
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
      return input.revisionRef.current
    },
    async loadAuthoritativeEvents() {
      throw new Error('not used')
    },
  }
}

function renderedContext(input: { readonly value: string | null; readonly capturedRevision: number }): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: 'Sheet1',
      address: 'K14',
      range: {
        startAddress: 'K14',
        endAddress: 'K14',
      },
    },
    viewport: {
      rowStart: 12,
      rowEnd: 16,
      colStart: 8,
      colEnd: 11,
    },
    rendered: {
      capturedAtUnixMs: Date.now(),
      capturedRevision: input.capturedRevision,
      batchId: input.capturedRevision,
      selection: {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'K14',
          endAddress: 'K14',
        },
        rowCount: 1,
        columnCount: 1,
        cellCount: 1,
        truncated: false,
        rows: [
          [
            {
              address: 'K14',
              input: input.value,
              value:
                input.value === null
                  ? { tag: ValueTag.Empty }
                  : {
                      tag: ValueTag.String,
                      value: input.value,
                    },
              formula: null,
              displayFormat: input.value,
              styleId: null,
              numberFormatId: null,
              style: null,
            },
          ],
        ],
      },
      visibleRange: null,
    },
  }
}

describe('workbook agent rendered freshness', () => {
  it('fails closed for rendered tools when the authoritative head revision cannot be loaded', async () => {
    const engine = await createEngine()
    const revisionRef = { current: 5 }
    const zeroSyncService = createZeroSyncService(engine, { revisionRef })
    zeroSyncService.getWorkbookHeadRevision = vi.fn(async () => {
      throw new Error('head revision unavailable')
    })
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const service = createWorkbookAgentService(zeroSyncService, {
      codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
        capturedOptions.current = options
        return fakeCodex
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
          threadId: 'thr-rendered-fail-closed',
          context: renderedContext({
            value: 'stale-visible-value',
            capturedRevision: 1,
          }),
        },
      })
      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Read rendered state',
        },
      })

      await expect(
        capturedOptions.current?.handleDynamicToolCall({
          threadId: snapshot.threadId,
          turnId: 'turn-1',
          callId: 'call-read-rendered',
          tool: 'read_rendered_selection',
          arguments: {},
        }),
      ).rejects.toThrow('head revision unavailable')
    } finally {
      await service.close()
    }
  })

  it('waits for exact rendered target proof after a mutation, not just the captured revision', async () => {
    vi.useRealTimers()
    const engine = await createEngine()
    const revisionRef = { current: 2 }
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const service = createWorkbookAgentService(createZeroSyncService(engine, { revisionRef }), {
      codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
        capturedOptions.current = options
        fakeCodex.currentOptions = options
        return fakeCodex
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
          threadId: 'thr-rendered-freshness',
          context: renderedContext({
            value: null,
            capturedRevision: 2,
          }),
        },
      })
      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Write and verify exact rendered state',
          context: renderedContext({
            value: null,
            capturedRevision: 2,
          }),
        },
      })

      const callPromise = capturedOptions.current?.handleDynamicToolCall({
        threadId: snapshot.threadId,
        turnId: 'turn-1',
        callId: 'call-write-k14',
        tool: 'write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'K14',
          values: [['tool-check-freshness']],
        },
      })
      if (!callPromise) {
        throw new Error('Expected dynamic tool handler to be captured')
      }

      setTimeout(() => {
        void service.updateContext({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            context: renderedContext({
              value: null,
              capturedRevision: 3,
            }),
          },
        })
      }, 20)
      setTimeout(() => {
        void service.updateContext({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            context: renderedContext({
              value: 'tool-check-freshness',
              capturedRevision: 3,
            }),
          },
        })
      }, 120)

      const payload = readDynamicToolJson(await callPromise)
      expect(payload['status']).toBe('applied')
      expect(payload['mutationReceipt']).toEqual(
        expect.objectContaining({
          renderedReadback: expect.objectContaining({
            matched: true,
            stale: false,
            capturedRevision: 3,
            incompleteReason: null,
          }),
        }),
      )
    } finally {
      await service.close()
    }
  })
})
