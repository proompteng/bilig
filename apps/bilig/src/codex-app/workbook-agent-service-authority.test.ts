import type { WorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { WorkbookAgentWorkflowRun } from '@bilig/contracts'
import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it, vi } from 'vitest'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { WorkbookAgentThreadStateRecord } from '../zero/workbook-chat-thread-store.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { createWorkbookAgentService } from './workbook-agent-service.js'

import {
  FakeCodexTransport,
  createReviewQueueItem,
  createVisibleSceneProof,
  createZeroSyncStub,
  readDynamicToolJson,
  startWorkbookAgentTestTurn,
  waitForWorkflowStatus,
} from './workbook-agent-service.test-helpers.js'

describe('workbook agent service workflow authority handoff', () => {
  it('does not apply mutating workflow commands after cancellation during authority handoff', async () => {
    const fakeCodex = new FakeCodexTransport()
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
    const getWorkbookHeadRevision = vi.fn(async () => {
      resolveHeadRevisionRequested()
      await headRevisionBlocked
      return 7
    })
    const upsertWorkflowRunStatuses: WorkbookAgentWorkflowRun['status'][] = []
    const upsertWorkbookWorkflowRun = vi.fn(async (_documentId: string, run: WorkbookAgentWorkflowRun) => {
      upsertWorkflowRunStatuses.push(run.status)
    })
    const applyAgentCommandBundle = vi.fn(async (_documentId, _bundle, preview) => ({
      revision: 8,
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
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
          return await task(runtime)
        },
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        getWorkbookHeadRevision,
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

      const runningSnapshot = await service.startWorkflow({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          workflowTemplate: 'createSheet',
          name: 'Forecast',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      await headRevisionRequested
      const runId = runningSnapshot.workflowRuns[0]?.runId
      if (!runId) {
        throw new Error('Expected running workflow id')
      }

      await service.cancelWorkflow({
        documentId: 'doc-1',
        threadId: 'thr-test',
        runId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      releaseHeadRevision()
      await new Promise((resolve) => setTimeout(resolve, 0))

      const finalSnapshot = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      expect(finalSnapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'createSheet',
          status: 'cancelled',
          summary: 'Cancelled workflow: Create Sheet',
          artifact: null,
        }),
      )
      expect(applyAgentCommandBundle).not.toHaveBeenCalled()
      expect(appendWorkbookAgentRun).not.toHaveBeenCalled()
      expect(upsertWorkflowRunStatuses).toEqual(['running', 'cancelled'])
      expect(finalSnapshot.entries).not.toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Create Sheet',
          }),
        ]),
      )
    } finally {
      releaseHeadRevision()
      await service.close()
    }
  })

  it('does not apply mutating workflow commands after cancellation during authoritative preview', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    let releasePreview!: () => void
    const previewBlocked = new Promise<void>((resolve) => {
      releasePreview = resolve
    })
    let resolvePreviewStarted!: () => void
    const previewStarted = new Promise<void>((resolve) => {
      resolvePreviewStarted = resolve
    })
    const upsertWorkflowRunStatuses: WorkbookAgentWorkflowRun['status'][] = []
    const upsertWorkbookWorkflowRun = vi.fn(async (_documentId: string, run: WorkbookAgentWorkflowRun) => {
      upsertWorkflowRunStatuses.push(run.status)
    })
    const applyAgentCommandBundle = vi.fn(async (_documentId, _bundle, preview) => ({
      revision: 8,
      preview,
    }))
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>) {
          resolvePreviewStarted()
          await previewBlocked
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
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        async getWorkbookHeadRevision() {
          return 7
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

      const runningSnapshot = await service.startWorkflow({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          workflowTemplate: 'createSheet',
          name: 'Forecast',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      await previewStarted
      const runId = runningSnapshot.workflowRuns[0]?.runId
      if (!runId) {
        throw new Error('Expected running workflow id')
      }

      await service.cancelWorkflow({
        documentId: 'doc-1',
        threadId: 'thr-test',
        runId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      releasePreview()
      await new Promise((resolve) => setTimeout(resolve, 0))

      const finalSnapshot = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      expect(finalSnapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'createSheet',
          status: 'cancelled',
          summary: 'Cancelled workflow: Create Sheet',
          artifact: null,
        }),
      )
      expect(applyAgentCommandBundle).not.toHaveBeenCalled()
      expect(appendWorkbookAgentRun).not.toHaveBeenCalled()
      expect(upsertWorkflowRunStatuses).toEqual(['running', 'cancelled'])
      expect(finalSnapshot.entries).not.toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Create Sheet',
          }),
        ]),
      )
    } finally {
      releasePreview()
      await service.close()
    }
  })

  it('applies rename-sheet workflows immediately in private threads', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Revenue')
    const applyAgentCommandBundle = vi.fn(async (_documentId, bundle: WorkbookAgentCommandBundle, preview) => {
      for (const command of bundle.commands) {
        if (command.kind === 'renameSheet') {
          engine.renameSheet(command.currentName, command.nextName)
        }
      }
      return {
        revision: 9,
        preview,
      }
    })
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
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
          return await task(runtime)
        },
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
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
        body: {
          context: {
            selection: {
              sheetName: 'Revenue',
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

      const runningSnapshot = await service.startWorkflow({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          workflowTemplate: 'renameCurrentSheet',
          name: 'Forecast',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'renameCurrentSheet',
          status: 'completed',
          summary:
            'Workflow mutation verification incomplete at revision r9: No target cell range was available for authoritative readback.',
          mutationExecuted: true,
          verificationComplete: false,
          mutationStatus: 'verification_incomplete',
          mutationReceipt: expect.objectContaining({
            status: 'verification_incomplete',
            toolName: 'workflow:renameCurrentSheet',
          }),
          artifact: expect.objectContaining({
            title: 'Rename Sheet Preview',
          }),
        }),
      )
      expect(snapshot.reviewQueueItems).toEqual([])
      expect(applyAgentCommandBundle).toHaveBeenCalledWith(
        'doc-1',
        expect.objectContaining({
          commands: [
            expect.objectContaining({
              kind: 'renameSheet',
              currentName: 'Revenue',
              nextName: 'Forecast',
            }),
          ],
        }),
        expect.anything(),
        expect.objectContaining({
          userID: 'alex@example.com',
        }),
      )
      expect(appendWorkbookAgentRun).toHaveBeenCalledTimes(1)
      expect(snapshot.executionRecords).toEqual([
        expect.objectContaining({
          summary: 'Rename sheet Revenue to Forecast',
          appliedRevision: 9,
          appliedBy: 'auto',
        }),
      ])
      expect(snapshot.context?.selection.sheetName).toBe('Forecast')
      expect(JSON.stringify(snapshot.context)).not.toContain('Revenue')
    } finally {
      await service.close()
    }
  })

  it('stages outlier-highlight change sets from durable workflows', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Revenue')
    engine.setCellValue('Revenue', 'A1', 'Region')
    engine.setCellValue('Revenue', 'B1', 'Revenue')
    engine.setCellValue('Revenue', 'A2', 'West')
    engine.setCellValue('Revenue', 'B2', 100)
    engine.setCellValue('Revenue', 'A3', 'East')
    engine.setCellValue('Revenue', 'B3', 105)
    engine.setCellValue('Revenue', 'A4', 'North')
    engine.setCellValue('Revenue', 'B4', 98)
    engine.setCellValue('Revenue', 'A5', 'South')
    engine.setCellValue('Revenue', 'B5', 102)
    engine.setCellValue('Revenue', 'A6', 'Enterprise')
    engine.setCellValue('Revenue', 'B6', 450)
    const getWorkbookHeadRevision = vi.fn(async () => 7)
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined)
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
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
          return await task(runtime)
        },
        getWorkbookHeadRevision,
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
        body: {
          context: {
            selection: {
              sheetName: 'Revenue',
              address: 'A1',
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

      const runningSnapshot = await service.startWorkflow({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          workflowTemplate: 'highlightCurrentSheetOutliers',
          sheetName: 'Revenue',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(getWorkbookHeadRevision).toHaveBeenCalledWith('doc-1')
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'highlightCurrentSheetOutliers',
          title: 'Highlight Current Sheet Outliers',
          status: 'completed',
          artifact: expect.objectContaining({
            title: 'Current Sheet Outlier Highlights',
            text: expect.stringContaining('## Highlighted Numeric Outliers'),
          }),
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'stage-outlier-highlights',
              status: 'completed',
            }),
          ]),
        }),
      )
      expect(snapshot.reviewQueueItems).toEqual([])
      expect(snapshot.executionRecords).toEqual([
        expect.objectContaining({
          appliedBy: 'auto',
          commands: [
            expect.objectContaining({
              kind: 'formatRange',
              range: expect.objectContaining({
                sheetName: 'Revenue',
                startAddress: 'B6',
                endAddress: 'B6',
              }),
              patch: expect.objectContaining({
                fill: expect.objectContaining({
                  backgroundColor: '#FEF3C7',
                }),
              }),
            }),
          ],
        }),
      ])
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: expect.stringContaining('Applied automatically workbook change set'),
          }),
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Highlight Current Sheet Outliers',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Revenue',
                startAddress: 'B6',
                endAddress: 'B6',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('allows Codex dynamic tools to start durable workflows inside the active thread', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined)
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
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options
          return fakeCodex
        },
      },
    )

    try {
      const created = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })
      await startWorkbookAgentTestTurn(service, {
        threadId: created.threadId,
      })

      const result = await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-test',
        turnId: 'turn-1',
        callId: 'call-start-workflow',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'summarizeWorkbook',
        },
      })

      expect(result?.success).toBe(true)
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(snapshot.workflowRuns).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            workflowTemplate: 'summarizeWorkbook',
            status: 'completed',
            steps: expect.arrayContaining([
              expect.objectContaining({
                stepId: 'inspect-workbook',
                status: 'completed',
              }),
            ]),
            artifact: expect.objectContaining({
              title: 'Workbook Summary',
            }),
          }),
        ]),
      )
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Started workflow: Summarize Workbook',
          }),
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Summarize Workbook',
          }),
        ]),
      )
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
    } finally {
      await service.close()
    }
  })

  it('applies direct structural tool commands immediately inside the active thread', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
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
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options
          return fakeCodex
        },
      },
    )

    try {
      const created = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {},
      })
      await startWorkbookAgentTestTurn(service, {
        threadId: created.threadId,
      })

      const result = await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-test',
        turnId: 'turn-1',
        callId: 'call-create-sheet',
        tool: 'bilig_create_sheet',
        arguments: {
          name: 'Prepaid Expenses',
        },
      })

      expect(result?.success).toBe(true)
      const output = result?.contentItems.find((item) => item.type === 'inputText')
      expect(output?.type).toBe('inputText')
      const text = output && 'text' in output ? output.text : ''
      expect(text).toContain('"applied": false')
      expect(text).toContain('"mutationExecuted": true')
      expect(text).toContain('"verificationComplete": false')
      expect(text).toContain('"status": "verification_incomplete"')
      expect(text).toContain('"staged": false')
      expect(text).toContain('"reviewQueued": false')
      expect(text).toContain('"queuedForTurnApply": false')
      expect(text).toContain('"revision": 7')

      const snapshot = service.getSnapshot({
        documentId: 'doc-1',
        threadId: 'thr-test',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      expect(snapshot.reviewQueueItems).toEqual([])
      expect(snapshot.executionRecords).toEqual([
        expect.objectContaining({
          summary: 'Create sheet Prepaid Expenses',
          appliedRevision: 7,
          appliedBy: 'auto',
        }),
      ])
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: expect.stringContaining('Applied automatically workbook change set'),
          }),
        ]),
      )
      expect(applyAgentCommandBundle).toHaveBeenCalledWith(
        'doc-1',
        expect.objectContaining({
          commands: [
            {
              kind: 'createSheet',
              name: 'Prepaid Expenses',
            },
          ],
        }),
        expect.objectContaining({
          structuralChanges: ['Create sheet Prepaid Expenses'],
        }),
        expect.objectContaining({
          userID: 'alex@example.com',
        }),
      )
      expect(appendWorkbookAgentRun).toHaveBeenCalledTimes(1)
    } finally {
      await service.close()
    }
  })

  it('uses the request turn actor and context for shared-thread workflow starts', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet2', 'C7', 99)
    let durableThreadState: WorkbookAgentThreadStateRecord | null = {
      documentId: 'doc-1',
      threadId: 'thr-shared',
      actorUserId: 'alex@example.com',
      scope: 'shared',
      executionPolicy: 'ownerReview',
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
      reviewQueueItems: [],
      updatedAtUnixMs: 100,
    }
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined)
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
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
          return await task(runtime)
        },
        async loadWorkbookAgentThreadState() {
          return durableThreadState ? structuredClone(durableThreadState) : null
        },
        async saveWorkbookAgentThreadState(record: WorkbookAgentThreadStateRecord) {
          durableThreadState = structuredClone(record)
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options
          return fakeCodex
        },
      },
    )

    try {
      await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      const caseySnapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      await service.startTurn({
        documentId: 'doc-1',
        threadId: caseySnapshot.threadId,
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Summarize my current sheet',
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
        },
      })

      await service.updateContext({
        documentId: 'doc-1',
        threadId: caseySnapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
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
        },
      })

      const result = await capturedOptions.current?.handleDynamicToolCall({
        threadId: 'thr-shared',
        turnId: 'turn-1',
        callId: 'call-start-workflow',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'summarizeCurrentSheet',
        },
      })

      expect(result?.success).toBe(true)
      const snapshot = await waitForWorkflowStatus(service, caseySnapshot.threadId, 'casey@example.com', 'completed')
      expect(snapshot.workflowRuns).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            workflowTemplate: 'summarizeCurrentSheet',
            startedByUserId: 'casey@example.com',
            summary: 'Summarized Sheet2 with 1 populated cell and 0 tables.',
            artifact: expect.objectContaining({
              text: expect.stringContaining('Sheet: Sheet2'),
            }),
          }),
        ]),
      )
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      const startedByUserIds = upsertWorkbookWorkflowRun.mock.calls.map(
        (call) =>
          (
            call.at(1) as
              | {
                  startedByUserId?: string
                }
              | undefined
          )?.startedByUserId ?? null,
      )
      expect(startedByUserIds).toEqual(['casey@example.com', 'casey@example.com'])
    } finally {
      await service.close()
    }
  })

  it('refreshes rendered browser context for the active turn when the same user posts a context update', async () => {
    const fakeCodex = new FakeCodexTransport()
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null }
    const service = createWorkbookAgentService(createZeroSyncStub(), {
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
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'A1',
              range: {
                startAddress: 'A1',
                endAddress: 'A1',
              },
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

      await service.startTurn({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Verify rendered state',
        },
      })

      await service.updateContext({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'A1',
              range: {
                startAddress: 'A1',
                endAddress: 'A1',
              },
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
            rendered: {
              capturedAtUnixMs: 100,
              capturedRevision: 3,
              batchId: 1,
              visibleSceneProof: createVisibleSceneProof(3),
              selection: {
                range: {
                  sheetName: 'Sheet1',
                  startAddress: 'A1',
                  endAddress: 'A1',
                },
                rowCount: 1,
                columnCount: 1,
                cellCount: 1,
                truncated: false,
                rows: [
                  [
                    {
                      address: 'A1',
                      input: 42,
                      value: { tag: 1, value: 42 },
                      formula: null,
                      displayFormat: null,
                      styleId: null,
                      numberFormatId: null,
                      style: null,
                    },
                  ],
                ],
              },
              visibleRange: null,
            },
          },
        },
      })

      const result = await capturedOptions.current?.handleDynamicToolCall({
        threadId: snapshot.threadId,
        turnId: 'turn-1',
        callId: 'call-rendered-selection',
        tool: 'read_rendered_selection',
        arguments: {},
      })
      const payload = readDynamicToolJson(result)

      expect(payload).toEqual(
        expect.objectContaining({
          renderedReadback: expect.objectContaining({
            available: true,
            matched: true,
            stale: false,
            capturedRevision: 3,
            capturedBatchId: 1,
            incompleteReason: null,
          }),
        }),
      )
    } finally {
      await service.close()
    }
  })

  it('rejects starting a second workflow while one is still running', async () => {
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
        async upsertWorkbookWorkflowRun(_documentId, run) {
          if (run.status === 'running') {
            resolveRunningPersisted()
          }
        },
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

      const firstWorkflow = service.startWorkflow({
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

      await expect(
        service.startWorkflow({
          documentId: 'doc-1',
          threadId: 'thr-test',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            workflowTemplate: 'describeRecentChanges',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_WORKFLOW_ALREADY_RUNNING',
        statusCode: 409,
      })

      releaseInspection()
      await firstWorkflow
      await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
    } finally {
      releaseInspection()
      await service.close()
    }
  })

  it('rejects workflow starts when the expected active turn no longer owns the session', async () => {
    const fakeCodex = new FakeCodexTransport()
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
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
        body: {},
      })
      await startWorkbookAgentTestTurn(service, {
        threadId: snapshot.threadId,
        prompt: 'Start a workflow',
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

      const staleWorkflowStart = {
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        expectedActiveTurnId: 'turn-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          workflowTemplate: 'summarizeWorkbook',
        },
      }
      await expect(service.startWorkflow(staleWorkflowStart)).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_STALE_TOOL_CALL',
        statusCode: 409,
        retryable: false,
      })
      expect(upsertWorkbookWorkflowRun).not.toHaveBeenCalled()
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }).workflowRuns,
      ).toEqual([])
    } finally {
      await service.close()
    }
  })

  it('clears legacy private review items before starting a new workflow', async () => {
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-existing',
            actorUserId: 'alex@example.com',
            scope: 'private',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-existing',
                documentId: 'doc-1',
                threadId: 'thr-existing',
                turnId: 'turn-1',
                goalText: 'Normalize the imported range',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A20',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 20,
                sharedReview: null,
              }),
            ],
            updatedAtUnixMs: 100,
          }
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
          threadId: 'thr-existing',
        },
      })

      expect(snapshot.reviewQueueItems).toEqual([])
      expect(snapshot.lastError).toBe('Private workbook threads no longer keep queued review items. Replay the request to apply it again.')

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
      expect(running.workflowRuns[0]?.status).toBe('running')
    } finally {
      await service.close()
    }
  })
})
