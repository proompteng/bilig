import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it, vi } from 'vitest'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { createWorkbookAgentService } from './workbook-agent-service.js'

import {
  FakeCodexTransport,
  createReviewQueueItem,
  createZeroSyncStub,
  waitForWorkflowStatus,
} from './workbook-agent-service.test-helpers.js'

describe('workbook agent service durable workflow execution', () => {
  it('limits workflow runner and auto-apply to the rollout allowlist', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-test',
            actorUserId: 'alex@example.com',
            scope: 'private',
            executionPolicy: 'autoApplyAll',
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
          allowlistedUserIds: ['pat@example.com'],
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

      await expect(
        service.startWorkflow({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          body: {
            workflowTemplate: 'summarizeWorkbook',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_WORKFLOW_RUNNER_ROLLOUT_BLOCKED',
        statusCode: 409,
        retryable: false,
      })

      expect(snapshot.reviewQueueItems).toEqual([])
      expect(snapshot.lastError).toBe('Private workbook threads no longer keep queued review items. Replay the request to apply it again.')
    } finally {
      await service.close()
    }
  })

  it('reports observability snapshot counts for rollout and runtime state', async () => {
    const fakeCodex = new FakeCodexTransport()
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
      featureFlags: {
        allowlistedUserIds: ['alex@example.com', 'pat@example.com'],
        allowlistedDocumentIds: ['doc-1'],
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

      const snapshot = service.getObservabilitySnapshot()
      expect(snapshot.enabled).toBe(true)
      expect(snapshot.featureFlags.allowlistedUserCount).toBe(2)
      expect(snapshot.featureFlags.allowlistedDocumentCount).toBe(1)
      expect(snapshot.sessions.sessionCount).toBe(1)
      expect(snapshot.pool.maxClients).toBeGreaterThan(0)
    } finally {
      await service.close()
    }
  })

  it('starts durable read/report workflows and records completed runs', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A1)')
    let inspectWorkbookCallCount = 0
    const inspectWorkbook = async <T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>): Promise<T> => {
      inspectWorkbookCallCount += 1
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
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        inspectWorkbook,
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
          workflowTemplate: 'summarizeWorkbook',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(inspectWorkbookCallCount).toBe(1)
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(snapshot.workflowRuns[0]).toEqual(
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
            kind: 'markdown',
            title: 'Workbook Summary',
          }),
        }),
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
    } finally {
      await service.close()
    }
  })

  it('runs durable formula issue workflows with cited issue reports', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellFormula('Sheet1', 'B1', '1/0')
    engine.setCellFormula('Sheet1', 'C1', 'LEN(A1:A2)')
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
          workflowTemplate: 'findFormulaIssues',
          sheetName: 'Sheet1',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'findFormulaIssues',
          title: 'Find Formula Issues',
          status: 'completed',
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'scan-formula-cells',
              status: 'completed',
            }),
          ]),
          artifact: expect.objectContaining({
            title: 'Formula Issues',
            text: expect.stringContaining('## Formula Issues'),
          }),
        }),
      )
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Started workflow: Find Formula Issues',
          }),
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Find Formula Issues',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Sheet1',
                startAddress: 'B1',
                endAddress: 'B1',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('stages formula-highlight change sets from durable workflows', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellFormula('Sheet1', 'B1', '1/0')
    engine.setCellFormula('Sheet1', 'C1', 'LEN(A1:A2)')
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
          workflowTemplate: 'highlightFormulaIssues',
          sheetName: 'Sheet1',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(getWorkbookHeadRevision).toHaveBeenCalledWith('doc-1')
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'highlightFormulaIssues',
          title: 'Highlight Formula Issues',
          status: 'completed',
          artifact: expect.objectContaining({
            title: 'Formula Issue Highlights',
            text: expect.stringContaining('## Highlighted Formula Issues'),
          }),
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'stage-issue-highlights',
              status: 'completed',
            }),
          ]),
        }),
      )
      expect(snapshot.reviewQueueItems).toEqual([])
      expect(snapshot.executionRecords).toEqual([
        expect.objectContaining({
          appliedBy: 'auto',
          commands: expect.arrayContaining([
            expect.objectContaining({
              kind: 'formatRange',
              range: expect.objectContaining({
                sheetName: 'Sheet1',
                startAddress: 'B1',
                endAddress: 'B1',
              }),
              patch: expect.objectContaining({
                fill: expect.objectContaining({
                  backgroundColor: '#FEE2E2',
                }),
              }),
            }),
          ]),
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
            text: 'Completed workflow: Highlight Formula Issues',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Sheet1',
                startAddress: 'B1',
                endAddress: 'B1',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('stages formula-repair change sets from durable workflows', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellValue('Sheet1', 'A2', 45)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'B2', '1/0')
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
          workflowTemplate: 'repairFormulaIssues',
          sheetName: 'Sheet1',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(getWorkbookHeadRevision).toHaveBeenCalledWith('doc-1')
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'repairFormulaIssues',
          title: 'Repair Formula Issues',
          status: 'completed',
          artifact: expect.objectContaining({
            title: 'Formula Repair Preview',
            text: expect.stringContaining('## Formula Repair Preview'),
          }),
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'stage-formula-repairs',
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
              kind: 'writeRange',
              sheetName: 'Sheet1',
              startAddress: 'B2',
              values: [[{ formula: 'A2*2' }]],
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
            text: 'Completed workflow: Repair Formula Issues',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Sheet1',
                startAddress: 'B2',
                endAddress: 'B2',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('stages header-normalization change sets from durable workflows', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Imports')
    engine.setCellValue('Imports', 'A1', 'order_id')
    engine.setCellValue('Imports', 'B1', ' customer name ')
    engine.setCellValue('Imports', 'C1', 'customer_name')
    engine.setCellValue('Imports', 'A2', 1001)
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
              sheetName: 'Imports',
              address: 'A2',
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
          workflowTemplate: 'normalizeCurrentSheetHeaders',
          sheetName: 'Imports',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(getWorkbookHeadRevision).toHaveBeenCalledWith('doc-1')
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'normalizeCurrentSheetHeaders',
          title: 'Normalize Current Sheet Headers',
          status: 'completed',
          artifact: expect.objectContaining({
            title: 'Header Normalization Preview',
            text: expect.stringContaining('## Header Normalization Preview'),
          }),
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'stage-header-normalization',
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
              kind: 'writeRange',
              sheetName: 'Imports',
              startAddress: 'A1',
              values: [['Order Id', 'Customer Name', 'Customer Name 2']],
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
            text: 'Completed workflow: Normalize Current Sheet Headers',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Imports',
                startAddress: 'A1',
                endAddress: 'C1',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('runs durable current-sheet summary workflows from the active selection context', async () => {
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
    engine.setCellFormula('Revenue', 'B2', 'SUM(B3:B5)')
    engine.setFreezePane('Revenue', 1, 0)
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
              colEnd: 8,
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
          workflowTemplate: 'summarizeCurrentSheet',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'summarizeCurrentSheet',
          title: 'Summarize Current Sheet',
          status: 'completed',
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'inspect-current-sheet',
              status: 'completed',
            }),
          ]),
          artifact: expect.objectContaining({
            title: 'Current Sheet Summary',
            text: expect.stringContaining('Sheet: Revenue'),
          }),
        }),
      )
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Started workflow: Summarize Current Sheet',
          }),
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Summarize Current Sheet',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Revenue',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('runs durable dependency trace workflows from the current selection context', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'C1', 'B1+1')
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
          scope: 'shared',
          executionPolicy: 'ownerReview',
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B1',
            },
            viewport: {
              rowStart: 0,
              rowEnd: 10,
              colStart: 0,
              colEnd: 5,
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
          workflowTemplate: 'traceSelectionDependencies',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'traceSelectionDependencies',
          title: 'Trace Selection Dependencies',
          status: 'completed',
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'trace-links',
              status: 'completed',
            }),
          ]),
          artifact: expect.objectContaining({
            title: 'Dependency Trace',
            text: expect.stringContaining('Root: Sheet1!B1'),
          }),
        }),
      )
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Trace Selection Dependencies',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Sheet1',
                startAddress: 'B1',
                endAddress: 'B1',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('runs durable current-cell explanation workflows from the active selection', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'C1', 'B1+1')
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
          scope: 'shared',
          executionPolicy: 'ownerReview',
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B1',
            },
            viewport: {
              rowStart: 0,
              rowEnd: 10,
              colStart: 0,
              colEnd: 5,
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
          workflowTemplate: 'explainSelectionCell',
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'explainSelectionCell',
          title: 'Explain Current Cell',
          status: 'completed',
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'explain-cell',
              status: 'completed',
            }),
          ]),
          artifact: expect.objectContaining({
            title: 'Current Cell',
            text: expect.stringContaining('Cell: Sheet1!B1'),
          }),
        }),
      )
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Explain Current Cell',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Sheet1',
                startAddress: 'B1',
                endAddress: 'B1',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('runs durable workbook search workflows with query input', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Revenue')
    engine.setCellValue('Revenue', 'A1', 'Region')
    engine.setCellValue('Revenue', 'B1', 'Revenue')
    engine.setCellFormula('Revenue', 'B2', 'SUM(B3:B5)')
    engine.setCellValue('Revenue', 'A2', 'West')
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
          workflowTemplate: 'searchWorkbookQuery',
          query: 'revenue',
          limit: 5,
        },
      })

      expect(runningSnapshot.workflowRuns[0]?.status).toBe('running')
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'searchWorkbookQuery',
          title: 'Search Workbook',
          status: 'completed',
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: 'search-workbook',
              status: 'completed',
            }),
          ]),
          artifact: expect.objectContaining({
            title: 'Workbook Search',
            text: expect.stringContaining('Query: revenue'),
          }),
        }),
      )
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Search Workbook',
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: 'range',
                sheetName: 'Revenue',
              }),
            ]),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('applies create-sheet workflows immediately in private threads', async () => {
    const fakeCodex = new FakeCodexTransport()
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const getWorkbookHeadRevision = vi.fn(async () => 7)
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined)
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
      const snapshot = await waitForWorkflowStatus(service, 'thr-test', 'alex@example.com', 'completed')
      expect(getWorkbookHeadRevision).toHaveBeenCalledWith('doc-1')
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2)
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: 'createSheet',
          status: 'completed',
          summary:
            'Workflow mutation verification incomplete at revision r8: No target cell range was available for authoritative readback.',
          mutationExecuted: true,
          verificationComplete: false,
          mutationStatus: 'verification_incomplete',
          mutationReceipt: expect.objectContaining({
            status: 'verification_incomplete',
            toolName: 'workflow:createSheet',
          }),
          artifact: expect.objectContaining({
            title: 'Create Sheet Preview',
          }),
        }),
      )
      expect(snapshot.reviewQueueItems).toEqual([])
      expect(applyAgentCommandBundle).toHaveBeenCalledWith(
        'doc-1',
        expect.objectContaining({
          commands: [expect.objectContaining({ kind: 'createSheet', name: 'Forecast' })],
        }),
        expect.anything(),
        expect.objectContaining({
          userID: 'alex@example.com',
        }),
      )
      expect(appendWorkbookAgentRun).toHaveBeenCalledTimes(1)
      expect(snapshot.executionRecords).toEqual([
        expect.objectContaining({
          summary: 'Create sheet Forecast',
          appliedRevision: 8,
          appliedBy: 'auto',
        }),
      ])
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: 'Applied automatically workbook change set at revision r8: Create sheet Forecast',
          }),
          expect.objectContaining({
            kind: 'system',
            text: 'Completed workflow: Create Sheet',
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })
})
