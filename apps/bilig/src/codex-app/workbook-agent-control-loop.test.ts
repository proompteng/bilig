import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import {
  createWorkbookAgentCommandBundle,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import { ValueTag } from '@bilig/protocol'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { z } from 'zod'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import { applyWorkbookAgentCommandBundleWithUndoCapture } from '../zero/workbook-agent-apply.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookChangeRecord } from '../zero/workbook-change-store.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { handleWorkbookAgentToolCall } from './workbook-agent-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-control-loop',
    replicaId: 'server:test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 42)
  return engine
}

function createZeroSyncHarness(
  engine: SpreadsheetEngine,
  input?: {
    readonly headRevision?: number
    readonly calculatedRevision?: number
    readonly changes?: readonly WorkbookChangeRecord[]
  },
): ZeroSyncService {
  const headRevision = input?.headRevision ?? 1
  const calculatedRevision = input?.calculatedRevision ?? headRevision
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
    async inspectWorkbook(_documentId, task) {
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: headRevision,
          calculatedRevision,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-30T12:00:00.000Z',
        }),
        headRevision,
        calculatedRevision,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    },
    async applyServerMutator() {
      throw new Error('not used')
    },
    async applyAgentCommandBundle() {
      throw new Error('not used')
    },
    async listWorkbookChanges() {
      return [...(input?.changes ?? [])]
    },
    async listWorkbookAgentRuns() {
      return []
    },
    async listWorkbookAgentThreadRuns() {
      return []
    },
    async appendWorkbookAgentRun() {
      throw new Error('not used')
    },
    async listWorkbookAgentThreadSummaries() {
      return []
    },
    async loadWorkbookAgentThreadState() {
      return null
    },
    async saveWorkbookAgentThreadState() {
      throw new Error('not used')
    },
    async listWorkbookThreadWorkflowRuns() {
      return []
    },
    async upsertWorkbookWorkflowRun() {
      throw new Error('not used')
    },
    async getWorkbookHeadRevision() {
      return headRevision
    },
    async loadAuthoritativeEvents() {
      throw new Error('not used')
    },
  }
}

function createBundle(command: WorkbookAgentCommand): WorkbookAgentCommandBundle {
  return createWorkbookAgentCommandBundle({
    bundleId: 'bundle-control-loop',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'Verify workbook agent control loop',
    baseRevision: 1,
    context: null,
    commands: [command],
    now: 1,
  })
}

function createExecutionRecord(input: {
  readonly bundle: WorkbookAgentCommandBundle
  readonly appliedRevision: number
  readonly afterInput: string | number | boolean | null
  readonly includePreviewDiff?: boolean
}): WorkbookAgentExecutionRecord {
  const range = input.bundle.affectedRanges[0]
  if (!range) {
    throw new Error('Expected an affected range')
  }
  return {
    id: 'run-control-loop',
    bundleId: input.bundle.id,
    documentId: input.bundle.documentId,
    threadId: input.bundle.threadId,
    turnId: input.bundle.turnId,
    actorUserId: 'alex@example.com',
    goalText: input.bundle.goalText,
    planText: null,
    summary: input.bundle.summary,
    scope: input.bundle.scope,
    riskClass: input.bundle.riskClass,
    acceptedScope: 'full',
    appliedBy: 'auto',
    baseRevision: input.bundle.baseRevision,
    appliedRevision: input.appliedRevision,
    context: input.bundle.context,
    commands: input.bundle.commands,
    preview: {
      ranges: input.bundle.affectedRanges,
      structuralChanges: [],
      cellDiffs:
        input.includePreviewDiff === false
          ? []
          : [
              {
                sheetName: range.sheetName,
                address: range.startAddress,
                beforeInput: null,
                beforeFormula: null,
                afterInput: input.afterInput,
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
    },
    createdAtUnixMs: 2,
    appliedAtUnixMs: 2,
  }
}

function readToolJson(response: Awaited<ReturnType<typeof handleWorkbookAgentToolCall>>): unknown {
  const text = response.contentItems[0]
  expect(text?.type).toBe('inputText')
  return JSON.parse(text && 'text' in text ? text.text : '')
}

describe('workbook agent control loop receipts', () => {
  it('explicitly reports applied writes as not staged and not queued', async () => {
    const engine = await createEngine()
    const sentinel = 'agent-control-loop-sentinel'
    const zeroSyncService = createZeroSyncHarness(engine, {
      headRevision: 2,
      changes: [
        {
          revision: 2,
          actorUserId: 'alex@example.com',
          clientMutationId: null,
          eventKind: 'applyAgentCommandBundle',
          summary: 'Write cells in Sheet1!F6',
          sheetId: null,
          sheetName: 'Sheet1',
          anchorAddress: 'F6',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'F6',
            endAddress: 'F6',
          },
          rangeInvalid: false,
          undoBundle: {
            kind: 'engineOps',
            ops: [],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAtUnixMs: 2,
        },
      ],
    })
    const uiContext: WorkbookAgentUiContext = {
      selection: {
        sheetName: 'Sheet1',
        address: 'F6',
        range: {
          startAddress: 'F6',
          endAddress: 'F6',
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: 20,
        colStart: 0,
        colEnd: 10,
      },
      rendered: {
        capturedAtUnixMs: 10,
        capturedRevision: 2,
        batchId: 2,
        selection: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'F6',
            endAddress: 'F6',
          },
          rowCount: 1,
          columnCount: 1,
          cellCount: 1,
          truncated: false,
          rows: [
            [
              {
                address: 'F6',
                input: sentinel,
                value: { tag: ValueTag.String, value: sentinel },
                formula: null,
                displayFormat: sentinel,
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
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => {
      const bundle = createBundle(command)
      applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
      return {
        bundle,
        executionRecord: createExecutionRecord({
          bundle,
          appliedRevision: 2,
          afterInput: sentinel,
        }),
      }
    })

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-explicit-apply-receipt',
        tool: 'write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'F6',
          values: [[sentinel]],
        },
      },
    )

    const payload = z
      .object({
        applied: z.literal(true),
        staged: z.literal(false),
        queuedForTurnApply: z.literal(false),
        revision: z.literal(2),
        mutationReceipt: z.object({
          status: z.literal('applied'),
          authoritativeReadback: z.object({
            matched: z.literal(true),
          }),
          renderedReadback: z.object({
            matched: z.literal(true),
            stale: z.literal(false),
            capturedRevision: z.literal(2),
          }),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.mutationReceipt.renderedReadback.capturedRevision).toBe(2)
  })

  it('derives authoritative write_range proof when the server preview omits cell diffs', async () => {
    const engine = await createEngine()
    const sentinel = 'authoritative-derived-proof'
    const zeroSyncService = createZeroSyncHarness(engine, {
      headRevision: 2,
      changes: [
        {
          revision: 2,
          actorUserId: 'alex@example.com',
          clientMutationId: null,
          eventKind: 'applyAgentCommandBundle',
          summary: 'Write cells in Sheet1!H8',
          sheetId: null,
          sheetName: 'Sheet1',
          anchorAddress: 'H8',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'H8',
            endAddress: 'H8',
          },
          rangeInvalid: false,
          undoBundle: {
            kind: 'engineOps',
            ops: [],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAtUnixMs: 2,
        },
      ],
    })
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => {
      const bundle = createBundle(command)
      applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
      return {
        bundle,
        executionRecord: createExecutionRecord({
          bundle,
          appliedRevision: 2,
          afterInput: sentinel,
          includePreviewDiff: false,
        }),
      }
    })

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-derived-authoritative-proof',
        tool: 'write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'H8',
          values: [[sentinel]],
        },
      },
    )

    const payload = z
      .object({
        mutationReceipt: z.object({
          authoritativeReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
            incompleteReason: z.null(),
          }),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.mutationReceipt.authoritativeReadback.matched).toBe(true)
  })

  it('uses the current workbook revision for apply_and_verify rendered freshness', async () => {
    const engine = await createEngine()
    engine.setCellValue('Sheet1', 'B2', 'stale rendered target')
    const zeroSyncService = createZeroSyncHarness(engine, {
      headRevision: 5,
      calculatedRevision: 5,
    })
    const awaitRenderedRevision = vi.fn(async () => undefined)
    const uiContext: WorkbookAgentUiContext = {
      selection: {
        sheetName: 'Sheet1',
        address: 'B2',
        range: {
          startAddress: 'B2',
          endAddress: 'B2',
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
      rendered: {
        capturedAtUnixMs: 10,
        capturedRevision: 4,
        batchId: 4,
        selection: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
          rowCount: 1,
          columnCount: 1,
          cellCount: 1,
          truncated: false,
          rows: [
            [
              {
                address: 'B2',
                input: 'stale rendered target',
                value: { tag: ValueTag.String, value: 'stale rendered target' },
                formula: null,
                displayFormat: 'stale rendered target',
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

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext,
        zeroSyncService,
        awaitRenderedRevision,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-apply-and-verify-freshness',
        tool: 'apply_and_verify',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
        },
      },
    )

    const payload = z
      .object({
        status: z.literal('verification_incomplete'),
        verificationComplete: z.literal(false),
        appliedRevision: z.literal(5),
        renderedReadback: z.array(
          z.object({
            stale: z.literal(true),
            capturedRevision: z.literal(4),
            incompleteReason: z.string(),
          }),
        ),
      })
      .parse(readToolJson(response))
    expect(payload.renderedReadback[0]?.incompleteReason).toContain('older than the requested verification revision')
    expect(awaitRenderedRevision).toHaveBeenCalledWith(5)
  })

  it('proves an exact rendered target inside a truncated viewport', async () => {
    const engine = await createEngine()
    engine.setCellValue('Sheet1', 'B2', 'visible exact target')
    const zeroSyncService = createZeroSyncHarness(engine, {
      headRevision: 3,
    })
    const uiContext: WorkbookAgentUiContext = {
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 99,
        colStart: 0,
        colEnd: 20,
      },
      rendered: {
        capturedAtUnixMs: 10,
        capturedRevision: 3,
        batchId: 3,
        selection: null,
        visibleRange: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'C3',
          },
          rowCount: 3,
          columnCount: 3,
          cellCount: 9,
          truncated: true,
          rows: [
            [
              {
                address: 'A1',
                input: 42,
                value: { tag: ValueTag.Number, value: 42 },
                formula: null,
                displayFormat: '42',
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'B1',
                input: null,
                value: { tag: ValueTag.Empty },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'C1',
                input: null,
                value: { tag: ValueTag.Empty },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
            [
              {
                address: 'A2',
                input: null,
                value: { tag: ValueTag.Empty },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'B2',
                input: 'visible exact target',
                value: { tag: ValueTag.String, value: 'visible exact target' },
                formula: null,
                displayFormat: 'visible exact target',
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'C2',
                input: null,
                value: { tag: ValueTag.Empty },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
            [
              {
                address: 'A3',
                input: null,
                value: { tag: ValueTag.Empty },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'B3',
                input: null,
                value: { tag: ValueTag.Empty },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'C3',
                input: null,
                value: { tag: ValueTag.Empty },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
          ],
        },
      },
    }

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-rendered-truncated-exact-cell',
        tool: 'read_rendered_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B2',
        },
      },
    )

    const payload = z
      .object({
        renderedReadback: z.object({
          matched: z.literal(true),
          stale: z.literal(false),
          truncated: z.literal(false),
          sourceTruncated: z.literal(true),
          incompleteReason: z.null(),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.renderedReadback.sourceTruncated).toBe(true)
  })

  it('reports partial selection confidence when browser confirmation is not proven', async () => {
    const engine = await createEngine()
    const zeroSyncService = createZeroSyncHarness(engine)
    const updateUiContext = vi.fn(async () => undefined)

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A1',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        updateUiContext,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-selection-confidence',
        tool: 'set_selection',
        arguments: {
          sheetName: 'Sheet1',
          address: 'C3',
        },
      },
    )

    const payload = z
      .object({
        updated: z.literal(true),
        verificationComplete: z.literal(false),
        selectionConfidence: z.object({
          level: z.literal('model_only'),
          browserConfirmed: z.literal(false),
          reason: z.string(),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.selectionConfidence.reason).toContain('browser')
    expect(updateUiContext).toHaveBeenCalledTimes(1)
  })

  it('marks undo verification incomplete when rendered proof is stale', async () => {
    const engine = await createEngine()
    engine.setCellValue('Sheet1', 'F6', 'temporary')
    const workbookChanges: WorkbookChangeRecord[] = [
      {
        revision: 2,
        actorUserId: 'alex@example.com',
        clientMutationId: null,
        eventKind: 'applyAgentCommandBundle',
        summary: 'Write cells in Sheet1!F6',
        sheetId: null,
        sheetName: 'Sheet1',
        anchorAddress: 'F6',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'F6',
          endAddress: 'F6',
        },
        rangeInvalid: false,
        undoBundle: {
          kind: 'engineOps',
          ops: [],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAtUnixMs: 2,
      },
    ]
    let headRevision = 2
    const zeroSyncService = createZeroSyncHarness(engine, {
      headRevision,
      changes: workbookChanges,
    })
    zeroSyncService.getWorkbookHeadRevision = vi.fn(async () => headRevision)
    zeroSyncService.applyServerMutator = vi.fn(async () => {
      engine.clearCell('Sheet1', 'F6')
      headRevision = 3
    })
    zeroSyncService.inspectWorkbook = async (_documentId, task) => {
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: headRevision,
          calculatedRevision: headRevision,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-30T12:00:00.000Z',
        }),
        headRevision,
        calculatedRevision: headRevision,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    }
    const awaitRenderedRevision = vi.fn(async () => undefined)

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'F6',
            range: {
              startAddress: 'F6',
              endAddress: 'F6',
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 20,
            colStart: 0,
            colEnd: 10,
          },
          rendered: {
            capturedAtUnixMs: 10,
            capturedRevision: 2,
            batchId: 2,
            selection: {
              range: {
                sheetName: 'Sheet1',
                startAddress: 'F6',
                endAddress: 'F6',
              },
              rowCount: 1,
              columnCount: 1,
              cellCount: 1,
              truncated: false,
              rows: [
                [
                  {
                    address: 'F6',
                    input: 'temporary',
                    value: { tag: ValueTag.String, value: 'temporary' },
                    formula: null,
                    displayFormat: 'temporary',
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
        zeroSyncService,
        awaitRenderedRevision,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-undo-stale-rendered-proof',
        tool: 'undo_workbook_mutation',
        arguments: {
          revision: 2,
        },
      },
    )

    const payload = z
      .object({
        undone: z.literal(true),
        applied: z.literal(true),
        staged: z.literal(false),
        queuedForTurnApply: z.literal(false),
        status: z.literal('verification_incomplete'),
        verificationComplete: z.literal(false),
        revision: z.object({
          before: z.literal(2),
          after: z.literal(3),
          reverted: z.literal(2),
        }),
        verification: z.object({
          renderedReadback: z.array(
            z.object({
              stale: z.literal(true),
              capturedRevision: z.literal(2),
              incompleteReason: z.string(),
            }),
          ),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.verification.renderedReadback[0]?.incompleteReason).toContain('older than the requested verification revision')
    expect(awaitRenderedRevision).toHaveBeenCalledWith(3)
  })
})
