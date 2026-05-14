import { SpreadsheetEngine } from '@bilig/core'
import {
  createWorkbookAgentCommandBundle,
  type CodexDynamicToolCallResult,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { ValueTag } from '@bilig/protocol'
import type { AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'
import { describe, expect, it } from 'vitest'
import { z } from 'zod'
import { applyWorkbookAgentCommandBundleWithUndoCapture } from '../zero/workbook-agent-apply.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookChangeRecord } from '../zero/workbook-change-store.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { buildWorkbookAgentVerificationReport, stageWorkbookAgentCommandResult } from './workbook-agent-mutation-receipt.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:mutation-receipt-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 'Seed')
  engine.setCellValue('Sheet1', 'B2', 'Before')
  engine.setCellValue('Sheet1', 'C3', 'Verified value')
  return engine
}

function createZeroSyncHarness(
  engine: SpreadsheetEngine,
  options: {
    readonly headRevision?: number
    readonly calculatedRevision?: number
    readonly changes?: readonly WorkbookChangeRecord[]
  } = {},
) {
  const zeroSyncService: ZeroSyncService = {
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
      const revision = options.headRevision ?? 1
      const calculatedRevision = options.calculatedRevision ?? revision
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision,
          calculatedRevision,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-12T12:00:00.000Z',
        }),
        headRevision: revision,
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
      return [...(options.changes ?? [])]
    },
    async listWorkbookAgentRuns() {
      return []
    },
    async appendWorkbookAgentRun() {
      throw new Error('not used')
    },
    async listWorkbookAgentThreadRuns() {
      return []
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
      return options.headRevision ?? 1
    },
    async loadAuthoritativeEvents() {
      return {
        afterRevision: options.headRevision ?? 1,
        headRevision: options.headRevision ?? 1,
        calculatedRevision: options.calculatedRevision ?? options.headRevision ?? 1,
        events: [],
      } satisfies AuthoritativeWorkbookEventBatch
    },
  }
  return { zeroSyncService }
}

function createBundle(command: WorkbookAgentCommand, bundleId: string): WorkbookAgentCommandBundle {
  return createWorkbookAgentCommandBundle({
    bundleId,
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'mutation receipt test',
    baseRevision: 1,
    now: 1,
    context: null,
    commands: [command],
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
    throw new Error('Expected affected range for execution record')
  }
  return {
    id: `run-${input.bundle.id}`,
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

function createRenderedContext(input: {
  readonly address: string
  readonly value: string | null
  readonly capturedRevision: number
  readonly styleId?: string | null
  readonly numberFormatId?: string | null
}): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: 'Sheet1',
      address: input.address,
      range: {
        startAddress: input.address,
        endAddress: input.address,
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
      capturedRevision: input.capturedRevision,
      batchId: input.capturedRevision,
      selection: {
        range: {
          sheetName: 'Sheet1',
          startAddress: input.address,
          endAddress: input.address,
        },
        rowCount: 1,
        columnCount: 1,
        cellCount: 1,
        truncated: false,
        rows: [
          [
            {
              address: input.address,
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
              styleId: input.styleId ?? null,
              numberFormatId: input.numberFormatId ?? null,
              style: null,
            },
          ],
        ],
      },
      visibleRange: null,
    },
  }
}

function parsePayload(result: CodexDynamicToolCallResult): unknown {
  expect(result.success).toBe(true)
  const item = result.contentItems[0]
  expect(item?.type).toBe('inputText')
  return JSON.parse(item && 'text' in item ? item.text : '')
}

const stagedPayloadSchema = z.object({
  applied: z.literal(false),
  staged: z.literal(true),
  reviewQueued: z.literal(true),
  queuedForTurnApply: z.literal(false),
  status: z.literal('staged'),
  bundleId: z.string(),
  mutationReceipt: z.object({
    status: z.literal('staged'),
    authoritativeReadback: z.object({
      requested: z.literal(false),
    }),
    renderedReadback: z.object({
      requested: z.literal(false),
    }),
    undo: z.object({
      available: z.literal(false),
      reasonUnavailable: z.string(),
    }),
    warnings: z.array(z.string()),
  }),
})

const queuedPayloadSchema = z.object({
  applied: z.literal(false),
  staged: z.literal(false),
  reviewQueued: z.literal(false),
  queuedForTurnApply: z.literal(true),
  status: z.literal('queued'),
  bundleId: z.string(),
  mutationReceipt: z.object({
    status: z.literal('queued'),
    warnings: z.array(z.string()),
  }),
})

const appliedPayloadSchema = z.object({
  applied: z.literal(true),
  staged: z.literal(false),
  reviewQueued: z.literal(false),
  queuedForTurnApply: z.literal(false),
  status: z.literal('verification_incomplete'),
  revision: z.literal(2),
  mutationReceipt: z.object({
    status: z.literal('verification_incomplete'),
    authoritativeReadback: z.object({
      requested: z.literal(true),
      matched: z.literal(true),
      incompleteReason: z.null(),
    }),
    renderedReadback: z.object({
      requested: z.literal(true),
      available: z.literal(false),
      matched: z.null(),
      incompleteReason: z.string(),
    }),
    undo: z.object({
      available: z.literal(true),
      token: z.literal('revision:2'),
    }),
    warnings: z.array(z.string()),
  }),
  verification: z.object({
    appliedRevision: z.literal(2),
  }),
})

describe('workbook agent mutation receipt helpers', () => {
  it('returns a review-queued staged payload when the stage command only produces a bundle', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [['Review later']],
    }

    const result = await stageWorkbookAgentCommandResult(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: async () => createBundle(command, 'bundle-staged'),
      },
      command,
      'writeRange',
    )

    const payload = stagedPayloadSchema.parse(parsePayload(result))
    expect(payload.bundleId).toBe('bundle-staged')
    expect(payload.mutationReceipt.undo.reasonUnavailable).toContain('has not been applied yet')
    expect(payload.mutationReceipt.warnings).toContain(
      'Workbook change set is waiting for owner review and has not modified the workbook yet.',
    )
  })

  it('returns a queued payload when the stage command defers apply to the turn loop', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [['Queued apply']],
    }

    const result = await stageWorkbookAgentCommandResult(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: async () => ({
          bundle: createBundle(command, 'bundle-queued'),
          executionRecord: null,
          disposition: 'queuedForTurnApply',
        }),
      },
      command,
      'writeRange',
    )

    const payload = queuedPayloadSchema.parse(parsePayload(result))
    expect(payload.bundleId).toBe('bundle-queued')
    expect(payload.mutationReceipt.warnings).toContain(
      'Queued workbook change sets are not completed mutations. The assistant must wait for apply and verify before claiming success.',
    )
  })

  it('derives authoritative proof and undo metadata for applied writes without rendered context', async () => {
    const engine = await createEngine()
    const appliedValue = 'Applied by receipt'
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [[appliedValue]],
    }
    const bundle = createBundle(command, 'bundle-applied')
    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
    const { zeroSyncService } = createZeroSyncHarness(engine, {
      headRevision: 2,
      calculatedRevision: 2,
      changes: [
        {
          revision: 2,
          actorUserId: 'alex@example.com',
          clientMutationId: null,
          eventKind: 'applyAgentCommandBundle',
          summary: 'Write cells in Sheet1!B2',
          sheetId: null,
          sheetName: 'Sheet1',
          anchorAddress: 'B2',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
          undoBundle,
          revertedByRevision: null,
          revertsRevision: null,
          createdAtUnixMs: 2,
        },
      ],
    })

    const result = await stageWorkbookAgentCommandResult(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: createExecutionRecord({
            bundle,
            appliedRevision: 2,
            afterInput: appliedValue,
            includePreviewDiff: false,
          }),
        }),
      },
      command,
      'writeRange',
    )

    const payload = appliedPayloadSchema.parse(parsePayload(result))
    expect(payload.mutationReceipt.renderedReadback.incompleteReason).toContain('No browser-rendered context')
    expect(payload.mutationReceipt.warnings).toContain('No browser-rendered context was attached to this tool call.')
  })

  it('does not report applied when authoritative readback disagrees with the claimed write', async () => {
    const engine = await createEngine()
    engine.setCellValue('Sheet1', 'B2', 'Different committed value')
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [['Claimed value']],
    }
    const bundle = createBundle(command, 'bundle-authoritative-mismatch')
    const { zeroSyncService } = createZeroSyncHarness(engine, {
      headRevision: 2,
      calculatedRevision: 2,
    })

    const result = await stageWorkbookAgentCommandResult(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: createRenderedContext({
          address: 'B2',
          value: 'Different committed value',
          capturedRevision: 2,
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: createExecutionRecord({
            bundle,
            appliedRevision: 2,
            afterInput: 'Claimed value',
          }),
        }),
      },
      command,
      'writeRange',
    )

    const payload = z
      .object({
        applied: z.literal(true),
        status: z.literal('verification_incomplete'),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          authoritativeReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(false),
            incompleteReason: z.string(),
          }),
          renderedReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
          }),
          warnings: z.array(z.string()),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.authoritativeReadback.incompleteReason).toContain('Authoritative readback did not match')
    expect(payload.mutationReceipt.warnings).toContain('Authoritative readback did not match preview expectations.')
  })

  it('reports applied for format mutations when authoritative and rendered proof agree', async () => {
    const engine = await createEngine()
    const command: WorkbookAgentCommand = {
      kind: 'formatRange',
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
    }
    const bundle = createBundle(command, 'bundle-format-needs-proof')
    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
    const committedCell = engine.getCell('Sheet1', 'B2')
    expect(committedCell.styleId).toBeTruthy()
    const { zeroSyncService } = createZeroSyncHarness(engine, {
      headRevision: 2,
      calculatedRevision: 2,
      changes: [
        {
          revision: 2,
          actorUserId: 'alex@example.com',
          clientMutationId: null,
          eventKind: 'applyAgentCommandBundle',
          summary: 'Format cells in Sheet1!B2',
          sheetId: null,
          sheetName: 'Sheet1',
          anchorAddress: 'B2',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
          undoBundle,
          revertedByRevision: null,
          revertsRevision: null,
          createdAtUnixMs: 2,
        },
      ],
    })

    const result = await stageWorkbookAgentCommandResult(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: createRenderedContext({
          address: 'B2',
          value: 'Before',
          capturedRevision: 2,
          styleId: committedCell.styleId ?? null,
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: createExecutionRecord({
            bundle,
            appliedRevision: 2,
            afterInput: null,
            includePreviewDiff: false,
          }),
        }),
      },
      command,
      'formatRange',
    )

    const payload = z
      .object({
        applied: z.literal(true),
        status: z.literal('applied'),
        mutationReceipt: z.object({
          status: z.literal('applied'),
          authoritativeReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
            incompleteReason: z.null(),
          }),
          renderedReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
          }),
          warnings: z.array(z.string()),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.warnings).toEqual([])
  })

  it('builds verification reports with matching rendered readback and optional audits disabled', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine, {
      headRevision: 3,
      calculatedRevision: 3,
    })

    const report = await buildWorkbookAgentVerificationReport({
      context: {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: createRenderedContext({
          address: 'C3',
          value: 'Verified value',
          capturedRevision: 3,
        }),
        zeroSyncService,
        stageCommand: async () =>
          createBundle(
            {
              kind: 'writeRange',
              sheetName: 'Sheet1',
              startAddress: 'C3',
              values: [['unused']],
            },
            'bundle-unused',
          ),
      },
      revision: 3,
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'C3',
        },
      ],
      includeFormulaIssues: false,
      includeInvariants: false,
    })

    expect(report.appliedRevision).toBe(3)
    expect(report.recalculationStatus.upToDate).toBe(true)
    expect(report.authoritativeReadback).toHaveLength(1)
    expect(report.renderedReadback).toEqual([
      expect.objectContaining({
        requested: true,
        available: true,
        matched: true,
        stale: false,
        capturedRevision: 3,
        incompleteReason: null,
      }),
    ])
    expect(report.formulaIssues).toBeNull()
    expect(report.invariants).toBeNull()
  })
})
