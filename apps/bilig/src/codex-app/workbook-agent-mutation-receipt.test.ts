import {
  createWorkbookAgentCommandBundle,
  type CodexDynamicToolCallResult,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import type { AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'
import { describe, expect, it } from 'vitest'
import { z } from 'zod'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import { applyWorkbookAgentCommandBundleWithUndoCapture } from '../zero/workbook-agent-apply.js'
import type { WorkbookChangeRecord } from '../zero/workbook-change-store.js'
import { stageWorkbookAgentCommandResult } from './workbook-agent-mutation-receipt.js'
import { buildWorkbookAgentVisibleCommitBarrierOutcome } from './workbook-agent-visible-commit-barrier.js'

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
    readonly changesError?: Error
  } = {},
) {
  const zeroSyncService: ZeroSyncService = {
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
    async applyWorkbookPlanData() {
      throw new Error('not used')
    },
    async listWorkbookChanges() {
      if (options.changesError) {
        throw options.changesError
      }
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
  readonly sceneProof?: Partial<NonNullable<NonNullable<WorkbookAgentUiContext['rendered']>['visibleSceneProof']>> | null
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
      visibleSceneProof: input.sceneProof === null ? null : createVisibleSceneProof(input.sceneProof ?? {}, input.capturedRevision),
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

function createVisibleSceneProof(
  overrides: Partial<NonNullable<NonNullable<WorkbookAgentUiContext['rendered']>['visibleSceneProof']>>,
  revision: number,
): NonNullable<NonNullable<WorkbookAgentUiContext['rendered']>['visibleSceneProof']> {
  const revisionText = String(revision)
  return {
    rendererMode: 'typegpu-v3',
    frameProofStatus: 'presented',
    frameProofSignature: `frame-${revisionText}`,
    presentedFrameProofSignature: `frame-${revisionText}`,
    currentSceneEpochSignature: `epoch-${revisionText}`,
    currentSceneOwnershipSignature: `scene-${revisionText}`,
    presentedSceneEpochSignature: `epoch-${revisionText}`,
    presentedSceneOwnershipSignature: `scene-${revisionText}`,
    currentSceneEpoch: `tile-${revisionText}`,
    presentedSceneEpoch: `tile-${revisionText}`,
    currentFillHandleRevision: `fill-${revisionText}`,
    presentedFillHandleRevision: `fill-${revisionText}`,
    currentSelectionRevision: `selection-${revisionText}`,
    presentedSelectionRevision: `selection-${revisionText}`,
    currentViewportRevision: `viewport-${revisionText}`,
    presentedViewportRevision: `viewport-${revisionText}`,
    currentSemanticMutationRevision: revisionText,
    presentedSemanticMutationRevision: revisionText,
    currentWorkbookRevision: revisionText,
    presentedWorkbookRevision: revisionText,
    gridAuthoritativeRevision: revisionText,
    typeGpuAuthoritativeRevision: revisionText,
    visibleAuthoritativeRevision: revisionText,
    tileSceneRevision: `tile-${revisionText}`,
    visibleRenderRevision: `tile-${revisionText}`,
    hasPresentedFrame: true,
    hasPresentedVisibleFrame: true,
    frameProofMatchesPresentedFrame: true,
    visibleSceneEpochMatchesPresentedFrame: true,
    visibleSceneOwnershipMatchesPresentedFrame: true,
    visibleAuthoritativeRevisionMatchesGrid: true,
    visibleRenderRevisionMatchesTileScene: true,
    ...overrides,
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
  mutationExecuted: z.literal(false),
  verificationComplete: z.literal(false),
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
    semanticReadback: z.object({
      requested: z.literal(false),
      matched: z.null(),
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
  mutationExecuted: z.literal(false),
  verificationComplete: z.literal(false),
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
  applied: z.literal(false),
  mutationExecuted: z.literal(true),
  verificationComplete: z.literal(false),
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
    semanticReadback: z.object({
      requested: z.literal(true),
      matched: z.literal(false),
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

describe('workbook agent mutation receipt applied and staged proofs', () => {
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

  it('exposes staged review items through the shared visible commit barrier outcome', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [['Review later']],
    }
    const bundle = createBundle(command, 'bundle-staged-barrier')

    const outcome = await buildWorkbookAgentVisibleCommitBarrierOutcome({
      context: {
        documentId: 'doc-1',
        uiContext: null,
        zeroSyncService,
      },
      toolName: 'writeRange',
      normalized: {
        bundle,
        executionRecord: null,
        disposition: 'reviewQueued',
      },
    })

    expect(outcome).toMatchObject({
      mutationExecuted: false,
      verificationComplete: false,
      status: 'staged',
      appliedRevision: null,
      mutationReceipt: {
        toolName: 'writeRange',
        status: 'staged',
      },
    })
    expect(outcome.summary).toContain('the workbook is unchanged until this is applied')
    expect(outcome.mutationReceipt.warnings).toContain(
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
          rangeInvalid: false,
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
    expect(payload.mutationReceipt.semanticReadback.incompleteReason).toContain('No browser-rendered context')
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
        applied: z.literal(false),
        mutationExecuted: z.literal(true),
        verificationComplete: z.literal(false),
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
          semanticReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(false),
            incompleteReason: z.string(),
          }),
          warnings: z.array(z.string()),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.authoritativeReadback.incompleteReason).toContain('Authoritative readback did not match')
    expect(payload.mutationReceipt.semanticReadback.incompleteReason).toContain('Authoritative readback')
    expect(payload.mutationReceipt.warnings).toContain('Authoritative readback did not match preview expectations.')
  })

  it('does not report applied when rendered proof is stale even if authoritative proof and undo agree', async () => {
    const engine = await createEngine()
    const appliedValue = 'Rendered proof must be fresh'
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [[appliedValue]],
    }
    const bundle = createBundle(command, 'bundle-stale-rendered-proof')
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
          rangeInvalid: false,
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
          value: appliedValue,
          capturedRevision: 1,
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: createExecutionRecord({
            bundle,
            appliedRevision: 2,
            afterInput: appliedValue,
          }),
        }),
      },
      command,
      'writeRange',
    )

    const payload = z
      .object({
        applied: z.literal(false),
        mutationExecuted: z.literal(true),
        verificationComplete: z.literal(false),
        status: z.literal('verification_incomplete'),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          authoritativeReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
          }),
          renderedReadback: z.object({
            requested: z.literal(true),
            matched: z.null(),
            stale: z.literal(true),
            incompleteReason: z.string(),
          }),
          semanticReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(false),
            incompleteReason: z.string(),
          }),
          undo: z.object({
            available: z.literal(true),
          }),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.status).toBe('verification_incomplete')
    expect(payload.mutationReceipt.semanticReadback.matched).toBe(false)
    expect(payload.mutationReceipt.semanticReadback.incompleteReason).toContain('Rendered')
  })

  it('does not report applied when only the first mutated target range is visibly proven', async () => {
    const engine = await createEngine()
    const firstCommand: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [['First visible target']],
    }
    const secondCommand: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'C3',
      values: [['Second hidden target']],
    }
    const bundle = createWorkbookAgentCommandBundle({
      bundleId: 'bundle-multi-range-render-proof',
      documentId: 'doc-1',
      threadId: 'thr-1',
      turnId: 'turn-1',
      goalText: 'mutation receipt test',
      baseRevision: 1,
      now: 1,
      context: null,
      commands: [firstCommand, secondCommand],
    })
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
          summary: 'Write cells in Sheet1!B2 and Sheet1!C3',
          sheetId: null,
          sheetName: 'Sheet1',
          anchorAddress: 'B2',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'C3',
          },
          rangeInvalid: false,
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
          value: 'First visible target',
          capturedRevision: 2,
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: {
            id: `run-${bundle.id}`,
            bundleId: bundle.id,
            documentId: bundle.documentId,
            threadId: bundle.threadId,
            turnId: bundle.turnId,
            actorUserId: 'alex@example.com',
            goalText: bundle.goalText,
            planText: null,
            summary: bundle.summary,
            scope: bundle.scope,
            riskClass: bundle.riskClass,
            acceptedScope: 'full',
            appliedBy: 'auto',
            baseRevision: bundle.baseRevision,
            appliedRevision: 2,
            context: bundle.context,
            commands: bundle.commands,
            preview: {
              ranges: bundle.affectedRanges,
              structuralChanges: [],
              cellDiffs: [
                {
                  sheetName: 'Sheet1',
                  address: 'B2',
                  beforeInput: null,
                  beforeFormula: null,
                  afterInput: 'First visible target',
                  afterFormula: null,
                  changeKinds: ['input'],
                },
                {
                  sheetName: 'Sheet1',
                  address: 'C3',
                  beforeInput: null,
                  beforeFormula: null,
                  afterInput: 'Second hidden target',
                  afterFormula: null,
                  changeKinds: ['input'],
                },
              ],
              effectSummary: {
                displayedCellDiffCount: 2,
                truncatedCellDiffs: false,
                inputChangeCount: 2,
                formulaChangeCount: 0,
                styleChangeCount: 0,
                numberFormatChangeCount: 0,
                structuralChangeCount: 0,
              },
            },
            createdAtUnixMs: 2,
            appliedAtUnixMs: 2,
          },
        }),
      },
      firstCommand,
      'writeRange',
    )

    const payload = z
      .object({
        applied: z.literal(false),
        status: z.literal('verification_incomplete'),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          authoritativeReadback: z.object({
            matched: z.literal(true),
          }),
          renderedReadback: z.object({
            matched: z.null(),
            available: z.literal(false),
            rangeProofs: z.array(
              z.object({
                requestedRange: z.object({
                  sheetName: z.literal('Sheet1'),
                  startAddress: z.string(),
                  endAddress: z.string(),
                }),
                matched: z.union([z.literal(true), z.null()]),
              }),
            ),
            incompleteReason: z.string(),
          }),
          semanticReadback: z.object({
            matched: z.literal(false),
          }),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.renderedReadback.rangeProofs).toHaveLength(2)
    expect(payload.mutationReceipt.renderedReadback.rangeProofs[0]?.matched).toBe(true)
    expect(payload.mutationReceipt.renderedReadback.rangeProofs[1]?.requestedRange.startAddress).toBe('C3')
    expect(payload.mutationReceipt.renderedReadback.rangeProofs[1]?.matched).toBeNull()
    expect(payload.mutationReceipt.renderedReadback.incompleteReason).toContain('Requested range was not captured')
  })

  it('does not truncate rendered proof after the first three target ranges', async () => {
    const engine = await createEngine()
    const commands: WorkbookAgentCommand[] = [
      {
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'B2',
        values: [['First visible target']],
      },
      {
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'C3',
        values: [['Second hidden target']],
      },
      {
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'D4',
        values: [['Third hidden target']],
      },
      {
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'E5',
        values: [['Fourth hidden target']],
      },
    ]
    const bundle = createWorkbookAgentCommandBundle({
      bundleId: 'bundle-four-range-render-proof',
      documentId: 'doc-1',
      threadId: 'thr-1',
      turnId: 'turn-1',
      goalText: 'mutation receipt test',
      baseRevision: 1,
      now: 1,
      context: null,
      commands,
    })
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
          summary: 'Write cells in Sheet1!B2, Sheet1!C3, Sheet1!D4, and Sheet1!E5',
          sheetId: null,
          sheetName: 'Sheet1',
          anchorAddress: 'B2',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'E5',
          },
          rangeInvalid: false,
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
          value: 'First visible target',
          capturedRevision: 2,
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: {
            id: `run-${bundle.id}`,
            bundleId: bundle.id,
            documentId: bundle.documentId,
            threadId: bundle.threadId,
            turnId: bundle.turnId,
            actorUserId: 'alex@example.com',
            goalText: bundle.goalText,
            planText: null,
            summary: bundle.summary,
            scope: bundle.scope,
            riskClass: bundle.riskClass,
            acceptedScope: 'full',
            appliedBy: 'auto',
            baseRevision: bundle.baseRevision,
            appliedRevision: 2,
            context: bundle.context,
            commands: bundle.commands,
            preview: {
              ranges: bundle.affectedRanges,
              structuralChanges: [],
              cellDiffs: [
                {
                  sheetName: 'Sheet1',
                  address: 'B2',
                  beforeInput: null,
                  beforeFormula: null,
                  afterInput: 'First visible target',
                  afterFormula: null,
                  changeKinds: ['input'],
                },
                {
                  sheetName: 'Sheet1',
                  address: 'C3',
                  beforeInput: null,
                  beforeFormula: null,
                  afterInput: 'Second hidden target',
                  afterFormula: null,
                  changeKinds: ['input'],
                },
                {
                  sheetName: 'Sheet1',
                  address: 'D4',
                  beforeInput: null,
                  beforeFormula: null,
                  afterInput: 'Third hidden target',
                  afterFormula: null,
                  changeKinds: ['input'],
                },
                {
                  sheetName: 'Sheet1',
                  address: 'E5',
                  beforeInput: null,
                  beforeFormula: null,
                  afterInput: 'Fourth hidden target',
                  afterFormula: null,
                  changeKinds: ['input'],
                },
              ],
              effectSummary: {
                displayedCellDiffCount: 4,
                truncatedCellDiffs: false,
                inputChangeCount: 4,
                formulaChangeCount: 0,
                styleChangeCount: 0,
                numberFormatChangeCount: 0,
                structuralChangeCount: 0,
              },
            },
            createdAtUnixMs: 2,
            appliedAtUnixMs: 2,
          },
        }),
      },
      commands[0]!,
      'writeRange',
    )

    const payload = z
      .object({
        applied: z.literal(false),
        status: z.literal('verification_incomplete'),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          authoritativeReadback: z.object({
            matched: z.literal(true),
            ranges: z.array(z.unknown()),
          }),
          renderedReadback: z.object({
            matched: z.null(),
            available: z.literal(false),
            rangeProofs: z.array(
              z.object({
                requestedRange: z.object({
                  sheetName: z.literal('Sheet1'),
                  startAddress: z.string(),
                  endAddress: z.string(),
                }),
                matched: z.union([z.literal(true), z.null()]),
              }),
            ),
            incompleteReason: z.string(),
          }),
          semanticReadback: z.object({
            matched: z.literal(false),
          }),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.authoritativeReadback.ranges).toHaveLength(4)
    expect(payload.mutationReceipt.renderedReadback.rangeProofs).toHaveLength(4)
    expect(payload.mutationReceipt.renderedReadback.rangeProofs.at(-1)?.requestedRange.startAddress).toBe('E5')
    expect(payload.mutationReceipt.renderedReadback.rangeProofs.at(-1)?.matched).toBeNull()
    expect(payload.mutationReceipt.renderedReadback.incompleteReason).toContain('Requested range was not captured')
  })

  it('does not report applied when undo proof is missing even if readbacks agree', async () => {
    const engine = await createEngine()
    const appliedValue = 'Undo proof required'
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [[appliedValue]],
    }
    const bundle = createBundle(command, 'bundle-missing-undo-proof')
    applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
    const { zeroSyncService } = createZeroSyncHarness(engine, {
      headRevision: 2,
      calculatedRevision: 2,
      changes: [],
    })

    const result = await stageWorkbookAgentCommandResult(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: createRenderedContext({
          address: 'B2',
          value: appliedValue,
          capturedRevision: 2,
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: createExecutionRecord({
            bundle,
            appliedRevision: 2,
            afterInput: appliedValue,
          }),
        }),
      },
      command,
      'writeRange',
    )

    const payload = z
      .object({
        applied: z.literal(false),
        mutationExecuted: z.literal(true),
        verificationComplete: z.literal(false),
        status: z.literal('verification_incomplete'),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          authoritativeReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
          }),
          renderedReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
          }),
          semanticReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
            incompleteReason: z.null(),
          }),
          undo: z.object({
            available: z.literal(false),
            lookupFailed: z.literal(false),
            reasonUnavailable: z.string(),
          }),
          warnings: z.array(z.string()),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.undo.reasonUnavailable).toContain('No persisted undo metadata')
    expect(payload.mutationReceipt.warnings).toContain('No persisted undo metadata was returned for the applied revision.')
  })

  it('surfaces undo history lookup failure instead of treating it as missing metadata', async () => {
    const engine = await createEngine()
    const appliedValue = 'History lookup failure'
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [[appliedValue]],
    }
    const bundle = createBundle(command, 'bundle-undo-lookup-failure')
    applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
    const { zeroSyncService } = createZeroSyncHarness(engine, {
      headRevision: 2,
      calculatedRevision: 2,
      changesError: new Error('history store unavailable'),
    })

    const result = await stageWorkbookAgentCommandResult(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: createRenderedContext({
          address: 'B2',
          value: appliedValue,
          capturedRevision: 2,
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: createExecutionRecord({
            bundle,
            appliedRevision: 2,
            afterInput: appliedValue,
          }),
        }),
      },
      command,
      'writeRange',
    )

    const payload = z
      .object({
        applied: z.literal(false),
        mutationExecuted: z.literal(true),
        verificationComplete: z.literal(false),
        status: z.literal('verification_incomplete'),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          undo: z.object({
            available: z.literal(false),
            lookupFailed: z.literal(true),
            reasonUnavailable: z.string(),
          }),
          warnings: z.array(z.string()),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.undo.reasonUnavailable).toBe(
      'Undo metadata lookup failed for applied revision r2: history store unavailable',
    )
    expect(payload.mutationReceipt.warnings).toContain('Undo metadata lookup failed for applied revision r2: history store unavailable')
  })

  it('does not summarize verification-incomplete execution records as fully applied', async () => {
    const engine = await createEngine()
    const appliedValue = 'Needs visible proof'
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [[appliedValue]],
    }
    const bundle = createBundle(command, 'bundle-incomplete-summary')
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
          rangeInvalid: false,
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
          }),
        }),
      },
      command,
      'writeRange',
    )

    const payload = z
      .object({
        status: z.literal('verification_incomplete'),
        summary: z.string(),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          warnings: z.array(z.string()),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.summary).toContain('Verification incomplete')
    expect(payload.summary).toContain('No browser-rendered context')
    expect(payload.summary).not.toContain('Applied workbook change set')
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
          rangeInvalid: false,
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
          semanticReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
            incompleteReason: z.null(),
          }),
          warnings: z.array(z.string()),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.warnings).toEqual([])
  })
})
