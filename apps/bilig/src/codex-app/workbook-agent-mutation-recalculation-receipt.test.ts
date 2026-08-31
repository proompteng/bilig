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
import { stageWorkbookAgentCommandResult } from './workbook-agent-mutation-receipt.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:mutation-recalculation-receipt-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'B2', 'Before')
  return engine
}

function createZeroSyncHarness(
  engine: SpreadsheetEngine,
  options: {
    readonly headRevision: number
    readonly calculatedRevision: number
    readonly changes: readonly WorkbookChangeRecord[]
  },
): ZeroSyncService {
  return {
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
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: options.headRevision,
          calculatedRevision: options.calculatedRevision,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-12T12:00:00.000Z',
        }),
        headRevision: options.headRevision,
        calculatedRevision: options.calculatedRevision,
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
      return [...options.changes]
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
      return options.headRevision
    },
    async loadAuthoritativeEvents() {
      return {
        afterRevision: options.headRevision,
        headRevision: options.headRevision,
        calculatedRevision: options.calculatedRevision,
        events: [],
      } satisfies AuthoritativeWorkbookEventBatch
    },
  }
}

function createBundle(command: WorkbookAgentCommand): WorkbookAgentCommandBundle {
  return createWorkbookAgentCommandBundle({
    bundleId: 'bundle-recalc-receipt',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'mutation recalculation receipt test',
    baseRevision: 1,
    now: 1,
    context: null,
    commands: [command],
  })
}

function createExecutionRecord(bundle: WorkbookAgentCommandBundle, appliedValue: string): WorkbookAgentExecutionRecord {
  return {
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
          afterInput: appliedValue,
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

function createUiContext(value: string, revision: number): WorkbookAgentUiContext {
  const revisionText = String(revision)
  return {
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
      capturedRevision: revision,
      batchId: revision,
      visibleSceneProof: {
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
      },
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
              input: value,
              value: {
                tag: ValueTag.String,
                value,
              },
              formula: null,
              displayFormat: value,
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

function parsePayload(result: CodexDynamicToolCallResult): unknown {
  expect(result.success).toBe(true)
  const item = result.contentItems[0]
  expect(item?.type).toBe('inputText')
  return JSON.parse(item && 'text' in item ? item.text : '')
}

async function applyWriteAndBuildReceipt(calculatedRevision: number): Promise<unknown> {
  const engine = await createEngine()
  const value = 'Recalculation proof required'
  const command: WorkbookAgentCommand = {
    kind: 'writeRange',
    sheetName: 'Sheet1',
    startAddress: 'B2',
    values: [[value]],
  }
  const bundle = createBundle(command)
  const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
  const changes: WorkbookChangeRecord[] = [
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
  ]

  const result = await stageWorkbookAgentCommandResult(
    {
      documentId: 'doc-1',
      session: { userID: 'alex@example.com', roles: ['editor'] },
      uiContext: createUiContext(value, 2),
      zeroSyncService: createZeroSyncHarness(engine, {
        headRevision: 2,
        calculatedRevision,
        changes,
      }),
      stageCommand: async () => ({
        bundle,
        executionRecord: createExecutionRecord(bundle, value),
      }),
    },
    command,
    'writeRange',
  )
  return parsePayload(result)
}

describe('workbook agent mutation recalculation receipt', () => {
  it('does not report applied until recalculation reaches the applied revision', async () => {
    const payload = z
      .object({
        applied: z.literal(false),
        verificationComplete: z.literal(false),
        status: z.literal('verification_incomplete'),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          recalculation: z.object({
            requested: z.literal(true),
            matched: z.literal(false),
            headRevision: z.literal(2),
            calculatedRevision: z.literal(1),
            requiredRevision: z.literal(2),
            incompleteReason: z.string(),
          }),
          authoritativeReadback: z.object({
            matched: z.literal(true),
          }),
          renderedReadback: z.object({
            matched: z.literal(true),
          }),
          semanticReadback: z.object({
            matched: z.literal(true),
          }),
          undo: z.object({
            available: z.literal(true),
          }),
          warnings: z.array(z.string()),
        }),
      })
      .parse(await applyWriteAndBuildReceipt(1))

    expect(payload.mutationReceipt.recalculation.incompleteReason).toContain('calculated r1 < required r2')
    expect(payload.mutationReceipt.warnings).toContain(payload.mutationReceipt.recalculation.incompleteReason)
  })

  it('includes passed recalculation proof in fully applied mutation receipts', async () => {
    const payload = z
      .object({
        applied: z.literal(true),
        verificationComplete: z.literal(true),
        status: z.literal('applied'),
        mutationReceipt: z.object({
          status: z.literal('applied'),
          recalculation: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
            headRevision: z.literal(2),
            calculatedRevision: z.literal(2),
            requiredRevision: z.literal(2),
            incompleteReason: z.null(),
          }),
          warnings: z.array(z.string()).length(0),
        }),
      })
      .parse(await applyWriteAndBuildReceipt(2))

    expect(payload.mutationReceipt.recalculation.matched).toBe(true)
  })
})
