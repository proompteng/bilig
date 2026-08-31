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
describe('workbook agent mutation receipt rendered proof freshness', () => {
  it('rejects applied status when rendered proof comes from a stale TypeGPU scene', async () => {
    const engine = await createEngine()
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [['Visible value']],
    }
    const bundle = createBundle(command, 'bundle-stale-scene-proof')
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
          value: 'Visible value',
          capturedRevision: 2,
          sceneProof: {
            presentedSceneOwnershipSignature: 'scene-1',
            visibleSceneOwnershipMatchesPresentedFrame: false,
          },
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: createExecutionRecord({
            bundle,
            appliedRevision: 2,
            afterInput: 'Visible value',
          }),
        }),
      },
      command,
      'writeRange',
    )

    const payload = z
      .object({
        applied: z.literal(false),
        status: z.literal('verification_incomplete'),
        summary: z.string(),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          authoritativeReadback: z.object({
            matched: z.literal(true),
          }),
          renderedReadback: z.object({
            matched: z.null(),
            stale: z.literal(true),
            visibleSceneProof: z.object({
              matched: z.literal(false),
              visibleSceneOwnershipMatchesPresentedFrame: z.literal(false),
              invalidReasons: z.array(z.string()),
            }),
          }),
          warnings: z.array(z.string()),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.summary).toContain('Verification incomplete')
    expect(payload.mutationReceipt.renderedReadback.visibleSceneProof.invalidReasons).toContain(
      'Presented visible-scene ownership does not match the current scene.',
    )
    expect(payload.mutationReceipt.warnings).toContain('Rendered TypeGPU visible-scene proof is incomplete or stale.')
  })

  it('rejects applied status when a rendered proof presents an older semantic mutation ownership', async () => {
    const engine = await createEngine()
    const command: WorkbookAgentCommand = {
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'B2',
      values: [['Visible value']],
    }
    const bundle = createBundle(command, 'bundle-stale-semantic-proof')
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
          value: 'Visible value',
          capturedRevision: 2,
          sceneProof: {
            presentedSemanticMutationRevision: '1',
          },
        }),
        zeroSyncService,
        stageCommand: async () => ({
          bundle,
          executionRecord: createExecutionRecord({
            bundle,
            appliedRevision: 2,
            afterInput: 'Visible value',
          }),
        }),
      },
      command,
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
            stale: z.literal(true),
            visibleSceneProof: z.object({
              matched: z.literal(false),
              visibleSemanticMutationRevisionMatchesPresentedFrame: z.literal(false),
              invalidReasons: z.array(z.string()),
            }),
          }),
        }),
      })
      .parse(parsePayload(result))
    expect(payload.mutationReceipt.renderedReadback.visibleSceneProof.invalidReasons).toContain(
      'Presented semantic mutation revision does not match the current authoritative scene.',
    )
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
