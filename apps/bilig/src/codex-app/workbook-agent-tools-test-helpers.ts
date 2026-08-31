import { createWorkbookAgentCommandBundle, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { SpreadsheetEngine } from '@bilig/core'
import { expect } from 'vitest'
import { z } from 'zod'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { handleWorkbookAgentToolCall } from './workbook-agent-tools.js'
export type { JsonValue, WorkbookAgentCommandBundle } from '@bilig/agent-api'
export type { WorkbookAgentUiContext } from '@bilig/contracts'
export { describe, expect, it, vi } from 'vitest'
export { z } from 'zod'
export type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
export type { ZeroSyncService } from '../zero/service.js'
export { applyWorkbookAgentCommandBundleWithUndoCapture } from '../zero/workbook-agent-apply.js'
export type { WorkbookChangeRecord } from '../zero/workbook-change-store.js'
export { handleWorkbookAgentToolCall } from './workbook-agent-tools.js'

export async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 42)
  engine.setCellValue('Sheet1', 'A2', 'Gross Margin')
  engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A1)')
  engine.setCellFormula('Sheet1', 'C1', '1/0')
  engine.setCellFormula('Sheet1', 'D1', 'LEN(A1:A2)')
  engine.setRangeStyle(
    { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
    {
      fill: { backgroundColor: '#fef3c7' },
      font: { bold: true },
    },
  )
  engine.setRangeNumberFormat(
    { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
    {
      kind: 'currency',
      currency: 'USD',
      decimals: 2,
      useGrouping: true,
      negativeStyle: 'minus',
      zeroStyle: 'zero',
    },
  )
  engine.createSheet('Ops Search')
  engine.setCellValue('Ops Search', 'A1', 'Northwind Import')
  return engine
}

export function createVisibleSceneProof(
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
  }
}

export function createZeroSyncHarness(engine: SpreadsheetEngine) {
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
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: 1,
          calculatedRevision: 1,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-06T12:00:00.000Z',
        }),
        headRevision: 1,
        calculatedRevision: 1,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    },
    async applyServerMutator() {
      throw new Error('not used')
    },
    async applyAgentCommandBundle() {
      return {
        revision: 1,
        preview: {
          ranges: [],
          structuralChanges: [],
          cellDiffs: [],
          effectSummary: {
            displayedCellDiffCount: 0,
            truncatedCellDiffs: false,
            inputChangeCount: 0,
            formulaChangeCount: 0,
            styleChangeCount: 0,
            numberFormatChangeCount: 0,
            structuralChangeCount: 0,
          },
        },
      }
    },
    async applyWorkbookPlanData() {
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
      return 1
    },
    async loadAuthoritativeEvents() {
      throw new Error('not used')
    },
  }
  return { zeroSyncService }
}

export function createBundle(command: WorkbookAgentCommandBundle['commands'][number]): WorkbookAgentCommandBundle {
  return {
    id: 'bundle-1',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'Update cells',
    summary: 'Stage workbook changes',
    scope: 'selection',
    riskClass: 'low',
    baseRevision: 1,
    createdAtUnixMs: 1,
    context: null,
    commands: [command],
    affectedRanges: [],
    estimatedAffectedCells: 2,
  }
}

export function createDerivedBundle(command: WorkbookAgentCommandBundle['commands'][number]): WorkbookAgentCommandBundle {
  return createWorkbookAgentCommandBundle({
    bundleId: 'bundle-derived',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'Update cells',
    baseRevision: 1,
    context: null,
    commands: [command],
    now: 1,
  })
}

export function readToolJson(response: Awaited<ReturnType<typeof handleWorkbookAgentToolCall>>): unknown {
  const text = response.contentItems[0]
  expect(text?.type).toBe('inputText')
  return JSON.parse(text && 'text' in text ? text.text : '')
}

export const tableListPayloadSchema = z.object({
  tables: z.array(
    z.object({
      sheetName: z.string(),
    }),
  ),
})

export const invariantPayloadSchema = z.object({
  summary: z.object({
    ok: z.boolean(),
  }),
  problems: z.array(
    z.object({
      message: z.string(),
    }),
  ),
})

export const renderedSelectionPayloadSchema = z.object({
  authoritativeReadback: z.object({
    rows: z.array(z.array(z.object({ value: z.unknown() }))),
  }),
  renderedReadback: z.object({
    available: z.boolean(),
    capturedRevision: z.number().nullable(),
    capturedBatchId: z.number().nullable(),
    range: z.object({
      rows: z.array(z.array(z.object({ style: z.unknown(), input: z.unknown() }))),
    }),
  }),
})

export const workbookSummarySchema = z.object({
  summary: z.object({
    sheetCount: z.number(),
    tableCount: z.number(),
    pivotCount: z.number(),
    spillCount: z.number(),
    hiddenRowIndexCount: z.number(),
    hiddenColumnIndexCount: z.number(),
  }),
  sheets: z.array(
    z.object({
      name: z.string(),
      freezePane: z
        .object({
          rows: z.number(),
          cols: z.number(),
        })
        .nullable(),
      filterCount: z.number(),
      sortCount: z.number(),
      tableCount: z.number(),
      pivotCount: z.number(),
      spillCount: z.number(),
      rowMetadata: z.object({
        hiddenIndexCount: z.number(),
        explicitSizeIndexCount: z.number(),
      }),
      columnMetadata: z.object({
        hiddenIndexCount: z.number(),
        explicitSizeIndexCount: z.number(),
      }),
    }),
  ),
})
