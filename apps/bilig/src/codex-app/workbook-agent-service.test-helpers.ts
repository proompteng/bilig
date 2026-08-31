import {
  isWorkbookAgentReviewQueueItem,
  toAppliedWorkbookCommandResult,
  toWorkbookAgentCommandBundle,
  toWorkbookAgentReviewQueueItem,
  type CodexDynamicToolCallResult,
  type CodexServerNotification,
  type CodexThread,
  type WorkbookAgentCommandBundle,
  type CodexTurn,
} from '@bilig/agent-api'
import type { WorkbookAgentUiContext, WorkbookAgentWorkflowRun } from '@bilig/contracts'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import { expect, vi } from 'vitest'
import type { ZeroSyncService } from '../zero/service.js'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import type { CodexAppServerTransport } from './codex-app-server-client.js'
import type { WorkbookAgentService } from './workbook-agent-service.js'

export class FakeCodexTransport implements CodexAppServerTransport {
  private readonly listeners = new Set<(notification: CodexServerNotification) => void>()
  private turnCounter = 0
  private threadCounter = 0
  lastThreadStartInput: Parameters<CodexAppServerTransport['threadStart']>[0] | null = null
  lastThreadResumeInput: { threadId: string } | null = null
  resumeError: unknown = null
  uniqueThreadStart = false
  nextTurn: Promise<CodexTurn> | null = null
  resumedThread: CodexThread | null = null
  closeCount = 0

  async ensureReady() {
    return {
      userAgent: 'fake',
      codexHome: '/tmp/fake-codex',
      platformFamily: 'unix',
      platformOs: 'macos',
    }
  }

  subscribe(listener: (notification: CodexServerNotification) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  async threadStart(input: Parameters<CodexAppServerTransport['threadStart']>[0]) {
    this.lastThreadStartInput = input
    this.threadCounter += 1
    return {
      id: this.uniqueThreadStart ? `thr-test-${String(this.threadCounter)}` : 'thr-test',
      preview: '',
      turns: [],
    }
  }

  async threadResume(input: { threadId: string }) {
    this.lastThreadResumeInput = input
    if (this.resumeError) {
      throw this.resumeError
    }
    if (this.resumedThread) {
      return {
        ...this.resumedThread,
        id: input.threadId,
      }
    }
    return {
      id: input.threadId,
      preview: '',
      turns: [],
    }
  }

  async turnStart(): Promise<CodexTurn> {
    if (this.nextTurn) {
      return await this.nextTurn
    }
    this.turnCounter += 1
    return {
      id: `turn-${String(this.turnCounter)}`,
      status: 'inProgress',
      items: [],
      error: null,
    }
  }

  async turnInterrupt() {}

  async close() {
    this.closeCount += 1
  }

  emit(notification: CodexServerNotification): void {
    this.listeners.forEach((listener) => listener(notification))
  }
}

export function createPreviewSummary(overrides: Record<string, unknown> = {}) {
  return {
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
    ...overrides,
  }
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

export function createRenderedContextForServiceTest(input: {
  readonly capturedRevision: number
  readonly stringId: number
  readonly value: string
}): WorkbookAgentUiContext {
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
      rowEnd: 20,
      colStart: 0,
      colEnd: 10,
    },
    rendered: {
      capturedAtUnixMs: input.capturedRevision * 100,
      capturedRevision: input.capturedRevision,
      batchId: input.capturedRevision,
      visibleSceneProof: createVisibleSceneProof(input.capturedRevision),
      selection: null,
      visibleRange: {
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
              input: input.value,
              value: {
                tag: ValueTag.String,
                value: input.value,
                stringId: input.stringId,
              },
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
}

export function createDurableRunningWorkflowRun(): WorkbookAgentWorkflowRun {
  return {
    runId: 'workflow-existing',
    threadId: 'thr-existing',
    startedByUserId: 'alex@example.com',
    workflowTemplate: 'summarizeWorkbook',
    title: 'Summarize Workbook',
    summary: 'Running workbook summary workflow.',
    status: 'running',
    createdAtUnixMs: 100,
    updatedAtUnixMs: 100,
    completedAtUnixMs: null,
    errorMessage: null,
    steps: [
      {
        stepId: 'inspect-workbook',
        label: 'Inspect workbook structure',
        status: 'running',
        summary: 'Reading durable workbook structure and layout metadata.',
        updatedAtUnixMs: 100,
      },
    ],
    artifact: null,
  }
}

export function createReviewQueueItem(bundle: WorkbookAgentCommandBundle) {
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: bundle.sharedReview ? 'ownerReview' : 'manual',
    ...(bundle.sharedReview ? { sharedReview: bundle.sharedReview } : {}),
  })
}

export function getPrimaryReviewBundle(snapshot: { reviewQueueItems: readonly unknown[] }): WorkbookAgentCommandBundle | null {
  const [reviewItem] = snapshot.reviewQueueItems
  if (!isWorkbookAgentReviewQueueItem(reviewItem)) {
    return null
  }
  return toWorkbookAgentCommandBundle(reviewItem)
}

export function readDynamicToolJson(result: CodexDynamicToolCallResult | undefined): Record<string, unknown> {
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

function isUnknownRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

type ApplyAgentCommandBundleResult = Awaited<ReturnType<ZeroSyncService['applyAgentCommandBundle']>>

function withAppliedCommandResultProof(
  bundle: WorkbookAgentCommandBundle,
  result: ApplyAgentCommandBundleResult,
): ApplyAgentCommandBundleResult {
  if (result.commandResult !== undefined) {
    return result
  }
  return {
    ...result,
    commandResult: toAppliedWorkbookCommandResult({
      bundle,
      revision: result.revision,
    }),
  }
}

async function createWorkbookRuntimeStub(documentId = 'doc-1'): Promise<WorkbookRuntime> {
  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `server:${documentId}:test`,
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 42)
  return {
    documentId,
    engine,
    projection: buildWorkbookSourceProjectionFromEngine(documentId, engine, {
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
}

export function createZeroSyncStub(overrides: Partial<ZeroSyncService> = {}): ZeroSyncService {
  const applyAgentCommandBundleOverride = overrides.applyAgentCommandBundle
  const normalizedOverrides = { ...overrides }
  if (applyAgentCommandBundleOverride !== undefined) {
    normalizedOverrides.applyAgentCommandBundle = async (...args) => {
      const result = await applyAgentCommandBundleOverride(...args)
      return withAppliedCommandResultProof(args[1], result)
    }
  }

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
    async inspectWorkbook<T>(documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>) {
      return await task(await createWorkbookRuntimeStub(documentId))
    },
    async applyServerMutator() {},
    async applyAgentCommandBundle(_documentId, bundle) {
      const revision = Math.max(2, bundle.baseRevision + 1)
      return withAppliedCommandResultProof(bundle, { revision, preview: createPreviewSummary() })
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
      return 1
    },
    async loadAuthoritativeEvents() {
      throw new Error('not used')
    },
    ...normalizedOverrides,
  }
}

export async function waitForWorkflowStatus(
  service: WorkbookAgentService,
  threadId: string,
  userId: string,
  status: 'running' | 'completed' | 'failed' | 'cancelled',
): Promise<ReturnType<WorkbookAgentService['getSnapshot']>> {
  await vi.waitFor(() => {
    const workflowRun = service.getSnapshot({
      documentId: 'doc-1',
      threadId,
      session: {
        userID: userId,
        roles: ['editor'],
      },
    }).workflowRuns[0]
    if (workflowRun?.status === 'failed' && status !== 'failed') {
      throw new Error(workflowRun.errorMessage ?? 'Workflow failed')
    }
    expect(workflowRun?.status).toBe(status)
  })
  return service.getSnapshot({
    documentId: 'doc-1',
    threadId,
    session: {
      userID: userId,
      roles: ['editor'],
    },
  })
}

export async function startWorkbookAgentTestTurn(
  service: WorkbookAgentService,
  input: {
    readonly threadId: string
    readonly userId?: string
    readonly prompt?: string
  },
): Promise<void> {
  await service.startTurn({
    documentId: 'doc-1',
    threadId: input.threadId,
    session: {
      userID: input.userId ?? 'alex@example.com',
      roles: ['editor'],
    },
    body: {
      prompt: input.prompt ?? 'Run workbook tool',
    },
  })
}
