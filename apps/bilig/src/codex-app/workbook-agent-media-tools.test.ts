import { SpreadsheetEngine } from '@bilig/core'
import {
  createWorkbookAgentCommandBundle,
  WORKBOOK_AGENT_TOOL_NAMES,
  type CodexDynamicToolCallResult,
  type WorkbookAgentCommand,
} from '@bilig/agent-api'
import { describe, expect, it, vi } from 'vitest'
import type { AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'
import { z } from 'zod'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { handleWorkbookAgentToolCall } from './workbook-agent-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:media-test',
  })
  await engine.ready()
  engine.createSheet('Dashboard')
  engine.setImage({
    id: 'Revenue Image',
    sheetName: 'Dashboard',
    address: 'B2',
    sourceUrl: 'https://example.com/revenue.png',
    rows: 8,
    cols: 5,
    altText: 'Revenue image',
  })
  engine.setShape({
    id: 'Review Callout',
    sheetName: 'Dashboard',
    address: 'G4',
    shapeType: 'textBox',
    rows: 3,
    cols: 4,
    text: 'Review',
    fillColor: '#ffeeaa',
  })
  return engine
}

function createZeroSyncHarness(engine: SpreadsheetEngine) {
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
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: 1,
          calculatedRevision: 1,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-12T12:00:00.000Z',
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
      throw new Error('not used')
    },
    async listWorkbookChanges() {
      return []
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
      return 1
    },
    async loadAuthoritativeEvents() {
      return {
        afterRevision: 1,
        headRevision: 1,
        calculatedRevision: 1,
        events: [],
      } satisfies AuthoritativeWorkbookEventBatch
    },
  }
  return { zeroSyncService }
}

function createBundle(command: WorkbookAgentCommand) {
  return createWorkbookAgentCommandBundle({
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'media test',
    baseRevision: 1,
    now: 1,
    context: null,
    commands: [command],
  })
}

function parsePayload(result: CodexDynamicToolCallResult): unknown {
  expect(result.success).toBe(true)
  const item = result.contentItems[0]
  expect(item?.type).toBe('inputText')
  return JSON.parse(item && 'text' in item ? item.text : '')
}

const listImagesPayloadSchema = z.object({
  imageCount: z.number(),
  images: z.array(
    z.object({
      id: z.string(),
      sheetName: z.string(),
      address: z.string(),
      sourceUrl: z.string(),
    }),
  ),
})

const listShapesPayloadSchema = z.object({
  shapeCount: z.number(),
  shapes: z.array(
    z.object({
      id: z.string(),
      sheetName: z.string(),
      address: z.string(),
      shapeType: z.string(),
    }),
  ),
})

const stagedMediaPayloadSchema = z.object({
  staged: z.boolean(),
  bundleId: z.string(),
  affectedRanges: z.array(
    z.object({
      sheetName: z.string(),
      startAddress: z.string(),
      endAddress: z.string(),
      role: z.string(),
    }),
  ),
})

describe('workbook agent media tools', () => {
  it('lists workbook images and shapes from the authoritative runtime', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const imagesResult = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'deleteImage', id: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-list-images',
        tool: WORKBOOK_AGENT_TOOL_NAMES.listImages,
        arguments: {},
      },
    )
    const shapesResult = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'deleteShape', id: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-list-shapes',
        tool: WORKBOOK_AGENT_TOOL_NAMES.listShapes,
        arguments: {},
      },
    )

    const imagesPayload = listImagesPayloadSchema.parse(parsePayload(imagesResult))
    expect(imagesPayload.imageCount).toBe(1)
    expect(imagesPayload.images).toContainEqual(
      expect.objectContaining({
        id: 'Revenue Image',
        sheetName: 'Dashboard',
        address: 'B2',
      }),
    )

    const shapesPayload = listShapesPayloadSchema.parse(parsePayload(shapesResult))
    expect(shapesPayload.shapeCount).toBe(1)
    expect(shapesPayload.shapes).toContainEqual(
      expect.objectContaining({
        id: 'Review Callout',
        sheetName: 'Dashboard',
        address: 'G4',
        shapeType: 'textBox',
      }),
    )
  })

  it('stages insert-image, move-image, insert-shape, and update-shape commands', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    const insertImageResult = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-insert-image',
        tool: WORKBOOK_AGENT_TOOL_NAMES.insertImage,
        arguments: {
          id: 'Margin Image',
          sourceUrl: 'https://example.com/margin.png',
          sheetName: 'Dashboard',
          address: 'J2',
          rows: 7,
          cols: 4,
          altText: 'Margin image',
        },
      },
    )
    const moveImageResult = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-move-image',
        tool: WORKBOOK_AGENT_TOOL_NAMES.moveImage,
        arguments: {
          id: 'Revenue Image',
          sheetName: 'Dashboard',
          address: 'C3',
          rows: 9,
          cols: 6,
        },
      },
    )
    const insertShapeResult = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-insert-shape',
        tool: WORKBOOK_AGENT_TOOL_NAMES.insertShape,
        arguments: {
          id: 'Status Pill',
          shapeType: 'roundedRectangle',
          sheetName: 'Dashboard',
          address: 'L5',
          rows: 2,
          cols: 3,
          text: 'Ready',
          fillColor: '#d1fae5',
        },
      },
    )
    const updateShapeResult = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-update-shape',
        tool: WORKBOOK_AGENT_TOOL_NAMES.updateShape,
        arguments: {
          id: 'Review Callout',
          sheetName: 'Dashboard',
          address: 'H5',
          text: 'Approved',
          fillColor: '#d1fae5',
        },
      },
    )

    expect(stageCommand).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        kind: 'upsertImage',
        image: expect.objectContaining({
          id: 'Margin Image',
          sheetName: 'Dashboard',
          address: 'J2',
          rows: 7,
          cols: 4,
        }),
      }),
    )
    expect(stageCommand).toHaveBeenNthCalledWith(
      2,
      expect.objectContaining({
        kind: 'upsertImage',
        image: expect.objectContaining({
          id: 'Revenue Image',
          address: 'C3',
          rows: 9,
          cols: 6,
        }),
      }),
    )
    expect(stageCommand).toHaveBeenNthCalledWith(
      3,
      expect.objectContaining({
        kind: 'upsertShape',
        shape: expect.objectContaining({
          id: 'Status Pill',
          shapeType: 'roundedRectangle',
          address: 'L5',
        }),
      }),
    )
    expect(stageCommand).toHaveBeenNthCalledWith(
      4,
      expect.objectContaining({
        kind: 'upsertShape',
        shape: expect.objectContaining({
          id: 'Review Callout',
          address: 'H5',
          text: 'Approved',
          fillColor: '#d1fae5',
        }),
      }),
    )

    expect(stagedMediaPayloadSchema.parse(parsePayload(insertImageResult))).toEqual(
      expect.objectContaining({
        staged: true,
      }),
    )
    expect(stagedMediaPayloadSchema.parse(parsePayload(moveImageResult))).toEqual(
      expect.objectContaining({
        staged: true,
      }),
    )
    expect(stagedMediaPayloadSchema.parse(parsePayload(insertShapeResult))).toEqual(
      expect.objectContaining({
        staged: true,
      }),
    )
    expect(stagedMediaPayloadSchema.parse(parsePayload(updateShapeResult))).toEqual(
      expect.objectContaining({
        staged: true,
      }),
    )
  })

  it('stages delete-image and delete-shape commands', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    const deleteImageResult = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-delete-image',
        tool: WORKBOOK_AGENT_TOOL_NAMES.deleteImage,
        arguments: { id: 'Revenue Image' },
      },
    )
    const deleteShapeResult = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-delete-shape',
        tool: WORKBOOK_AGENT_TOOL_NAMES.deleteShape,
        arguments: { id: 'Review Callout' },
      },
    )

    expect(stageCommand).toHaveBeenNthCalledWith(1, {
      kind: 'deleteImage',
      id: 'Revenue Image',
    })
    expect(stageCommand).toHaveBeenNthCalledWith(2, {
      kind: 'deleteShape',
      id: 'Review Callout',
    })
    expect(stagedMediaPayloadSchema.parse(parsePayload(deleteImageResult))).toEqual(
      expect.objectContaining({
        staged: true,
      }),
    )
    expect(stagedMediaPayloadSchema.parse(parsePayload(deleteShapeResult))).toEqual(
      expect.objectContaining({
        staged: true,
      }),
    )
  })
})
