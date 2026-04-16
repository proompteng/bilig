import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  WORKBOOK_AGENT_TOOL_NAMES,
  normalizeWorkbookAgentToolName,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolCallResult,
  type CodexDynamicToolSpec,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import type { WorkbookImageSnapshot, WorkbookShapeSnapshot } from '@bilig/protocol'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { z } from 'zod'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { rangeOrSelectorJsonSchema, rangeOrSelectorSchema, resolveRangeOrSelectorRequest } from './workbook-agent-selector-tooling.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'

const DEFAULT_MEDIA_ROWS = 8
const DEFAULT_MEDIA_COLS = 6

const shapeTypeSchema = z.enum(['rectangle', 'roundedRectangle', 'ellipse', 'line', 'arrow', 'textBox'])

const listMediaArgsSchema = z.object({
  sheetName: z.string().trim().min(1).optional(),
})

const mediaAnchorArgsSchema = z
  .object({
    sheetName: z.string().trim().min(1).optional(),
    address: z.string().trim().min(1).optional(),
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
    rows: z.number().int().positive().max(200).optional(),
    cols: z.number().int().positive().max(200).optional(),
  })
  .refine(
    (value) =>
      (value.sheetName ? 1 : 0) === (value.address ? 1 : 0) &&
      (value.range ? 1 : 0) + (value.selector ? 1 : 0) + (value.sheetName ? 1 : 0) <= 1,
    {
      message: 'Provide either sheetName/address, range, or selector',
    },
  )

const insertImageArgsSchema = mediaAnchorArgsSchema
  .extend({
    id: z.string().trim().min(1),
    sourceUrl: z.string().trim().min(1),
    altText: z.string().trim().min(1).optional(),
  })
  .refine(
    (value) => value.range !== undefined || value.selector !== undefined || (value.sheetName !== undefined && value.address !== undefined),
    {
      message: 'Provide an explicit anchor, range, or selector',
    },
  )

const moveImageArgsSchema = mediaAnchorArgsSchema
  .extend({
    id: z.string().trim().min(1),
    altText: z.string().trim().min(1).optional(),
  })
  .refine(
    (value) =>
      value.rows !== undefined ||
      value.cols !== undefined ||
      value.altText !== undefined ||
      value.range !== undefined ||
      value.selector !== undefined ||
      (value.sheetName !== undefined && value.address !== undefined),
    {
      message: 'Provide a new target location, size, or altText',
    },
  )

const deleteImageArgsSchema = z.object({
  id: z.string().trim().min(1),
})

const insertShapeArgsSchema = mediaAnchorArgsSchema
  .extend({
    id: z.string().trim().min(1),
    shapeType: shapeTypeSchema,
    text: z.string().trim().min(1).optional(),
    fillColor: z.string().trim().min(1).optional(),
    strokeColor: z.string().trim().min(1).optional(),
  })
  .refine(
    (value) => value.range !== undefined || value.selector !== undefined || (value.sheetName !== undefined && value.address !== undefined),
    {
      message: 'Provide an explicit anchor, range, or selector',
    },
  )

const updateShapeArgsSchema = mediaAnchorArgsSchema
  .extend({
    id: z.string().trim().min(1),
    shapeType: shapeTypeSchema.optional(),
    text: z.string().trim().min(1).optional(),
    fillColor: z.string().trim().min(1).optional(),
    strokeColor: z.string().trim().min(1).optional(),
  })
  .refine(
    (value) =>
      value.rows !== undefined ||
      value.cols !== undefined ||
      value.shapeType !== undefined ||
      value.text !== undefined ||
      value.fillColor !== undefined ||
      value.strokeColor !== undefined ||
      value.range !== undefined ||
      value.selector !== undefined ||
      (value.sheetName !== undefined && value.address !== undefined),
    {
      message: 'Provide at least one shape update',
    },
  )

const deleteShapeArgsSchema = z.object({
  id: z.string().trim().min(1),
})

export const workbookAgentMediaToolSpecs = [
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.listImages,
    description: 'List workbook images with anchor positions, source URLs, and footprint size.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.listShapes,
    description: 'List workbook shapes with anchor positions, shape types, and footprint size.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.insertImage,
    description: 'Insert an image at an explicit anchor or selector target, with optional footprint sizing and alt text.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['id', 'sourceUrl'],
      properties: {
        id: { type: 'string' },
        sourceUrl: { type: 'string' },
        altText: { type: 'string' },
        sheetName: { type: 'string' },
        address: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        rows: { type: 'number' },
        cols: { type: 'number' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.moveImage,
    description:
      'Move or resize an existing image by id. You can target an explicit anchor or selector target, and optionally update alt text.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['id'],
      properties: {
        id: { type: 'string' },
        altText: { type: 'string' },
        sheetName: { type: 'string' },
        address: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        rows: { type: 'number' },
        cols: { type: 'number' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteImage,
    description: 'Delete an image by id.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['id'],
      properties: {
        id: { type: 'string' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.insertShape,
    description: 'Insert a workbook shape at an explicit anchor or selector target, with optional text and colors.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['id', 'shapeType'],
      properties: {
        id: { type: 'string' },
        shapeType: { type: 'string', enum: shapeTypeSchema.options },
        text: { type: 'string' },
        fillColor: { type: 'string' },
        strokeColor: { type: 'string' },
        sheetName: { type: 'string' },
        address: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        rows: { type: 'number' },
        cols: { type: 'number' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.updateShape,
    description: 'Update an existing workbook shape by id, including anchor position, footprint, text, colors, or shape type.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['id'],
      properties: {
        id: { type: 'string' },
        shapeType: { type: 'string', enum: shapeTypeSchema.options },
        text: { type: 'string' },
        fillColor: { type: 'string' },
        strokeColor: { type: 'string' },
        sheetName: { type: 'string' },
        address: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        rows: { type: 'number' },
        cols: { type: 'number' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteShape,
    description: 'Delete a workbook shape by id.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['id'],
      properties: {
        id: { type: 'string' },
      },
    },
  },
] satisfies readonly CodexDynamicToolSpec[]

export interface WorkbookAgentMediaToolContext {
  readonly documentId: string
  readonly session: SessionIdentity
  readonly uiContext: WorkbookAgentUiContext | null
  readonly zeroSyncService: ZeroSyncService
  readonly stageCommand: (command: WorkbookAgentCommand) => Promise<
    | WorkbookAgentCommandBundle
    | {
        readonly bundle: WorkbookAgentCommandBundle
        readonly executionRecord: WorkbookAgentExecutionRecord | null
        readonly disposition?: 'queuedForTurnApply' | 'reviewQueued'
      }
  >
}

function stringifyJson(value: unknown): string {
  return JSON.stringify(value, null, 2)
}

function textToolResult(text: string, success = true): CodexDynamicToolCallResult {
  return {
    success,
    contentItems: [{ type: 'inputText', text }],
  }
}

async function stageCommandResult(
  context: WorkbookAgentMediaToolContext,
  command: WorkbookAgentCommand,
): Promise<CodexDynamicToolCallResult> {
  const result = await context.stageCommand(command)
  const normalized = 'bundle' in result ? result : { bundle: result, executionRecord: null, disposition: 'reviewQueued' as const }
  const bundle = normalized.bundle
  if (normalized.executionRecord) {
    return textToolResult(
      stringifyJson({
        applied: true,
        staged: false,
        reviewQueued: false,
        bundleId: bundle.id,
        summary: `Applied workbook change set at revision r${String(normalized.executionRecord.appliedRevision)}: ${normalized.executionRecord.summary}`,
        revision: normalized.executionRecord.appliedRevision,
        scope: normalized.executionRecord.scope,
        riskClass: normalized.executionRecord.riskClass,
        estimatedAffectedCells: bundle.estimatedAffectedCells,
        affectedRanges: bundle.affectedRanges,
      }),
    )
  }
  if (normalized.disposition === 'queuedForTurnApply') {
    return textToolResult(
      stringifyJson({
        applied: false,
        staged: true,
        reviewQueued: false,
        queuedForTurnApply: true,
        bundleId: bundle.id,
        summary: `Queued workbook change set for turn apply: ${bundle.summary}`,
        scope: bundle.scope,
        riskClass: bundle.riskClass,
        estimatedAffectedCells: bundle.estimatedAffectedCells,
        affectedRanges: bundle.affectedRanges,
      }),
    )
  }
  return textToolResult(
    stringifyJson({
      applied: false,
      staged: true,
      reviewQueued: true,
      bundleId: bundle.id,
      summary: `Prepared workbook review item: ${bundle.summary}`,
      scope: bundle.scope,
      riskClass: bundle.riskClass,
      estimatedAffectedCells: bundle.estimatedAffectedCells,
      affectedRanges: bundle.affectedRanges,
    }),
  )
}

function normalizeRangeFootprint(range: { readonly sheetName: string; readonly startAddress: string; readonly endAddress: string }) {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    rows: Math.abs(end.row - start.row) + 1,
    cols: Math.abs(end.col - start.col) + 1,
    address: formatAddress(Math.min(start.row, end.row), Math.min(start.col, end.col)),
  }
}

type MediaAnchorInput = z.infer<typeof mediaAnchorArgsSchema>

function resolvePlacement(
  runtime: WorkbookRuntime,
  args: MediaAnchorInput,
  uiContext: WorkbookAgentUiContext | null,
  fallback:
    | {
        sheetName: string
        address: string
        rows: number
        cols: number
      }
    | undefined,
) {
  if (args.range || args.selector) {
    const resolved = resolveRangeOrSelectorRequest({
      runtime,
      args: {
        ...(args.range ? { range: args.range } : {}),
        ...(args.selector ? { selector: args.selector } : {}),
      },
      uiContext,
    })
    const footprint = normalizeRangeFootprint(resolved.range)
    return {
      sheetName: resolved.range.sheetName,
      address: footprint.address,
      rows: args.rows ?? footprint.rows,
      cols: args.cols ?? footprint.cols,
    }
  }
  if (args.sheetName && args.address) {
    return {
      sheetName: args.sheetName,
      address: args.address,
      rows: args.rows ?? fallback?.rows ?? DEFAULT_MEDIA_ROWS,
      cols: args.cols ?? fallback?.cols ?? DEFAULT_MEDIA_COLS,
    }
  }
  if (fallback) {
    return {
      sheetName: fallback.sheetName,
      address: fallback.address,
      rows: args.rows ?? fallback.rows,
      cols: args.cols ?? fallback.cols,
    }
  }
  throw new Error('Provide an explicit anchor, range, or selector')
}

function listImages(runtime: WorkbookRuntime, sheetName?: string): readonly WorkbookImageSnapshot[] {
  return runtime.engine
    .getImages()
    .filter((image) => sheetName === undefined || image.sheetName === sheetName)
    .map((image) => structuredClone(image))
}

function listShapes(runtime: WorkbookRuntime, sheetName?: string): readonly WorkbookShapeSnapshot[] {
  return runtime.engine
    .getShapes()
    .filter((shape) => sheetName === undefined || shape.sheetName === sheetName)
    .map((shape) => structuredClone(shape))
}

function requireImage(runtime: WorkbookRuntime, id: string): WorkbookImageSnapshot {
  const image = runtime.engine.getImage(id)
  if (!image) {
    throw new Error(`Image ${id} does not exist`)
  }
  return image
}

function requireShape(runtime: WorkbookRuntime, id: string): WorkbookShapeSnapshot {
  const shape = runtime.engine.getShape(id)
  if (!shape) {
    throw new Error(`Shape ${id} does not exist`)
  }
  return shape
}

export async function handleWorkbookAgentMediaToolCall(
  context: WorkbookAgentMediaToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult | null> {
  const toolName = normalizeWorkbookAgentToolName(request.tool)
  switch (toolName) {
    case WORKBOOK_AGENT_TOOL_NAMES.listImages: {
      const args = listMediaArgsSchema.parse(request.arguments)
      const images = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => listImages(runtime, args.sheetName))
      return textToolResult(
        stringifyJson({
          imageCount: images.length,
          images,
        }),
      )
    }
    case WORKBOOK_AGENT_TOOL_NAMES.listShapes: {
      const args = listMediaArgsSchema.parse(request.arguments)
      const shapes = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => listShapes(runtime, args.sheetName))
      return textToolResult(
        stringifyJson({
          shapeCount: shapes.length,
          shapes,
        }),
      )
    }
    case WORKBOOK_AGENT_TOOL_NAMES.insertImage: {
      const args = insertImageArgsSchema.parse(request.arguments)
      const image = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const placement = resolvePlacement(runtime, args, context.uiContext, undefined)
        return {
          id: args.id,
          sheetName: placement.sheetName,
          address: placement.address,
          sourceUrl: args.sourceUrl,
          rows: placement.rows,
          cols: placement.cols,
          ...(args.altText ? { altText: args.altText } : {}),
        } satisfies WorkbookImageSnapshot
      })
      return await stageCommandResult(context, {
        kind: 'upsertImage',
        image,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.moveImage: {
      const args = moveImageArgsSchema.parse(request.arguments)
      const image = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const existing = requireImage(runtime, args.id)
        const placement = resolvePlacement(runtime, args, context.uiContext, existing)
        return {
          ...existing,
          sheetName: placement.sheetName,
          address: placement.address,
          rows: placement.rows,
          cols: placement.cols,
          ...(args.altText !== undefined ? { altText: args.altText } : {}),
        } satisfies WorkbookImageSnapshot
      })
      return await stageCommandResult(context, {
        kind: 'upsertImage',
        image,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteImage: {
      const args = deleteImageArgsSchema.parse(request.arguments)
      return await stageCommandResult(context, {
        kind: 'deleteImage',
        id: args.id,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.insertShape: {
      const args = insertShapeArgsSchema.parse(request.arguments)
      const shape = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const placement = resolvePlacement(runtime, args, context.uiContext, undefined)
        return {
          id: args.id,
          sheetName: placement.sheetName,
          address: placement.address,
          shapeType: args.shapeType,
          rows: placement.rows,
          cols: placement.cols,
          ...(args.text ? { text: args.text } : {}),
          ...(args.fillColor ? { fillColor: args.fillColor } : {}),
          ...(args.strokeColor ? { strokeColor: args.strokeColor } : {}),
        } satisfies WorkbookShapeSnapshot
      })
      return await stageCommandResult(context, {
        kind: 'upsertShape',
        shape,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.updateShape: {
      const args = updateShapeArgsSchema.parse(request.arguments)
      const shape = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const existing = requireShape(runtime, args.id)
        const placement = resolvePlacement(runtime, args, context.uiContext, existing)
        return {
          ...existing,
          sheetName: placement.sheetName,
          address: placement.address,
          rows: placement.rows,
          cols: placement.cols,
          ...(args.shapeType ? { shapeType: args.shapeType } : {}),
          ...(args.text !== undefined ? { text: args.text } : {}),
          ...(args.fillColor !== undefined ? { fillColor: args.fillColor } : {}),
          ...(args.strokeColor !== undefined ? { strokeColor: args.strokeColor } : {}),
        } satisfies WorkbookShapeSnapshot
      })
      return await stageCommandResult(context, {
        kind: 'upsertShape',
        shape,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteShape: {
      const args = deleteShapeArgsSchema.parse(request.arguments)
      return await stageCommandResult(context, {
        kind: 'deleteShape',
        id: args.id,
      })
    }
    default:
      return null
  }
}
