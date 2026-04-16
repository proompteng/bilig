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
import type {
  CellStylePatch,
  CellRangeRef,
  WorkbookConditionalFormatRuleSnapshot,
  WorkbookConditionalFormatSnapshot,
} from '@bilig/protocol'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { setRangeStyleArgsSchema } from '@bilig/zero-sync'
import { z } from 'zod'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { rangeOrSelectorJsonSchema, rangeOrSelectorSchema, resolveRangeOrSelectorRequest } from './workbook-agent-selector-tooling.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'

const literalValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()])

const conditionalFormatRuleSchema = z.union([
  z
    .object({
      kind: z.literal('cellIs'),
      operator: z.enum(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual']),
      values: z.array(literalValueSchema).min(1).max(2),
    })
    .refine((value) => (['between', 'notBetween'].includes(value.operator) ? value.values.length === 2 : value.values.length === 1), {
      message: 'between and notBetween require two values; other operators require one value',
    }),
  z.object({
    kind: z.literal('textContains'),
    text: z.string().trim().min(1),
    caseSensitive: z.boolean().optional(),
  }),
  z.object({
    kind: z.literal('formula'),
    formula: z.string().trim().min(1),
  }),
  z.object({
    kind: z.literal('blanks'),
  }),
  z.object({
    kind: z.literal('notBlanks'),
  }),
])
type ConditionalFormatRuleInput = z.infer<typeof conditionalFormatRuleSchema>
type ConditionalFormatStyleInput = Exclude<z.infer<typeof setRangeStyleArgsSchema.shape.patch>, undefined>

const listConditionalFormatsArgsSchema = z
  .object({
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) <= 1, {
    message: 'Provide at most one of range or selector',
  })

const upsertConditionalFormatArgsSchema = z
  .object({
    id: z.string().trim().min(1).optional(),
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
    rule: conditionalFormatRuleSchema,
    style: setRangeStyleArgsSchema.shape.patch,
    stopIfTrue: z.boolean().optional(),
    priority: z.number().int().positive().optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: 'Provide exactly one of range or selector',
  })

const updateConditionalFormatArgsSchema = z
  .object({
    id: z.string().trim().min(1),
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
    rule: conditionalFormatRuleSchema.optional(),
    style: setRangeStyleArgsSchema.shape.patch.optional(),
    stopIfTrue: z.boolean().optional(),
    priority: z.number().int().positive().nullable().optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) <= 1, {
    message: 'Provide at most one of range or selector',
  })
  .refine(
    (value) =>
      value.rule !== undefined ||
      value.style !== undefined ||
      value.stopIfTrue !== undefined ||
      value.priority !== undefined ||
      value.range !== undefined ||
      value.selector !== undefined,
    {
      message: 'Provide at least one update field',
    },
  )

const removeConditionalFormatArgsSchema = z.object({
  id: z.string().trim().min(1),
})

export const workbookAgentConditionalFormatToolSpecs = [
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.getConditionalFormats,
    description: 'List workbook conditional formatting rules. Optionally filter to an explicit range or semantic selector.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.addConditionalFormat,
    description: 'Add a conditional formatting rule to an explicit range or semantic selector with a style patch.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['rule', 'style'],
      properties: {
        id: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        rule: { type: 'object' },
        style: { type: 'object' },
        stopIfTrue: { type: 'boolean' },
        priority: { type: 'number' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.updateConditionalFormat,
    description: 'Update an existing conditional formatting rule by id. Optional range or selector retargets the rule.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['id'],
      properties: {
        id: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        rule: { type: 'object' },
        style: { type: 'object' },
        stopIfTrue: { type: 'boolean' },
        priority: { type: 'number' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.removeConditionalFormat,
    description: 'Remove an existing conditional formatting rule by id.',
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

export interface WorkbookAgentConditionalFormatToolContext {
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
  context: WorkbookAgentConditionalFormatToolContext,
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

function normalizeRangeBounds(range: CellRangeRef): CellRangeRef & {
  startRow: number
  endRow: number
  startCol: number
  endCol: number
} {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    ...range,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

function rangesIntersect(left: CellRangeRef, right: CellRangeRef): boolean {
  const leftBounds = normalizeRangeBounds(left)
  const rightBounds = normalizeRangeBounds(right)
  return !(
    leftBounds.sheetName !== rightBounds.sheetName ||
    leftBounds.endRow < rightBounds.startRow ||
    rightBounds.endRow < leftBounds.startRow ||
    leftBounds.endCol < rightBounds.startCol ||
    rightBounds.endCol < leftBounds.startCol
  )
}

function listConditionalFormats(runtime: WorkbookRuntime) {
  return runtime.engine.exportSnapshot().sheets.flatMap((sheet) => runtime.engine.getConditionalFormats(sheet.name))
}

function normalizeConditionalFormatRule(rule: ConditionalFormatRuleInput): WorkbookConditionalFormatRuleSnapshot {
  switch (rule.kind) {
    case 'cellIs':
      return {
        kind: 'cellIs',
        operator: rule.operator,
        values: [...rule.values],
      }
    case 'textContains':
      return {
        kind: 'textContains',
        text: rule.text,
        ...(rule.caseSensitive !== undefined ? { caseSensitive: rule.caseSensitive } : {}),
      }
    case 'formula':
      return {
        kind: 'formula',
        formula: rule.formula,
      }
    case 'blanks':
      return { kind: 'blanks' }
    case 'notBlanks':
      return { kind: 'notBlanks' }
  }
}

function normalizeConditionalFormatStylePatch(style: ConditionalFormatStyleInput): CellStylePatch {
  const normalized: CellStylePatch = {}
  if (style.fill !== undefined) {
    if (style.fill === null) {
      normalized.fill = null
    } else {
      const fill: NonNullable<CellStylePatch['fill']> = {}
      if (style.fill.backgroundColor !== undefined) {
        fill.backgroundColor = style.fill.backgroundColor
      }
      normalized.fill = fill
    }
  }
  if (style.font !== undefined) {
    normalized.font =
      style.font === null
        ? null
        : {
            ...(style.font.family !== undefined ? { family: style.font.family } : {}),
            ...(style.font.size !== undefined ? { size: style.font.size } : {}),
            ...(style.font.bold !== undefined ? { bold: style.font.bold } : {}),
            ...(style.font.italic !== undefined ? { italic: style.font.italic } : {}),
            ...(style.font.underline !== undefined ? { underline: style.font.underline } : {}),
            ...(style.font.color !== undefined ? { color: style.font.color } : {}),
          }
  }
  if (style.alignment !== undefined) {
    normalized.alignment =
      style.alignment === null
        ? null
        : {
            ...(style.alignment.horizontal !== undefined ? { horizontal: style.alignment.horizontal } : {}),
            ...(style.alignment.vertical !== undefined ? { vertical: style.alignment.vertical } : {}),
            ...(style.alignment.wrap !== undefined ? { wrap: style.alignment.wrap } : {}),
            ...(style.alignment.indent !== undefined ? { indent: style.alignment.indent } : {}),
          }
  }
  if (style.borders !== undefined) {
    if (style.borders === null) {
      normalized.borders = null
    } else {
      const borders: NonNullable<CellStylePatch['borders']> = {}
      for (const sideName of ['top', 'right', 'bottom', 'left'] as const) {
        const side = style.borders[sideName]
        if (side === undefined) {
          continue
        }
        if (side === null) {
          borders[sideName] = null
          continue
        }
        borders[sideName] = {
          ...(side.style !== undefined ? { style: side.style } : {}),
          ...(side.weight !== undefined ? { weight: side.weight } : {}),
          ...(side.color !== undefined ? { color: side.color } : {}),
        }
      }
      normalized.borders = borders
    }
  }
  return normalized
}

export async function handleWorkbookAgentConditionalFormatToolCall(
  context: WorkbookAgentConditionalFormatToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult | null> {
  const normalizedTool = normalizeWorkbookAgentToolName(request.tool)
  switch (normalizedTool) {
    case WORKBOOK_AGENT_TOOL_NAMES.getConditionalFormats: {
      const args = listConditionalFormatsArgsSchema.parse(request.arguments)
      const payload = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const formats = listConditionalFormats(runtime)
        if (!args.range && !args.selector) {
          return {
            documentId: context.documentId,
            conditionalFormatCount: formats.length,
            conditionalFormats: formats,
          }
        }
        const resolved = resolveRangeOrSelectorRequest({
          runtime,
          args: {
            ...(args.range ? { range: args.range } : {}),
            ...(args.selector ? { selector: args.selector } : {}),
          },
          uiContext: context.uiContext,
        })
        return {
          documentId: context.documentId,
          conditionalFormatCount: formats.filter((format) => rangesIntersect(format.range, resolved.range)).length,
          conditionalFormats: formats.filter((format) => rangesIntersect(format.range, resolved.range)),
        }
      })
      return textToolResult(stringifyJson(payload))
    }
    case WORKBOOK_AGENT_TOOL_NAMES.addConditionalFormat: {
      const args = upsertConditionalFormatArgsSchema.parse(request.arguments)
      const format = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        if (args.style === undefined) {
          throw new Error('style is required')
        }
        const resolved = resolveRangeOrSelectorRequest({
          runtime,
          args: {
            ...(args.range ? { range: args.range } : {}),
            ...(args.selector ? { selector: args.selector } : {}),
          },
          uiContext: context.uiContext,
        })
        return {
          id: args.id ?? crypto.randomUUID(),
          range: resolved.range,
          rule: normalizeConditionalFormatRule(structuredClone(args.rule)),
          style: normalizeConditionalFormatStylePatch(structuredClone(args.style)),
          ...(args.stopIfTrue !== undefined ? { stopIfTrue: args.stopIfTrue } : {}),
          ...(args.priority !== undefined ? { priority: args.priority } : {}),
        } satisfies WorkbookConditionalFormatSnapshot
      })
      return await stageCommandResult(context, {
        kind: 'upsertConditionalFormat',
        format,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.updateConditionalFormat: {
      const args = updateConditionalFormatArgsSchema.parse(request.arguments)
      const format = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const existing = runtime.engine.getConditionalFormat(args.id)
        if (!existing) {
          throw new Error(`Conditional format ${args.id} does not exist`)
        }
        const resolvedRange =
          args.range || args.selector
            ? resolveRangeOrSelectorRequest({
                runtime,
                args: {
                  ...(args.range ? { range: args.range } : {}),
                  ...(args.selector ? { selector: args.selector } : {}),
                },
                uiContext: context.uiContext,
              }).range
            : existing.range
        return {
          ...existing,
          range: resolvedRange,
          ...(args.rule !== undefined
            ? {
                rule: normalizeConditionalFormatRule(structuredClone(args.rule)),
              }
            : {}),
          ...(args.style !== undefined
            ? {
                style: normalizeConditionalFormatStylePatch(structuredClone(args.style)),
              }
            : {}),
          ...(args.stopIfTrue !== undefined ? { stopIfTrue: args.stopIfTrue } : {}),
          ...(args.priority !== undefined ? (args.priority === null ? {} : { priority: args.priority }) : {}),
        } satisfies WorkbookConditionalFormatSnapshot
      })
      return await stageCommandResult(context, {
        kind: 'upsertConditionalFormat',
        format,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.removeConditionalFormat: {
      const args = removeConditionalFormatArgsSchema.parse(request.arguments)
      const existing = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const format = runtime.engine.getConditionalFormat(args.id)
        if (!format) {
          throw new Error(`Conditional format ${args.id} does not exist`)
        }
        return structuredClone(format)
      })
      return await stageCommandResult(context, {
        kind: 'deleteConditionalFormat',
        id: existing.id,
        range: existing.range,
      })
    }
    default:
      return null
  }
}
