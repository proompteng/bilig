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
import type { CellRangeRef, WorkbookRangeProtectionSnapshot, WorkbookSheetProtectionSnapshot } from '@bilig/protocol'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { z } from 'zod'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { rangeOrSelectorJsonSchema, rangeOrSelectorSchema, resolveRangeOrSelectorRequest } from './workbook-agent-selector-tooling.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'

const getProtectionStatusArgsSchema = z.object({
  sheetName: z.string().trim().min(1).optional(),
  range: rangeOrSelectorSchema.shape.range.optional(),
  selector: rangeOrSelectorSchema.shape.selector.optional(),
})

const sheetProtectionArgsSchema = z.object({
  sheetName: z.string().trim().min(1).optional(),
  hideFormulas: z.boolean().optional(),
})

const rangeProtectionArgsSchema = z
  .object({
    id: z.string().trim().min(1).optional(),
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
    hideFormulas: z.boolean().optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: 'Provide exactly one of range or selector',
  })

const deleteRangeProtectionArgsSchema = z
  .object({
    id: z.string().trim().min(1).optional(),
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
  })
  .refine((value) => value.id !== undefined || value.range !== undefined || value.selector !== undefined, {
    message: 'Provide an id or one of range or selector',
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) <= 1, {
    message: 'Provide at most one of range or selector',
  })

const hideFormulasArgsSchema = z
  .object({
    sheetName: z.string().trim().min(1).optional(),
    id: z.string().trim().min(1).optional(),
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
  })
  .refine((value) => (value.sheetName ? 1 : 0) + (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: 'Provide exactly one of sheetName, range, or selector',
  })

export const workbookAgentProtectionToolSpecs = [
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.getProtectionStatus,
    description: 'List workbook sheet and range protection status. Optionally filter by sheet, explicit range, or semantic selector.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.protectSheet,
    description: 'Protect a sheet, optionally hiding formulas across the entire sheet.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
        hideFormulas: { type: 'boolean' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.unprotectSheet,
    description: 'Remove sheet protection from a sheet.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.protectRange,
    description: 'Protect a range or selector target, optionally hiding formulas in that region.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        id: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        hideFormulas: { type: 'boolean' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.unprotectRange,
    description: 'Remove a range protection by id or exact target range.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        id: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.lockCells,
    description: 'Alias for protecting a range of cells.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        id: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        hideFormulas: { type: 'boolean' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.unlockCells,
    description: 'Alias for removing a range protection.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        id: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.hideFormulas,
    description: 'Hide formulas on a protected sheet or protected range. For ranges, creates or updates a protection record.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
        id: { type: 'string' },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
      },
    },
  },
] satisfies readonly CodexDynamicToolSpec[]

export interface WorkbookAgentProtectionToolContext {
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
  context: WorkbookAgentProtectionToolContext,
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
  const a = normalizeRangeBounds(left)
  const b = normalizeRangeBounds(right)
  return !(a.sheetName !== b.sheetName || a.endRow < b.startRow || b.endRow < a.startRow || a.endCol < b.startCol || b.endCol < a.startCol)
}

function listProtectionStatus(runtime: WorkbookRuntime) {
  return runtime.engine.exportSnapshot().sheets.map((sheet) => ({
    sheetName: sheet.name,
    sheetProtection: runtime.engine.getSheetProtection(sheet.name) ?? null,
    rangeProtections: runtime.engine.getRangeProtections(sheet.name),
  }))
}

function resolveSheetName(args: { sheetName?: string | undefined }, uiContext: WorkbookAgentUiContext | null): string {
  if (args.sheetName) {
    return args.sheetName
  }
  if (uiContext) {
    return uiContext.selection.sheetName
  }
  throw new Error('sheetName is required when there is no attached workbook context')
}

export async function handleWorkbookAgentProtectionToolCall(
  context: WorkbookAgentProtectionToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult | null> {
  const normalizedTool = normalizeWorkbookAgentToolName(request.tool)
  switch (normalizedTool) {
    case WORKBOOK_AGENT_TOOL_NAMES.getProtectionStatus: {
      const args = getProtectionStatusArgsSchema.parse(request.arguments)
      const payload = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const status = listProtectionStatus(runtime)
        if (args.sheetName) {
          return status.find((entry) => entry.sheetName === args.sheetName) ?? null
        }
        if (args.range || args.selector) {
          const resolved = resolveRangeOrSelectorRequest({
            runtime,
            args: {
              ...(args.range ? { range: args.range } : {}),
              ...(args.selector ? { selector: args.selector } : {}),
            },
            uiContext: context.uiContext,
          })
          return {
            sheetName: resolved.range.sheetName,
            sheetProtection: runtime.engine.getSheetProtection(resolved.range.sheetName) ?? null,
            rangeProtections: runtime.engine
              .getRangeProtections(resolved.range.sheetName)
              .filter((entry) => rangesIntersect(entry.range, resolved.range)),
          }
        }
        return status
      })
      return textToolResult(stringifyJson(payload))
    }
    case WORKBOOK_AGENT_TOOL_NAMES.protectSheet: {
      const args = sheetProtectionArgsSchema.parse(request.arguments)
      return await stageCommandResult(context, {
        kind: 'setSheetProtection',
        protection: {
          sheetName: resolveSheetName(args, context.uiContext),
          ...(args.hideFormulas !== undefined ? { hideFormulas: args.hideFormulas } : {}),
        },
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.unprotectSheet: {
      const args = sheetProtectionArgsSchema.parse(request.arguments)
      return await stageCommandResult(context, {
        kind: 'clearSheetProtection',
        sheetName: resolveSheetName(args, context.uiContext),
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.protectRange:
    case WORKBOOK_AGENT_TOOL_NAMES.lockCells: {
      const args = rangeProtectionArgsSchema.parse(request.arguments)
      const protection = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
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
          ...(args.hideFormulas !== undefined ? { hideFormulas: args.hideFormulas } : {}),
        } satisfies WorkbookRangeProtectionSnapshot
      })
      return await stageCommandResult(context, {
        kind: 'upsertRangeProtection',
        protection,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.unprotectRange:
    case WORKBOOK_AGENT_TOOL_NAMES.unlockCells: {
      const args = deleteRangeProtectionArgsSchema.parse(request.arguments)
      const protection = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        if (args.id) {
          const existing = runtime.engine.getRangeProtection(args.id)
          if (!existing) {
            throw new Error(`Range protection ${args.id} does not exist`)
          }
          return existing
        }
        const resolved = resolveRangeOrSelectorRequest({
          runtime,
          args: {
            ...(args.range ? { range: args.range } : {}),
            ...(args.selector ? { selector: args.selector } : {}),
          },
          uiContext: context.uiContext,
        })
        const normalized = normalizeRangeBounds(resolved.range)
        const matches = runtime.engine.getRangeProtections(normalized.sheetName).filter((entry) => {
          const candidate = normalizeRangeBounds(entry.range)
          return candidate.startAddress === normalized.startAddress && candidate.endAddress === normalized.endAddress
        })
        if (matches.length !== 1) {
          throw new Error(
            matches.length === 0
              ? 'No range protection matches that target'
              : 'Multiple range protections match that target; provide id explicitly',
          )
        }
        return matches[0]!
      })
      return await stageCommandResult(context, {
        kind: 'deleteRangeProtection',
        id: protection.id,
        range: protection.range,
      })
    }
    case WORKBOOK_AGENT_TOOL_NAMES.hideFormulas: {
      const args = hideFormulasArgsSchema.parse(request.arguments)
      if (args.sheetName) {
        const protection = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const existing = runtime.engine.getSheetProtection(args.sheetName!)
          return {
            sheetName: args.sheetName!,
            hideFormulas: true,
            ...(existing ? {} : {}),
          } satisfies WorkbookSheetProtectionSnapshot
        })
        return await stageCommandResult(context, {
          kind: 'setSheetProtection',
          protection,
        })
      }
      const protection = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const resolved = resolveRangeOrSelectorRequest({
          runtime,
          args: {
            ...(args.range ? { range: args.range } : {}),
            ...(args.selector ? { selector: args.selector } : {}),
          },
          uiContext: context.uiContext,
        })
        if (args.id) {
          const existing = runtime.engine.getRangeProtection(args.id)
          if (existing) {
            return {
              ...existing,
              hideFormulas: true,
              range: resolved.range,
            } satisfies WorkbookRangeProtectionSnapshot
          }
        }
        return {
          id: args.id ?? crypto.randomUUID(),
          range: resolved.range,
          hideFormulas: true,
        } satisfies WorkbookRangeProtectionSnapshot
      })
      return await stageCommandResult(context, {
        kind: 'upsertRangeProtection',
        protection,
      })
    }
    default:
      return null
  }
}
