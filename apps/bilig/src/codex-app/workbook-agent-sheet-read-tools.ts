import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  WORKBOOK_AGENT_TOOL_NAMES,
  normalizeWorkbookAgentToolName,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolCallResult,
  type CodexDynamicToolSpec,
} from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { z } from 'zod'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { resolveWorkbookSelector } from './workbook-selector-resolver.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'

const sheetNameArgsSchema = z.object({
  sheetName: z.string().trim().min(1).optional(),
})

const currentRegionArgsSchema = z.object({
  sheetName: z.string().trim().min(1).optional(),
  address: z.string().trim().min(1).optional(),
})

const axisMetadataArgsSchema = z.object({
  sheetName: z.string().trim().min(1),
  startIndex: z.number().int().nonnegative().optional(),
  count: z.number().int().positive().optional(),
})

export const workbookAgentSheetReadToolSpecs = [
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.listSheets,
    description: 'List workbook sheets in visible sort order.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {},
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.getSheetView,
    description: 'Read sheet-level view state, including used range, freeze panes, filters, and sorts for one sheet.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.getUsedRange,
    description: 'Read the used range for one sheet.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.getCurrentRegion,
    description: 'Read the current contiguous populated region around an explicit anchor or the current selection.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string' },
        address: { type: 'string' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.getRowMetadata,
    description: 'Read row metadata regions for a sheet, with optional start/count filtering.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['sheetName'],
      properties: {
        sheetName: { type: 'string' },
        startIndex: { type: 'number' },
        count: { type: 'number' },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.getColumnMetadata,
    description: 'Read column metadata regions for a sheet, with optional start/count filtering.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['sheetName'],
      properties: {
        sheetName: { type: 'string' },
        startIndex: { type: 'number' },
        count: { type: 'number' },
      },
    },
  },
] satisfies readonly CodexDynamicToolSpec[]

export interface WorkbookAgentSheetReadToolContext {
  readonly documentId: string
  readonly session: SessionIdentity
  readonly uiContext: WorkbookAgentUiContext | null
  readonly zeroSyncService: ZeroSyncService
}

function textToolResult(text: string, success = true): CodexDynamicToolCallResult {
  return {
    success,
    contentItems: [
      {
        type: 'inputText',
        text,
      },
    ],
  }
}

function stringifyJson(value: unknown): string {
  return JSON.stringify(value, null, 2)
}

function resolveSheetName(sheetName: string | undefined, uiContext: WorkbookAgentUiContext | null): string {
  if (sheetName) {
    return sheetName
  }
  if (uiContext) {
    return uiContext.selection.sheetName
  }
  throw new Error('sheetName is required when no browser workbook context exists')
}

function getSheetUsedRange(
  runtime: WorkbookRuntime,
  sheetName: string,
): {
  startAddress: string
  endAddress: string
  rowCount: number
  columnCount: number
  cellCount: number
} | null {
  const sheet = runtime.engine.exportSnapshot().sheets.find((entry) => entry.name === sheetName)
  if (!sheet || sheet.cells.length === 0) {
    return null
  }
  let startRow = Number.POSITIVE_INFINITY
  let endRow = Number.NEGATIVE_INFINITY
  let startCol = Number.POSITIVE_INFINITY
  let endCol = Number.NEGATIVE_INFINITY
  for (const cell of sheet.cells) {
    const parsed = parseCellAddress(cell.address, sheetName)
    startRow = Math.min(startRow, parsed.row)
    endRow = Math.max(endRow, parsed.row)
    startCol = Math.min(startCol, parsed.col)
    endCol = Math.max(endCol, parsed.col)
  }
  return {
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    rowCount: endRow - startRow + 1,
    columnCount: endCol - startCol + 1,
    cellCount: (endRow - startRow + 1) * (endCol - startCol + 1),
  }
}

function filterAxisMetadata<T extends { start: number; count: number }>(
  entries: readonly T[],
  startIndex: number | undefined,
  count: number | undefined,
): readonly T[] {
  if (startIndex === undefined || count === undefined) {
    return entries
  }
  const endIndex = startIndex + count - 1
  return entries.filter((entry) => {
    const entryEnd = entry.start + entry.count - 1
    return entry.start <= endIndex && entryEnd >= startIndex
  })
}

export async function handleWorkbookAgentSheetReadToolCall(
  context: WorkbookAgentSheetReadToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult | null> {
  const normalizedTool = normalizeWorkbookAgentToolName(request.tool)
  switch (normalizedTool) {
    case WORKBOOK_AGENT_TOOL_NAMES.listSheets: {
      const payload = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
        documentId: context.documentId,
        sheetCount: runtime.engine.exportSnapshot().sheets.length,
        sheets: runtime.engine
          .exportSnapshot()
          .sheets.toSorted((left, right) => left.order - right.order)
          .map((sheet) => ({
            name: sheet.name,
            order: sheet.order,
            usedRange: getSheetUsedRange(runtime, sheet.name),
          })),
      }))
      return textToolResult(stringifyJson(payload))
    }
    case WORKBOOK_AGENT_TOOL_NAMES.getSheetView: {
      const args = sheetNameArgsSchema.parse(request.arguments)
      const payload = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const sheetName = resolveSheetName(args.sheetName, context.uiContext)
        return {
          documentId: context.documentId,
          sheetName,
          usedRange: getSheetUsedRange(runtime, sheetName),
          freezePane: runtime.engine.getFreezePane(sheetName) ?? null,
          filters: runtime.engine.getFilters(sheetName),
          sorts: runtime.engine.getSorts(sheetName),
        }
      })
      return textToolResult(stringifyJson(payload))
    }
    case WORKBOOK_AGENT_TOOL_NAMES.getUsedRange: {
      const args = sheetNameArgsSchema.parse(request.arguments)
      const payload = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const sheetName = resolveSheetName(args.sheetName, context.uiContext)
        return {
          documentId: context.documentId,
          sheetName,
          usedRange: getSheetUsedRange(runtime, sheetName),
        }
      })
      return textToolResult(stringifyJson(payload))
    }
    case WORKBOOK_AGENT_TOOL_NAMES.getCurrentRegion: {
      const args = currentRegionArgsSchema.parse(request.arguments)
      const payload = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const resolution = resolveWorkbookSelector({
          runtime,
          selector: {
            kind: 'currentRegion',
            ...(args.sheetName && args.address
              ? {
                  anchor: {
                    sheet: args.sheetName,
                    address: args.address,
                  },
                }
              : {}),
          },
          uiContext: context.uiContext,
        })
        return {
          documentId: context.documentId,
          sheetName: resolution.derivedA1Ranges[0]?.sheetName ?? null,
          resolvedSelector: {
            displayLabel: resolution.displayLabel,
            derivedA1Ranges: resolution.derivedA1Ranges,
          },
        }
      })
      return textToolResult(stringifyJson(payload))
    }
    case WORKBOOK_AGENT_TOOL_NAMES.getRowMetadata: {
      const args = axisMetadataArgsSchema.parse(request.arguments)
      const payload = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
        documentId: context.documentId,
        sheetName: args.sheetName,
        entryCount: runtime.engine.getRowMetadata(args.sheetName).length,
        entries: filterAxisMetadata(runtime.engine.getRowMetadata(args.sheetName), args.startIndex, args.count),
      }))
      return textToolResult(stringifyJson(payload))
    }
    case WORKBOOK_AGENT_TOOL_NAMES.getColumnMetadata: {
      const args = axisMetadataArgsSchema.parse(request.arguments)
      const payload = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
        documentId: context.documentId,
        sheetName: args.sheetName,
        entryCount: runtime.engine.getColumnMetadata(args.sheetName).length,
        entries: filterAxisMetadata(runtime.engine.getColumnMetadata(args.sheetName), args.startIndex, args.count),
      }))
      return textToolResult(stringifyJson(payload))
    }
    default:
      return null
  }
}
