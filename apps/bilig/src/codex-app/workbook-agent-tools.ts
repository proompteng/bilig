import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  type CellRangeRef,
  type CellNumberFormatInput,
  type CellNumberFormatPreset,
  normalizeCellNumberFormatPreset,
} from '@bilig/protocol'
import { WORKBOOK_AGENT_TOOL_NAMES, normalizeWorkbookAgentToolName } from '@bilig/agent-api'
import type {
  CodexDynamicToolCallRequest,
  CodexDynamicToolCallResult,
  CodexDynamicToolSpec,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
  WorkbookAgentExecutionRecord,
  WorkbookAgentWriteCellInput,
} from '@bilig/agent-api'
import { setRangeNumberFormatArgsSchema } from '@bilig/zero-sync'
import type { WorkbookAgentUiContext, WorkbookAgentWorkflowRun, WorkbookViewport } from '@bilig/contracts'
import { z } from 'zod'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import {
  findWorkbookFormulaIssues,
  searchWorkbook,
  summarizeWorkbookStructure,
  traceWorkbookDependencies,
} from './workbook-agent-comprehension.js'
import {
  inspectWorkbookCell,
  inspectWorkbookContext,
  inspectWorkbookRange,
  normalizeWorkbookAgentUiContext,
} from './workbook-agent-inspection.js'
import { verifyWorkbookInvariants } from './workbook-agent-audit.js'
import { handleWorkbookAgentAnnotationToolCall, workbookAgentAnnotationToolSpecs } from './workbook-agent-annotation-tools.js'
import { handleWorkbookAgentAuditToolCall, workbookAgentAuditToolSpecs } from './workbook-agent-audit-tools.js'
import {
  handleWorkbookAgentConditionalFormatToolCall,
  workbookAgentConditionalFormatToolSpecs,
} from './workbook-agent-conditional-format-tools.js'
import { handleWorkbookAgentObjectToolCall, workbookAgentObjectToolSpecs } from './workbook-agent-object-tools.js'
import { handleWorkbookAgentMediaToolCall, workbookAgentMediaToolSpecs } from './workbook-agent-media-tools.js'
import { handleWorkbookAgentProtectionToolCall, workbookAgentProtectionToolSpecs } from './workbook-agent-protection-tools.js'
import { handleWorkbookAgentSheetReadToolCall, workbookAgentSheetReadToolSpecs } from './workbook-agent-sheet-read-tools.js'
import { handleWorkbookAgentValidationToolCall, workbookAgentValidationToolSpecs } from './workbook-agent-validation-tools.js'
import {
  cellRangeRefJsonSchema,
  cellRangeRefSchema,
  rangeOrSelectorJsonSchema,
  rangeOrSelectorSchema,
  readRangeToolArgsSchema,
  resolveFormulaRangeRequest,
  resolveRangeOrSelectorRequest,
  resolveReadRangeRequest,
  resolveTransferRangeRequest,
  resolveWriteRangeRequest,
  setFormulaToolArgsSchema,
  transferRangeToolArgsSchema,
  workbookSemanticSelectorJsonSchema,
  writeRangeToolArgsSchema,
} from './workbook-agent-selector-tooling.js'
import {
  normalizeWorkbookAgentStylePatch,
  workbookAgentStylePatchHasChanges,
  workbookAgentStylePatchJsonSchema,
  workbookAgentStylePatchSchema,
} from './workbook-agent-style-patches.js'
import {
  listWorkbookNamedRanges,
  listWorkbookTables,
  type ResolvedWorkbookSelector,
  workbookSemanticSelectorSchema,
} from './workbook-selector-resolver.js'
import {
  parseWorkbookAgentStructuralToolCommand,
  sortToolArgsSchema,
  workbookAgentStructuralToolSpecs,
} from './workbook-agent-structural-tools.js'

const MAX_MUTATION_RANGE_CELLS = 400
const MAX_READ_RANGE_CELLS = 4000

const inspectCellToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  address: z.string().min(1).optional(),
})
const formulaIssueToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  limit: z.number().int().positive().max(200).optional(),
})
const readRecentChangesToolArgsSchema = z.object({
  limit: z.number().int().positive().max(50).optional(),
})
const startWorkflowToolArgsSchema = z.discriminatedUnion('workflowTemplate', [
  z.object({
    workflowTemplate: z.literal('summarizeWorkbook'),
  }),
  z.object({
    workflowTemplate: z.literal('summarizeCurrentSheet'),
  }),
  z.object({
    workflowTemplate: z.literal('describeRecentChanges'),
  }),
  z.object({
    workflowTemplate: z.literal('findFormulaIssues'),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(200).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('highlightFormulaIssues'),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(200).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('repairFormulaIssues'),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(200).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('highlightCurrentSheetOutliers'),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(100).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('styleCurrentSheetHeaders'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('normalizeCurrentSheetHeaders'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('normalizeCurrentSheetNumberFormats'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('normalizeCurrentSheetWhitespace'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('fillCurrentSheetFormulasDown'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('traceSelectionDependencies'),
  }),
  z.object({
    workflowTemplate: z.literal('explainSelectionCell'),
  }),
  z.object({
    workflowTemplate: z.literal('searchWorkbookQuery'),
    query: z.string().trim().min(1),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(50).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('createCurrentSheetRollup'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('createCurrentSheetReviewTab'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('createSheet'),
    name: z.string().trim().min(1),
  }),
  z.object({
    workflowTemplate: z.literal('renameCurrentSheet'),
    name: z.string().trim().min(1),
  }),
  z.object({
    workflowTemplate: z.literal('hideCurrentRow'),
  }),
  z.object({
    workflowTemplate: z.literal('hideCurrentColumn'),
  }),
  z.object({
    workflowTemplate: z.literal('unhideCurrentRow'),
  }),
  z.object({
    workflowTemplate: z.literal('unhideCurrentColumn'),
  }),
])
export type WorkbookAgentStartWorkflowRequest = z.infer<typeof startWorkflowToolArgsSchema>
const searchWorkbookToolArgsSchema = z.object({
  query: z.string().trim().min(1),
  sheetName: z.string().min(1).optional(),
  limit: z.number().int().positive().max(50).optional(),
})
const traceDependenciesToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  address: z.string().min(1).optional(),
  direction: z.enum(['precedents', 'dependents', 'both']).optional(),
  depth: z.number().int().positive().max(4).optional(),
})
const setActiveSheetToolArgsSchema = z.object({
  sheetName: z.string().trim().min(1),
  address: z.string().trim().min(1).optional(),
})
const setSelectionToolArgsSchema = z.object({
  sheetName: z.string().trim().min(1).optional(),
  address: z.string().trim().min(1),
  endAddress: z.string().trim().min(1).optional(),
})
const readRenderedRangeToolArgsSchema = z.object({
  sheetName: z.string().trim().min(1),
  startAddress: z.string().trim().min(1),
  endAddress: z.string().trim().min(1),
})
const applyAndVerifyToolArgsSchema = z.object({
  range: cellRangeRefSchema.optional(),
  includeFormulaIssues: z.boolean().optional(),
  includeInvariants: z.boolean().optional(),
})

const clearRangeToolArgsSchema = rangeOrSelectorSchema
const formatRangeToolArgsSchema = z
  .object({
    range: cellRangeRefSchema.optional(),
    selector: workbookSemanticSelectorSchema.optional(),
    patch: workbookAgentStylePatchSchema.optional(),
    numberFormat: setRangeNumberFormatArgsSchema.shape.format.optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: 'Provide exactly one of range or selector',
  })
  .refine((value) => value.patch !== undefined || value.numberFormat !== undefined, {
    message: 'patch or numberFormat is required',
  })

type FormatRangeNumberFormatInput = NonNullable<z.infer<typeof setRangeNumberFormatArgsSchema.shape.format>>

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

function serializeSelectorResolution(resolution: ResolvedWorkbookSelector | null) {
  if (!resolution) {
    return null
  }
  return {
    objectType: resolution.objectType,
    displayLabel: resolution.displayLabel,
    resolvedRevision: resolution.resolvedRevision,
    derivedA1Ranges: resolution.derivedA1Ranges,
    table: resolution.table,
    namedRange: resolution.namedRange,
  }
}

function summarizeWorkbookChangeRecord(record: Awaited<ReturnType<ZeroSyncService['listWorkbookChanges']>>[number]) {
  return {
    revision: record.revision,
    actorUserId: record.actorUserId,
    eventKind: record.eventKind,
    summary: record.summary,
    sheetName: record.sheetName,
    anchorAddress: record.anchorAddress,
    range: record.range,
    createdAtUnixMs: record.createdAtUnixMs,
    revertedByRevision: record.revertedByRevision,
    revertsRevision: record.revertsRevision,
  }
}

function normalizeFormula(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula
}

function normalizeNumberFormatInput(input: FormatRangeNumberFormatInput): CellNumberFormatInput {
  if (typeof input === 'string') {
    return input
  }

  const preset: CellNumberFormatPreset = {
    kind: input.kind,
  }
  if (typeof input.currency === 'string') {
    preset.currency = input.currency
  }
  if (typeof input.decimals === 'number') {
    preset.decimals = input.decimals
  }
  if (typeof input.useGrouping === 'boolean') {
    preset.useGrouping = input.useGrouping
  }
  if (input.negativeStyle === 'minus' || input.negativeStyle === 'parentheses') {
    preset.negativeStyle = input.negativeStyle
  }
  if (input.zeroStyle === 'zero' || input.zeroStyle === 'dash') {
    preset.zeroStyle = input.zeroStyle
  }
  if (input.dateStyle === 'short' || input.dateStyle === 'iso') {
    preset.dateStyle = input.dateStyle
  }

  return normalizeCellNumberFormatPreset(preset)
}

function isNumericText(value: string): boolean {
  const trimmed = value.trim()
  if (trimmed.length === 0 || trimmed !== value) {
    return false
  }
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    return false
  }
  const unsigned = trimmed.replace(/^[+-]/, '')
  if (/^0\d/.test(unsigned)) {
    return false
  }
  const parsed = Number(trimmed)
  return Number.isFinite(parsed)
}

function coerceNumericText(value: string): number | string {
  return isNumericText(value) ? Number(value) : value
}

function excelDateSerialFromIsoDate(value: string): number {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value)
  if (!match) {
    throw new Error(`Date value must use YYYY-MM-DD format, received ${value}`)
  }
  const year = Number(match[1])
  const month = Number(match[2])
  const day = Number(match[3])
  const timestamp = Date.UTC(year, month - 1, day)
  const date = new Date(timestamp)
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error(`Invalid date value ${value}`)
  }
  const excelEpoch = Date.UTC(1899, 11, 30)
  return Math.trunc((timestamp - excelEpoch) / 86_400_000)
}

function normalizeBooleanInput(value: string | boolean): boolean {
  if (typeof value === 'boolean') {
    return value
  }
  const normalized = value.trim().toLowerCase()
  if (normalized === 'true') {
    return true
  }
  if (normalized === 'false') {
    return false
  }
  throw new Error(`Boolean value must be true or false, received ${value}`)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeWriteCellInput(cellInput: unknown): WorkbookAgentWriteCellInput {
  if (cellInput === null || typeof cellInput === 'number' || typeof cellInput === 'boolean') {
    return cellInput
  }
  if (typeof cellInput === 'string') {
    return cellInput.startsWith('=')
      ? {
          formula: `=${normalizeFormula(cellInput)}`,
        }
      : coerceNumericText(cellInput)
  }
  if (!isRecord(cellInput)) {
    throw new Error('Unsupported write_range cell input')
  }
  const record = cellInput
  if (record['type'] === 'blank') {
    return null
  }
  if (record['type'] === 'formula') {
    if (typeof record['formula'] !== 'string') {
      throw new Error('Typed formula cell requires a formula string')
    }
    return {
      formula: `=${normalizeFormula(record['formula'])}`,
    }
  }
  if (record['type'] === 'text') {
    if (typeof record['value'] !== 'string') {
      throw new Error('Typed text cell requires a string value')
    }
    return record['value']
  }
  if (record['type'] === 'number') {
    const value = record['value']
    if (typeof value === 'number') {
      return value
    }
    if (typeof value === 'string' && isNumericText(value)) {
      return Number(value)
    }
    throw new Error(`Typed number cell requires a finite number or numeric string, received ${String(value)}`)
  }
  if (record['type'] === 'date') {
    const value = record['value']
    if (typeof value === 'number' && Number.isFinite(value)) {
      return value
    }
    if (typeof value === 'string') {
      return excelDateSerialFromIsoDate(value)
    }
    throw new Error('Typed date cell requires an Excel serial number or YYYY-MM-DD string')
  }
  if (record['type'] === 'boolean') {
    const value = record['value']
    if (typeof value === 'string' || typeof value === 'boolean') {
      return normalizeBooleanInput(value)
    }
    throw new Error('Typed boolean cell requires a boolean or true/false string')
  }
  if (typeof record['formula'] === 'string') {
    return {
      formula: `=${normalizeFormula(record['formula'])}`,
    }
  }
  if ('value' in record) {
    const value = record['value']
    if (value === null || typeof value === 'number' || typeof value === 'boolean') {
      return value
    }
    if (typeof value === 'string') {
      return value.startsWith('=')
        ? {
            formula: `=${normalizeFormula(value)}`,
          }
        : coerceNumericText(value)
    }
  }
  throw new Error('Unsupported write_range cell input')
}

function normalizeRange(range: CellRangeRef): CellRangeRef & {
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

function countRangeCells(range: CellRangeRef): number {
  const bounds = normalizeRange(range)
  return (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1)
}

function countTotalRangeCells(ranges: readonly CellRangeRef[]): number {
  return ranges.reduce((sum, range) => sum + countRangeCells(range), 0)
}

function ensureRangeLimit(range: CellRangeRef, limit: number): void {
  const count = countRangeCells(range)
  if (count > limit) {
    throw new Error(
      `Range ${range.sheetName}!${range.startAddress}:${range.endAddress} has ${String(count)} cells; tool limit is ${String(limit)} cells per call`,
    )
  }
}

function resolveSelectionRange(context: WorkbookAgentUiContext | null): CellRangeRef {
  if (!context) {
    throw new Error('No browser workbook context is attached to this chat session')
  }
  return {
    sheetName: context.selection.sheetName,
    startAddress: context.selection.range?.startAddress ?? context.selection.address,
    endAddress: context.selection.range?.endAddress ?? context.selection.address,
  }
}

function resolveVisibleRange(context: WorkbookAgentUiContext | null): CellRangeRef {
  if (!context) {
    throw new Error('No browser workbook context is attached to this chat session')
  }
  return viewportToRange(context.selection.sheetName, context.viewport)
}

function rangesEqual(left: CellRangeRef, right: CellRangeRef): boolean {
  const normalizedLeft = normalizeRange(left)
  const normalizedRight = normalizeRange(right)
  return (
    normalizedLeft.sheetName === normalizedRight.sheetName &&
    normalizedLeft.startAddress === normalizedRight.startAddress &&
    normalizedLeft.endAddress === normalizedRight.endAddress
  )
}

function selectRenderedRange(
  context: WorkbookAgentUiContext | null,
  range: CellRangeRef,
): {
  readonly available: boolean
  readonly stale: boolean
  readonly capturedAtUnixMs: number | null
  readonly batchId: number | null
  readonly range: unknown
} {
  const rendered = context?.rendered
  if (!rendered) {
    return {
      available: false,
      stale: true,
      capturedAtUnixMs: null,
      batchId: null,
      range: null,
    }
  }
  const candidates = [rendered.selection, rendered.visibleRange].filter((entry): entry is NonNullable<typeof entry> => entry !== null)
  const match = candidates.find((entry) => rangesEqual(entry.range, range)) ?? null
  return {
    available: match !== null,
    stale: match === null,
    capturedAtUnixMs: rendered.capturedAtUnixMs,
    batchId: rendered.batchId,
    range: match,
  }
}

async function buildVerificationReport(input: {
  readonly context: WorkbookAgentToolContext
  readonly revision: number | null
  readonly ranges: readonly CellRangeRef[]
  readonly includeFormulaIssues?: boolean
  readonly includeInvariants?: boolean
}) {
  return await input.context.zeroSyncService.inspectWorkbook(input.context.documentId, async (runtime) => {
    const uiContext = normalizeWorkbookAgentUiContext(runtime, input.context.uiContext)
    const normalizedRanges = input.ranges.map((range) => ({
      sheetName: range.sheetName,
      startAddress: normalizeRange(range).startAddress,
      endAddress: normalizeRange(range).endAddress,
    }))
    const authoritativeReadback = normalizedRanges.map((range) => inspectWorkbookRange(runtime, range))
    const renderedReadback = normalizedRanges.map((range) => {
      const renderedRange = selectRenderedRange(uiContext, range)
      return {
        requestedRange: range,
        available: renderedRange.available,
        stale: renderedRange.stale,
        capturedAtUnixMs: renderedRange.capturedAtUnixMs,
        batchId: renderedRange.batchId,
        range: renderedRange.range,
      }
    })
    const formulaIssues =
      input.includeFormulaIssues === false
        ? null
        : findWorkbookFormulaIssues(runtime, {
            limit: 100,
          })
    const invariants = input.includeInvariants === false ? null : await verifyWorkbookInvariants(runtime, { roundTrip: true })
    return {
      appliedRevision: input.revision,
      recalculationStatus: {
        headRevision: runtime.headRevision,
        calculatedRevision: runtime.calculatedRevision,
        upToDate: runtime.calculatedRevision >= runtime.headRevision,
        lastMetrics: runtime.engine.getLastMetrics(),
      },
      authoritativeReadback,
      renderedReadback,
      formulaIssues,
      invariants,
    }
  })
}

function resolveInspectionTarget(
  context: WorkbookAgentUiContext | null,
  args: z.infer<typeof inspectCellToolArgsSchema>,
): {
  sheetName: string
  address: string
} {
  if (args.sheetName && args.address) {
    return {
      sheetName: args.sheetName,
      address: args.address,
    }
  }
  if (!context) {
    throw new Error('sheetName and address are required when no browser workbook context exists')
  }
  return context.selection
}

function viewportToRange(sheetName: string, viewport: WorkbookViewport): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(viewport.rowStart, viewport.colStart),
    endAddress: formatAddress(viewport.rowEnd, viewport.colEnd),
  }
}

function viewportAroundAddress(sheetName: string, address: string, base?: WorkbookViewport): WorkbookViewport {
  const parsed = parseCellAddress(address, sheetName)
  if (base) {
    const rowCount = Math.max(1, base.rowEnd - base.rowStart + 1)
    const colCount = Math.max(1, base.colEnd - base.colStart + 1)
    return {
      rowStart: Math.max(0, parsed.row),
      rowEnd: Math.max(0, parsed.row + rowCount - 1),
      colStart: Math.max(0, parsed.col),
      colEnd: Math.max(0, parsed.col + colCount - 1),
    }
  }
  return {
    rowStart: parsed.row,
    rowEnd: parsed.row + 20,
    colStart: parsed.col,
    colEnd: parsed.col + 10,
  }
}

export interface WorkbookAgentToolContext {
  readonly documentId: string
  readonly session: SessionIdentity
  readonly uiContext: WorkbookAgentUiContext | null
  readonly zeroSyncService: ZeroSyncService
  readonly stageCommand: (command: WorkbookAgentCommand) => Promise<WorkbookAgentCommandBundle | WorkbookAgentStageCommandResult>
  readonly updateUiContext?: (context: WorkbookAgentUiContext | null) => Promise<void>
  readonly startWorkflow?: (input: WorkbookAgentStartWorkflowRequest) => Promise<WorkbookAgentWorkflowRun>
}

export interface WorkbookAgentStageCommandResult {
  readonly bundle: WorkbookAgentCommandBundle
  readonly executionRecord: WorkbookAgentExecutionRecord | null
  readonly disposition?: 'queuedForTurnApply' | 'reviewQueued'
}

const rangeTargetJsonSchema = {
  oneOf: [cellRangeRefJsonSchema, workbookSemanticSelectorJsonSchema],
}

function createDynamicToolSpecs(): readonly CodexDynamicToolSpec[] {
  return [
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.getContext,
      description:
        'Read the current browser workbook context, including selection geometry, the visible viewport, freeze panes, and hidden or resized axes in view.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
      description: 'Read a workbook summary with sheet names, populated cell counts, and used ranges.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.setActiveSheet,
      description: 'Make a sheet the active browser sheet for the attached workbook view, optionally moving selection to one cell.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['sheetName'],
        properties: {
          sheetName: { type: 'string' },
          address: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.setSelection,
      description: 'Move the attached browser workbook selection to a cell or rectangular range.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['address'],
        properties: {
          sheetName: { type: 'string' },
          address: { type: 'string' },
          endAddress: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRenderedSelection,
      description: 'Read the latest browser-rendered snapshot for the selected cells and compare it with authoritative workbook state.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRenderedRange,
      description:
        'Read the latest browser-rendered snapshot for a range when it is cached by the attached view and compare it with authoritative workbook state.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['sheetName', 'startAddress', 'endAddress'],
        properties: {
          sheetName: { type: 'string' },
          startAddress: { type: 'string' },
          endAddress: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.applyAndVerify,
      description:
        'Verify the latest applied workbook state by recalculating, reading authoritative cells, reading rendered browser cells when cached, scanning formula issues, and checking invariants.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          range: cellRangeRefJsonSchema,
          includeFormulaIssues: { type: 'boolean' },
          includeInvariants: { type: 'boolean' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.listNamedRanges,
      description: 'List workbook named ranges and named references, including resolved cell ranges and structured references.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.listTables,
      description: 'List workbook tables with sheet location, range, header/totals settings, and column names.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    ...workbookAgentSheetReadToolSpecs,
    ...workbookAgentAnnotationToolSpecs,
    ...workbookAgentConditionalFormatToolSpecs,
    ...workbookAgentObjectToolSpecs,
    ...workbookAgentMediaToolSpecs,
    ...workbookAgentProtectionToolSpecs,
    ...workbookAgentValidationToolSpecs,
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRange,
      description:
        'Read a rectangular cell range or selector target, including inputs, formulas, style ids, number-format ids, referenced formatting records, and sheet-state metadata for that window.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          sheetName: { type: 'string' },
          startAddress: { type: 'string' },
          endAddress: { type: 'string' },
          selector: workbookSemanticSelectorJsonSchema,
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readSelection,
      description:
        'Read the currently selected range from the attached browser workbook context with values, formulas, formatting catalogs, and local sheet-state metadata.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange,
      description:
        'Read the currently visible viewport range from the attached browser workbook context with values, formulas, formatting catalogs, and local sheet-state metadata.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges,
      description: 'Read the most recent durable workbook changes, including revisions, summaries, and affected ranges.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          limit: { type: 'number' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
      description:
        'Run a built-in workbook workflow for durable summaries, analysis, cleanup, search, rollups, review tabs, and safe structural tasks.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['workflowTemplate'],
        properties: {
          workflowTemplate: {
            type: 'string',
            enum: [
              'summarizeWorkbook',
              'summarizeCurrentSheet',
              'describeRecentChanges',
              'findFormulaIssues',
              'highlightFormulaIssues',
              'repairFormulaIssues',
              'highlightCurrentSheetOutliers',
              'styleCurrentSheetHeaders',
              'normalizeCurrentSheetHeaders',
              'normalizeCurrentSheetNumberFormats',
              'normalizeCurrentSheetWhitespace',
              'fillCurrentSheetFormulasDown',
              'traceSelectionDependencies',
              'explainSelectionCell',
              'searchWorkbookQuery',
              'createCurrentSheetRollup',
              'createCurrentSheetReviewTab',
              'createSheet',
              'renameCurrentSheet',
              'hideCurrentRow',
              'hideCurrentColumn',
              'unhideCurrentRow',
              'unhideCurrentColumn',
            ],
          },
          query: { type: 'string' },
          sheetName: { type: 'string' },
          limit: { type: 'number' },
          name: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.inspectCell,
      description:
        'Explain one cell, including input, current value, formula, display format, style record, number-format record, version, cycle status, and direct precedents/dependents. Defaults to the current selection when no address is provided.',
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
      name: WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues,
      description: 'Scan the workbook for broken formulas, error cells, cycles, and formulas still running through the JS fallback path.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          sheetName: { type: 'string' },
          limit: { type: 'number' },
        },
      },
    },
    ...workbookAgentAuditToolSpecs,
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook,
      description: 'Search workbook sheet names, addresses, formulas, inputs, and displayed values through the warm local runtime.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['query'],
        properties: {
          query: { type: 'string' },
          sheetName: { type: 'string' },
          limit: { type: 'number' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.traceDependencies,
      description:
        'Trace workbook precedents and dependents from one cell for multiple hops. Defaults to the current selection when no address is provided.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          sheetName: { type: 'string' },
          address: { type: 'string' },
          direction: {
            type: 'string',
            enum: ['precedents', 'dependents', 'both'],
          },
          depth: { type: 'number' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.writeRange,
      description:
        'Write a rectangular matrix of spreadsheet inputs starting at a top-left address or selector target. Use primitives for literals, {formula} for formulas, and null to clear a cell.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['values'],
        properties: {
          sheetName: { type: 'string' },
          startAddress: { type: 'string' },
          selector: workbookSemanticSelectorJsonSchema,
          values: {
            type: 'array',
            items: {
              type: 'array',
              items: {
                oneOf: [
                  { type: 'string' },
                  { type: 'number' },
                  { type: 'boolean' },
                  { type: 'null' },
                  {
                    oneOf: [
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'value'],
                        properties: {
                          type: { type: 'string', const: 'text' },
                          value: { type: 'string' },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'value'],
                        properties: {
                          type: { type: 'string', const: 'number' },
                          value: { oneOf: [{ type: 'string' }, { type: 'number' }] },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'value'],
                        properties: {
                          type: { type: 'string', const: 'date' },
                          value: { oneOf: [{ type: 'string' }, { type: 'number' }] },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'value'],
                        properties: {
                          type: { type: 'string', const: 'boolean' },
                          value: { oneOf: [{ type: 'string' }, { type: 'boolean' }] },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type'],
                        properties: {
                          type: { type: 'string', const: 'blank' },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'formula'],
                        properties: {
                          type: { type: 'string', const: 'formula' },
                          formula: { type: 'string' },
                        },
                      },
                    ],
                  },
                  {
                    type: 'object',
                    additionalProperties: false,
                    required: ['value'],
                    properties: {
                      value: {
                        oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }],
                      },
                    },
                  },
                  {
                    type: 'object',
                    additionalProperties: false,
                    required: ['formula'],
                    properties: {
                      formula: { type: 'string' },
                    },
                  },
                ],
              },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.setFormula,
      description:
        'Write one or more formulas into a target range or selector-resolved anchor. Use this when the request is explicitly about formulas rather than generic cell input.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['formulas'],
        properties: {
          range: cellRangeRefJsonSchema,
          selector: workbookSemanticSelectorJsonSchema,
          formulas: {
            type: 'array',
            items: {
              type: 'array',
              items: { type: 'string' },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.clearRange,
      description: 'Clear a rectangular range of cells or a selector-resolved workbook region.',
      inputSchema: rangeOrSelectorJsonSchema,
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.formatRange,
      description:
        'Apply style and/or number-format changes to a range or selector target. Use patch for style properties and numberFormat for number formatting.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          range: cellRangeRefJsonSchema,
          selector: workbookSemanticSelectorJsonSchema,
          patch: workbookAgentStylePatchJsonSchema,
          numberFormat: {
            oneOf: [{ type: 'string' }, { type: 'object' }],
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.fillRange,
      description:
        'Fill a target range from a source range using spreadsheet fill semantics. Source and target may be concrete ranges or semantic selectors.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['source', 'target'],
        properties: {
          source: rangeTargetJsonSchema,
          target: rangeTargetJsonSchema,
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.copyRange,
      description: 'Copy a source range into a target range. Source and target may be concrete ranges or semantic selectors.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['source', 'target'],
        properties: {
          source: rangeTargetJsonSchema,
          target: rangeTargetJsonSchema,
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.moveRange,
      description: 'Move a source range into a target range. Source and target may be concrete ranges or semantic selectors.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['source', 'target'],
        properties: {
          source: rangeTargetJsonSchema,
          target: rangeTargetJsonSchema,
        },
      },
    },
    ...workbookAgentStructuralToolSpecs,
  ] satisfies readonly CodexDynamicToolSpec[]
}

export const workbookAgentDynamicToolSpecs = createDynamicToolSpecs()

async function stageCommandResult(context: WorkbookAgentToolContext, command: WorkbookAgentCommand): Promise<CodexDynamicToolCallResult> {
  const result = await context.stageCommand(command)
  const normalized: WorkbookAgentStageCommandResult =
    'bundle' in result ? result : { bundle: result, executionRecord: null, disposition: 'reviewQueued' }
  const bundle = normalized.bundle
  if (normalized.executionRecord) {
    const verificationRanges = bundle.affectedRanges
      .filter((range) => range.role === 'target')
      .slice(0, 3)
      .map((range) => ({
        sheetName: range.sheetName,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      }))
    const verification = await buildVerificationReport({
      context,
      revision: normalized.executionRecord.appliedRevision,
      ranges: verificationRanges,
    })
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
        verification,
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

function workflowToolResult(run: WorkbookAgentWorkflowRun): CodexDynamicToolCallResult {
  return textToolResult(
    stringifyJson({
      workflowRun: {
        runId: run.runId,
        workflowTemplate: run.workflowTemplate,
        title: run.title,
        summary: run.summary,
        status: run.status,
        completedAtUnixMs: run.completedAtUnixMs,
        errorMessage: run.errorMessage,
      },
      artifact: run.artifact,
    }),
  )
}

export async function handleWorkbookAgentToolCall(
  context: WorkbookAgentToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult> {
  try {
    const normalizedTool = normalizeWorkbookAgentToolName(request.tool)
    const sheetReadToolResult = await handleWorkbookAgentSheetReadToolCall(context, request)
    if (sheetReadToolResult) {
      return sheetReadToolResult
    }
    const auditToolResult = await handleWorkbookAgentAuditToolCall(context, request)
    if (auditToolResult) {
      return auditToolResult
    }
    const objectToolResult = await handleWorkbookAgentObjectToolCall(context, request)
    if (objectToolResult) {
      return objectToolResult
    }
    const mediaToolResult = await handleWorkbookAgentMediaToolCall(context, request)
    if (mediaToolResult) {
      return mediaToolResult
    }
    const protectionToolResult = await handleWorkbookAgentProtectionToolCall(context, request)
    if (protectionToolResult) {
      return protectionToolResult
    }
    const annotationToolResult = await handleWorkbookAgentAnnotationToolCall(context, request)
    if (annotationToolResult) {
      return annotationToolResult
    }
    const conditionalFormatToolResult = await handleWorkbookAgentConditionalFormatToolCall(context, request)
    if (conditionalFormatToolResult) {
      return conditionalFormatToolResult
    }
    const validationToolResult = await handleWorkbookAgentValidationToolCall(context, request)
    if (validationToolResult) {
      return validationToolResult
    }
    const structuralCommand = parseWorkbookAgentStructuralToolCommand(request)
    if (structuralCommand) {
      return await stageCommandResult(context, structuralCommand)
    }
    switch (normalizedTool) {
      case WORKBOOK_AGENT_TOOL_NAMES.getContext: {
        const text = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          inspectWorkbookContext(runtime, context.uiContext),
        )
        return textToolResult(text)
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readWorkbook: {
        const summary = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
          documentId: context.documentId,
          context: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          ...summarizeWorkbookStructure(runtime),
        }))
        return textToolResult(stringifyJson(summary))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setActiveSheet: {
        const args = setActiveSheetToolArgsSchema.parse(request.arguments)
        const nextContext = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const sheets = runtime.engine.exportSnapshot().sheets.map((sheet) => sheet.name)
          if (!sheets.includes(args.sheetName)) {
            throw new Error(`Sheet ${args.sheetName} does not exist`)
          }
          const currentContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const address = args.address ?? currentContext?.selection.address ?? 'A1'
          return normalizeWorkbookAgentUiContext(runtime, {
            selection: {
              sheetName: args.sheetName,
              address,
              range: {
                startAddress: address,
                endAddress: address,
              },
            },
            viewport: viewportAroundAddress(args.sheetName, address, currentContext?.viewport),
          })
        })
        if (!context.updateUiContext) {
          throw new Error('Active sheet control is not available in this workbook assistant session')
        }
        await context.updateUiContext(nextContext)
        return textToolResult(
          stringifyJson({
            updated: true,
            context: nextContext,
          }),
        )
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setSelection: {
        const args = setSelectionToolArgsSchema.parse(request.arguments)
        const nextContext = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const currentContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const sheetName = args.sheetName ?? currentContext?.selection.sheetName
          if (!sheetName) {
            throw new Error('sheetName is required when no browser workbook context exists')
          }
          const sheets = runtime.engine.exportSnapshot().sheets.map((sheet) => sheet.name)
          if (!sheets.includes(sheetName)) {
            throw new Error(`Sheet ${sheetName} does not exist`)
          }
          const start = parseCellAddress(args.address, sheetName)
          const end = parseCellAddress(args.endAddress ?? args.address, sheetName)
          const startAddress = formatAddress(Math.min(start.row, end.row), Math.min(start.col, end.col))
          const endAddress = formatAddress(Math.max(start.row, end.row), Math.max(start.col, end.col))
          return normalizeWorkbookAgentUiContext(runtime, {
            selection: {
              sheetName,
              address: args.address,
              range: {
                startAddress,
                endAddress,
              },
            },
            viewport: viewportAroundAddress(sheetName, args.address, currentContext?.viewport),
          })
        })
        if (!context.updateUiContext) {
          throw new Error('Selection control is not available in this workbook assistant session')
        }
        await context.updateUiContext(nextContext)
        return textToolResult(
          stringifyJson({
            updated: true,
            context: nextContext,
          }),
        )
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRenderedSelection: {
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const uiContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const range = resolveSelectionRange(uiContext)
          ensureRangeLimit(range, MAX_READ_RANGE_CELLS)
          return {
            authoritativeReadback: inspectWorkbookRange(runtime, range),
            renderedReadback: selectRenderedRange(uiContext, range),
          }
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRenderedRange: {
        const args = readRenderedRangeToolArgsSchema.parse(request.arguments)
        const range = normalizeRange({
          sheetName: args.sheetName,
          startAddress: args.startAddress,
          endAddress: args.endAddress,
        })
        ensureRangeLimit(range, MAX_READ_RANGE_CELLS)
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const normalizedContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          return {
            authoritativeReadback: inspectWorkbookRange(runtime, range),
            renderedReadback: selectRenderedRange(normalizedContext, range),
          }
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.applyAndVerify: {
        const args = applyAndVerifyToolArgsSchema.parse(request.arguments)
        const ranges = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          if (args.range) {
            return [normalizeRange(args.range)]
          }
          const uiContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          return uiContext ? [resolveSelectionRange(uiContext)] : []
        })
        const report = await buildVerificationReport({
          context,
          revision: null,
          ranges,
          ...(args.includeFormulaIssues !== undefined ? { includeFormulaIssues: args.includeFormulaIssues } : {}),
          ...(args.includeInvariants !== undefined ? { includeInvariants: args.includeInvariants } : {}),
        })
        return textToolResult(stringifyJson(report))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.listNamedRanges: {
        const namedRanges = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
          documentId: context.documentId,
          namedRangeCount: runtime.engine.getDefinedNames().length,
          namedRanges: listWorkbookNamedRanges(runtime),
        }))
        return textToolResult(stringifyJson(namedRanges))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.listTables: {
        const tables = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
          documentId: context.documentId,
          tableCount: runtime.engine.getTables().length,
          tables: listWorkbookTables(runtime),
        }))
        return textToolResult(stringifyJson(tables))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRange: {
        const args = readRangeToolArgsSchema.parse(request.arguments)
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const uiContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const resolved = resolveReadRangeRequest({
            runtime,
            args,
            uiContext,
          })
          const totalCells = countTotalRangeCells(resolved.ranges)
          if (totalCells > MAX_READ_RANGE_CELLS) {
            throw new Error(
              `Resolved selector spans ${String(totalCells)} cells; tool limit is ${String(MAX_READ_RANGE_CELLS)} cells per call`,
            )
          }
          const inspectedRanges = resolved.ranges.map((range) => inspectWorkbookRange(runtime, range))
          if (inspectedRanges.length === 1) {
            return {
              resolvedSelector: serializeSelectorResolution(resolved.resolution),
              ...inspectedRanges[0],
            }
          }
          return {
            resolvedSelector: serializeSelectorResolution(resolved.resolution),
            rangeCount: inspectedRanges.length,
            ranges: inspectedRanges,
          }
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readSelection: {
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const range = resolveSelectionRange(normalizeWorkbookAgentUiContext(runtime, context.uiContext))
          ensureRangeLimit(range, MAX_READ_RANGE_CELLS)
          return inspectWorkbookRange(runtime, range)
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange: {
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const range = resolveVisibleRange(normalizeWorkbookAgentUiContext(runtime, context.uiContext))
          ensureRangeLimit(range, MAX_READ_RANGE_CELLS)
          return inspectWorkbookRange(runtime, range)
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges: {
        const args = readRecentChangesToolArgsSchema.parse(request.arguments)
        const changes = await context.zeroSyncService.listWorkbookChanges(context.documentId, args.limit)
        return textToolResult(
          stringifyJson({
            documentId: context.documentId,
            changeCount: changes.length,
            changes: changes.map((record) => summarizeWorkbookChangeRecord(record)),
          }),
        )
      }
      case WORKBOOK_AGENT_TOOL_NAMES.startWorkflow: {
        const args = startWorkflowToolArgsSchema.parse(request.arguments)
        if (!context.startWorkflow) {
          throw new Error('Built-in workflow execution is not available in this session')
        }
        return workflowToolResult(await context.startWorkflow(args))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.inspectCell: {
        const args = inspectCellToolArgsSchema.parse(request.arguments)
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          inspectWorkbookCell(runtime, resolveInspectionTarget(normalizeWorkbookAgentUiContext(runtime, context.uiContext), args)),
        )
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues: {
        const args = formulaIssueToolArgsSchema.parse(request.arguments)
        const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          findWorkbookFormulaIssues(runtime, {
            ...(args.sheetName ? { sheetName: args.sheetName } : {}),
            ...(args.limit !== undefined ? { limit: args.limit } : {}),
          }),
        )
        return textToolResult(stringifyJson(report))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook: {
        const args = searchWorkbookToolArgsSchema.parse(request.arguments)
        const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          searchWorkbook(runtime, {
            query: args.query,
            ...(args.sheetName ? { sheetName: args.sheetName } : {}),
            ...(args.limit !== undefined ? { limit: args.limit } : {}),
          }),
        )
        return textToolResult(stringifyJson(report))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.traceDependencies: {
        const args = traceDependenciesToolArgsSchema.parse(request.arguments)
        const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const target = resolveInspectionTarget(normalizeWorkbookAgentUiContext(runtime, context.uiContext), args)
          return traceWorkbookDependencies(runtime, {
            sheetName: target.sheetName,
            address: target.address,
            ...(args.direction ? { direction: args.direction } : {}),
            ...(args.depth !== undefined ? { depth: args.depth } : {}),
          })
        })
        return textToolResult(stringifyJson(report))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.writeRange: {
        const args = writeRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveWriteRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        const values = args.values
        const start = parseCellAddress(resolved.startAddress, resolved.sheetName)
        const maxWidth = values.reduce((width, rowValues) => Math.max(width, rowValues.length), 0)
        const endAddress = formatAddress(start.row + values.length - 1, start.col + maxWidth - 1)
        ensureRangeLimit(
          {
            sheetName: resolved.sheetName,
            startAddress: resolved.startAddress,
            endAddress,
          },
          MAX_MUTATION_RANGE_CELLS,
        )
        return await stageCommandResult(context, {
          kind: 'writeRange',
          sheetName: resolved.sheetName,
          startAddress: resolved.startAddress,
          values: values.map((rowValues) => rowValues.map((cellInput) => normalizeWriteCellInput(cellInput))),
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setFormula: {
        const args = setFormulaToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveFormulaRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'setRangeFormulas',
          range: resolved.range,
          formulas: args.formulas,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.clearRange: {
        const args = clearRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'clearRange',
          range: resolved.range,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.formatRange: {
        const args = formatRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args: {
              ...(args.range ? { range: args.range } : {}),
              ...(args.selector ? { selector: args.selector } : {}),
            },
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        const formatCommand: Extract<WorkbookAgentCommand, { kind: 'formatRange' }> = {
          kind: 'formatRange',
          range: resolved.range,
        }
        if (args.patch !== undefined) {
          const normalizedPatch = normalizeWorkbookAgentStylePatch(args.patch)
          if (!workbookAgentStylePatchHasChanges(normalizedPatch)) {
            throw new Error('format_range patch did not include any supported style fields')
          }
          formatCommand.patch = normalizedPatch
        }
        if (args.numberFormat !== undefined) {
          formatCommand.numberFormat = normalizeNumberFormatInput(args.numberFormat)
        }
        return await stageCommandResult(context, formatCommand)
      }
      case WORKBOOK_AGENT_TOOL_NAMES.fillRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveTransferRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.source, MAX_MUTATION_RANGE_CELLS)
        ensureRangeLimit(resolved.target, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'fillRange',
          source: resolved.source,
          target: resolved.target,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.copyRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveTransferRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.source, MAX_MUTATION_RANGE_CELLS)
        ensureRangeLimit(resolved.target, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'copyRange',
          source: resolved.source,
          target: resolved.target,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.moveRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveTransferRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.source, MAX_MUTATION_RANGE_CELLS)
        ensureRangeLimit(resolved.target, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'moveRange',
          source: resolved.source,
          target: resolved.target,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setFilter: {
        const args = rangeOrSelectorSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'setFilter',
          range: resolved.range,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.clearFilter: {
        const args = rangeOrSelectorSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'clearFilter',
          range: resolved.range,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setSort: {
        const args = sortToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args: {
              ...(args.range ? { range: args.range } : {}),
              ...(args.selector ? { selector: args.selector } : {}),
            },
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'setSort',
          range: resolved.range,
          keys: args.keys.map((key) => ({
            keyAddress: key.keyAddress,
            direction: key.direction,
          })),
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.clearSort: {
        const args = rangeOrSelectorSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureRangeLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'clearSort',
          range: resolved.range,
        })
      }
      default:
        return textToolResult(`Unknown bilig tool: ${request.tool}`, false)
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    return textToolResult(`Tool ${request.tool} failed: ${message}`, false)
  }
}
