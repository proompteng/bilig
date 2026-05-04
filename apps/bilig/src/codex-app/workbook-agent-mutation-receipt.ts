import type {
  CodexDynamicToolCallResult,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewCellDiff,
  WorkbookAgentPreviewRange,
  WorkbookAgentWriteCellInput,
} from '@bilig/agent-api'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, LiteralInput } from '@bilig/protocol'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { verifyWorkbookInvariants } from './workbook-agent-audit.js'
import { findWorkbookFormulaIssues } from './workbook-agent-comprehension.js'
import { inspectWorkbookRange, normalizeWorkbookAgentUiContext } from './workbook-agent-inspection.js'
import { countWorkbookAgentRangeCells, createWorkbookAgentRangeChunkPlan, toWorkbookAgentRangeRef } from './workbook-agent-range-chunks.js'
import {
  emptyWorkbookRenderedReadbackProof,
  selectWorkbookRenderedReadback,
  type WorkbookRenderedReadbackProof,
  type WorkbookVerificationMismatch,
} from './workbook-agent-rendered-readback.js'
import { stringifyJson, textToolResult, type WorkbookAgentStageCommandResult } from './workbook-agent-tool-shared.js'

const MAX_VERIFICATION_RANGES = 3
const MAX_RECEIPT_READBACK_CELLS = 4000

export interface WorkbookAgentMutationReceiptRange {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
  readonly role: WorkbookAgentPreviewRange['role']
  readonly kind: 'values' | 'formulas' | 'formats' | 'tables' | 'objects' | 'selection' | 'sheet'
}

export interface WorkbookAuthoritativeReadbackProof {
  readonly requested: boolean
  readonly matched: boolean | null
  readonly ranges: readonly unknown[]
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
}

export interface WorkbookToolMutationReceipt {
  readonly toolName: string
  readonly status: 'applied' | 'staged' | 'queued' | 'failed' | 'verification_incomplete'
  readonly revision: {
    readonly before: number | null
    readonly after: number | null
  }
  readonly affectedRanges: readonly WorkbookAgentMutationReceiptRange[]
  readonly authoritativeReadback: WorkbookAuthoritativeReadbackProof
  readonly renderedReadback: WorkbookRenderedReadbackProof
  readonly undo: {
    readonly available: boolean
    readonly token: string | null
    readonly reasonUnavailable: string | null
  }
  readonly warnings: readonly string[]
}

export interface WorkbookAgentToolStageContext {
  readonly documentId: string
  readonly session: SessionIdentity
  readonly uiContext: WorkbookAgentUiContext | null
  readonly zeroSyncService: ZeroSyncService
  readonly stageCommand: (command: WorkbookAgentCommand) => Promise<WorkbookAgentCommandBundle | WorkbookAgentStageCommandResult>
}

function commandKind(command: WorkbookAgentCommand): WorkbookAgentMutationReceiptRange['kind'] {
  switch (command.kind) {
    case 'writeRange':
    case 'clearRange':
    case 'fillRange':
    case 'copyRange':
    case 'moveRange':
      return 'values'
    case 'setRangeFormulas':
      return 'formulas'
    case 'formatRange':
    case 'setDataValidation':
    case 'clearDataValidation':
    case 'upsertConditionalFormat':
    case 'deleteConditionalFormat':
    case 'setSheetProtection':
    case 'clearSheetProtection':
    case 'upsertRangeProtection':
    case 'deleteRangeProtection':
    case 'upsertCommentThread':
    case 'deleteCommentThread':
    case 'upsertNote':
    case 'deleteNote':
      return 'formats'
    case 'upsertTable':
    case 'deleteTable':
    case 'upsertPivotTable':
    case 'deletePivotTable':
      return 'tables'
    case 'upsertDefinedName':
    case 'deleteDefinedName':
    case 'upsertChart':
    case 'deleteChart':
    case 'upsertImage':
    case 'deleteImage':
    case 'upsertShape':
    case 'deleteShape':
      return 'objects'
    case 'createSheet':
    case 'renameSheet':
    case 'deleteSheet':
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns':
    case 'setFreezePane':
    case 'setFilter':
    case 'clearFilter':
    case 'setSort':
    case 'clearSort':
    case 'updateRowMetadata':
    case 'updateColumnMetadata':
      return 'sheet'
    default: {
      const exhaustive: never = command
      return exhaustive
    }
  }
}

function receiptRanges(bundle: WorkbookAgentCommandBundle): readonly WorkbookAgentMutationReceiptRange[] {
  return bundle.affectedRanges.map((range) => ({
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
    role: range.role,
    kind: commandKind(bundle.commands[0] ?? ({ kind: 'clearRange', range } as WorkbookAgentCommand)),
  }))
}

function verificationRanges(bundle: WorkbookAgentCommandBundle): readonly CellRangeRef[] {
  return bundle.affectedRanges
    .filter((range) => range.role === 'target')
    .slice(0, MAX_VERIFICATION_RANGES)
    .map((range) => toWorkbookAgentRangeRef(range))
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function valuesEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) {
    return true
  }
  return JSON.stringify(left) === JSON.stringify(right)
}

function isFormulaWriteCellInput(value: WorkbookAgentWriteCellInput): value is { readonly formula: string } {
  return typeof value === 'object' && value !== null && 'formula' in value && typeof value.formula === 'string'
}

function literalWriteCellInput(value: WorkbookAgentWriteCellInput): LiteralInput | null {
  if (isFormulaWriteCellInput(value)) {
    return null
  }
  if (typeof value === 'object' && value !== null && 'value' in value) {
    return value.value
  }
  return value
}

function deriveWriteRangePreviewDiffs(bundle: WorkbookAgentCommandBundle): readonly WorkbookAgentPreviewCellDiff[] {
  const diffs: WorkbookAgentPreviewCellDiff[] = []
  bundle.commands.forEach((command) => {
    if (command.kind === 'writeRange') {
      const start = parseCellAddress(command.startAddress, command.sheetName)
      command.values.forEach((rowValues, rowIndex) => {
        rowValues.forEach((cellInput, columnIndex) => {
          if (diffs.length >= MAX_RECEIPT_READBACK_CELLS) {
            return
          }
          const address = formatAddress(start.row + rowIndex, start.col + columnIndex)
          diffs.push({
            sheetName: command.sheetName,
            address,
            beforeInput: null,
            beforeFormula: null,
            afterInput: literalWriteCellInput(cellInput),
            afterFormula: isFormulaWriteCellInput(cellInput) ? cellInput.formula : null,
            changeKinds: isFormulaWriteCellInput(cellInput) ? ['formula'] : ['input'],
          })
        })
      })
      return
    }
    if (command.kind === 'setRangeFormulas') {
      const start = parseCellAddress(command.range.startAddress, command.range.sheetName)
      command.formulas.forEach((rowFormulas, rowIndex) => {
        rowFormulas.forEach((formula, columnIndex) => {
          if (diffs.length >= MAX_RECEIPT_READBACK_CELLS) {
            return
          }
          diffs.push({
            sheetName: command.range.sheetName,
            address: formatAddress(start.row + rowIndex, start.col + columnIndex),
            beforeInput: null,
            beforeFormula: null,
            afterInput: null,
            afterFormula: formula.startsWith('=') ? formula : `=${formula}`,
            changeKinds: ['formula'],
          })
        })
      })
    }
  })
  return diffs
}

function authoritativeRowsByAddress(readbacks: readonly unknown[]): Map<string, Record<string, unknown>> {
  const cells = new Map<string, Record<string, unknown>>()
  readbacks.forEach((readback) => {
    if (!isRecord(readback) || !Array.isArray(readback['rows']) || !isRecord(readback['range'])) {
      return
    }
    const sheetName = typeof readback['range']['sheetName'] === 'string' ? readback['range']['sheetName'] : ''
    readback['rows'].forEach((row) => {
      if (!Array.isArray(row)) {
        return
      }
      row.forEach((cell) => {
        if (isRecord(cell) && typeof cell['address'] === 'string') {
          cells.set(`${sheetName}!${cell['address']}`, cell)
        }
      })
    })
  })
  return cells
}

function collectComparablePreviewMismatches(input: {
  readonly previewDiffs: readonly WorkbookAgentPreviewCellDiff[]
  readonly readbacks: readonly unknown[]
}): {
  readonly matched: boolean | null
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
} {
  if (input.previewDiffs.length === 0) {
    return {
      matched: null,
      mismatches: [],
      incompleteReason: 'No value or formula preview expectations were available for authoritative comparison.',
    }
  }
  const cells = authoritativeRowsByAddress(input.readbacks)
  const mismatches: WorkbookVerificationMismatch[] = []
  let comparableCount = 0
  input.previewDiffs.forEach((diff) => {
    const cell = cells.get(`${diff.sheetName}!${diff.address}`)
    if (!cell) {
      mismatches.push({
        sheetName: diff.sheetName,
        address: diff.address,
        field: 'cell',
        expected: 'authoritative cell present',
        actual: null,
        source: 'authoritative',
      })
      return
    }
    if (diff.changeKinds.includes('input')) {
      comparableCount += 1
      const actualInput = cell['input'] ?? null
      const actualValue = cell['value'] ?? null
      if (!valuesEqual(diff.afterInput, actualInput) && !valuesEqual(diff.afterInput, actualValue)) {
        mismatches.push({
          sheetName: diff.sheetName,
          address: diff.address,
          field: 'input',
          expected: diff.afterInput,
          actual: {
            input: actualInput,
            value: actualValue,
          },
          source: 'authoritative',
        })
      }
    }
    if (diff.changeKinds.includes('formula')) {
      comparableCount += 1
      const actualFormula = cell['formula'] ?? null
      if (!valuesEqual(diff.afterFormula, actualFormula)) {
        mismatches.push({
          sheetName: diff.sheetName,
          address: diff.address,
          field: 'formula',
          expected: diff.afterFormula,
          actual: actualFormula,
          source: 'authoritative',
        })
      }
    }
  })
  if (comparableCount === 0) {
    return {
      matched: null,
      mismatches,
      incompleteReason: 'Preview contained only non-value changes, so authoritative value/formula matching was not applicable.',
    }
  }
  return {
    matched: mismatches.length === 0,
    mismatches,
    incompleteReason: mismatches.length === 0 ? null : 'Authoritative readback did not match preview expectations.',
  }
}

async function resolveUndoStatus(input: {
  readonly context: WorkbookAgentToolStageContext
  readonly appliedRevision: number | null
}): Promise<WorkbookToolMutationReceipt['undo']> {
  if (input.appliedRevision === null) {
    return {
      available: false,
      token: null,
      reasonUnavailable: 'Workbook mutation has not been applied yet.',
    }
  }
  const changes = await input.context.zeroSyncService.listWorkbookChanges(input.context.documentId, 25).catch(() => [])
  const matchingChange = changes.find((change) => change.revision === input.appliedRevision) ?? null
  if (matchingChange?.undoBundle) {
    return {
      available: true,
      token: `revision:${String(input.appliedRevision)}`,
      reasonUnavailable: null,
    }
  }
  return {
    available: false,
    token: null,
    reasonUnavailable: 'No persisted undo metadata was returned for the applied revision.',
  }
}

async function buildAuthoritativeReadback(input: {
  readonly context: WorkbookAgentToolStageContext
  readonly bundle: WorkbookAgentCommandBundle
  readonly executionRecord: WorkbookAgentExecutionRecord | null
  readonly ranges: readonly CellRangeRef[]
}): Promise<WorkbookAuthoritativeReadbackProof> {
  if (input.ranges.length === 0) {
    return {
      requested: false,
      matched: null,
      ranges: [],
      mismatches: [],
      incompleteReason: 'No target cell range was available for authoritative readback.',
    }
  }
  const readbacks = await input.context.zeroSyncService.inspectWorkbook(input.context.documentId, (runtime) =>
    input.ranges.map((range) => inspectWorkbookRange(runtime, range)),
  )
  const comparison = collectComparablePreviewMismatches({
    previewDiffs:
      input.executionRecord?.preview?.cellDiffs.length === 0
        ? deriveWriteRangePreviewDiffs(input.bundle)
        : (input.executionRecord?.preview?.cellDiffs ?? []),
    readbacks,
  })
  return {
    requested: true,
    matched: comparison.matched,
    ranges: readbacks,
    mismatches: comparison.mismatches,
    incompleteReason: comparison.incompleteReason,
  }
}

function firstAuthoritativeRows(readback: WorkbookAuthoritativeReadbackProof): readonly (readonly unknown[])[] | undefined {
  const first = readback.ranges[0]
  if (!isRecord(first) || !Array.isArray(first['rows'])) {
    return undefined
  }
  return first['rows'].filter((row): row is readonly unknown[] => Array.isArray(row))
}

function asReadonlyRows(rows: readonly unknown[]): readonly (readonly unknown[])[] {
  return rows.filter((row): row is unknown[] => Array.isArray(row))
}

async function buildRenderedReadback(input: {
  readonly context: WorkbookAgentToolStageContext
  readonly appliedRevision: number | null
  readonly ranges: readonly CellRangeRef[]
  readonly authoritativeReadback: WorkbookAuthoritativeReadbackProof
}): Promise<WorkbookRenderedReadbackProof> {
  const range = input.ranges[0] ?? null
  if (!range) {
    return emptyWorkbookRenderedReadbackProof({
      requested: false,
      reason: 'No target cell range was available for rendered readback.',
    })
  }
  const nextChunk =
    countWorkbookAgentRangeCells(range) > MAX_RECEIPT_READBACK_CELLS
      ? (createWorkbookAgentRangeChunkPlan(range, MAX_RECEIPT_READBACK_CELLS).chunks[1] ?? null)
      : null
  const uiContext = await input.context.zeroSyncService.inspectWorkbook(input.context.documentId, (runtime) =>
    normalizeWorkbookAgentUiContext(runtime, input.context.uiContext),
  )
  const authoritativeRows = firstAuthoritativeRows(input.authoritativeReadback)
  return selectWorkbookRenderedReadback({
    renderedContext: uiContext?.rendered,
    requestedRange: range,
    minBatchId: input.appliedRevision,
    nextChunk,
    ...(authoritativeRows !== undefined ? { authoritativeRows } : {}),
  })
}

export async function buildWorkbookAgentVerificationReport(input: {
  readonly context: WorkbookAgentToolStageContext
  readonly revision: number | null
  readonly ranges: readonly CellRangeRef[]
  readonly includeFormulaIssues?: boolean
  readonly includeInvariants?: boolean
}) {
  return await input.context.zeroSyncService.inspectWorkbook(input.context.documentId, async (runtime) => {
    const uiContext = normalizeWorkbookAgentUiContext(runtime, input.context.uiContext)
    const normalizedRanges = input.ranges.map((range) => toWorkbookAgentRangeRef(range))
    const authoritativeReadback = normalizedRanges.map((range) => inspectWorkbookRange(runtime, range))
    const renderedReadback = normalizedRanges.map((range) => {
      const authoritativeRange = inspectWorkbookRange(runtime, range)
      return selectWorkbookRenderedReadback({
        renderedContext: uiContext?.rendered,
        requestedRange: range,
        authoritativeRows: asReadonlyRows(authoritativeRange.rows),
        minBatchId: input.revision,
      })
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

async function buildMutationReceipt(input: {
  readonly context: WorkbookAgentToolStageContext
  readonly toolName: string
  readonly normalized: WorkbookAgentStageCommandResult
}): Promise<WorkbookToolMutationReceipt> {
  const { bundle, executionRecord } = input.normalized
  const ranges = executionRecord ? verificationRanges(bundle) : []
  const authoritativeReadback = executionRecord
    ? await buildAuthoritativeReadback({
        context: input.context,
        bundle,
        executionRecord,
        ranges,
      })
    : {
        requested: false,
        matched: null,
        ranges: [],
        mismatches: [],
        incompleteReason: 'Workbook mutation is not applied, so authoritative readback is not yet meaningful.',
      }
  const renderedReadback = executionRecord
    ? await buildRenderedReadback({
        context: input.context,
        appliedRevision: executionRecord.appliedRevision,
        ranges,
        authoritativeReadback,
      })
    : emptyWorkbookRenderedReadbackProof({
        requested: false,
        reason: 'Workbook mutation is not applied, so rendered readback is not yet meaningful.',
      })
  const undo = await resolveUndoStatus({
    context: input.context,
    appliedRevision: executionRecord?.appliedRevision ?? null,
  })
  const warnings: string[] = []
  if (!executionRecord && input.normalized.disposition === 'queuedForTurnApply') {
    warnings.push(
      'Queued workbook change sets are not completed mutations. The assistant must wait for apply and verify before claiming success.',
    )
  }
  if (!executionRecord && input.normalized.disposition === 'reviewQueued') {
    warnings.push('Workbook change set is waiting for owner review and has not modified the workbook yet.')
  }
  if (executionRecord && authoritativeReadback.matched !== true) {
    warnings.push(authoritativeReadback.incompleteReason ?? 'Authoritative readback did not prove the mutation.')
  }
  if (executionRecord && renderedReadback.matched !== true) {
    warnings.push(renderedReadback.incompleteReason ?? 'Rendered readback did not prove the mutation.')
  }
  if (executionRecord && !undo.available) {
    warnings.push(undo.reasonUnavailable ?? 'Undo status is unavailable.')
  }
  return {
    toolName: input.toolName,
    status: executionRecord
      ? renderedReadback.matched === true || !renderedReadback.requested
        ? 'applied'
        : 'verification_incomplete'
      : input.normalized.disposition === 'queuedForTurnApply'
        ? 'queued'
        : 'staged',
    revision: {
      before: bundle.baseRevision,
      after: executionRecord?.appliedRevision ?? null,
    },
    affectedRanges: receiptRanges(bundle),
    authoritativeReadback,
    renderedReadback,
    undo,
    warnings,
  }
}

export async function stageWorkbookAgentCommandResult(
  context: WorkbookAgentToolStageContext,
  command: WorkbookAgentCommand,
  toolName: string,
): Promise<CodexDynamicToolCallResult> {
  const result = await context.stageCommand(command)
  const normalized: WorkbookAgentStageCommandResult =
    'bundle' in result ? result : { bundle: result, executionRecord: null, disposition: 'reviewQueued' }
  const bundle = normalized.bundle
  const mutationReceipt = await buildMutationReceipt({
    context,
    toolName,
    normalized,
  })
  if (normalized.executionRecord) {
    const verification = await buildWorkbookAgentVerificationReport({
      context,
      revision: normalized.executionRecord.appliedRevision,
      ranges: verificationRanges(bundle),
    })
    return textToolResult(
      stringifyJson({
        applied: mutationReceipt.status === 'applied' || mutationReceipt.status === 'verification_incomplete',
        staged: false,
        reviewQueued: false,
        queuedForTurnApply: false,
        status: mutationReceipt.status,
        bundleId: bundle.id,
        summary: `Applied workbook change set at revision r${String(normalized.executionRecord.appliedRevision)}: ${normalized.executionRecord.summary}`,
        revision: normalized.executionRecord.appliedRevision,
        scope: normalized.executionRecord.scope,
        riskClass: normalized.executionRecord.riskClass,
        estimatedAffectedCells: bundle.estimatedAffectedCells,
        affectedRanges: bundle.affectedRanges,
        mutationReceipt,
        verification,
      }),
    )
  }
  if (normalized.disposition === 'queuedForTurnApply') {
    return textToolResult(
      stringifyJson({
        applied: false,
        staged: false,
        reviewQueued: false,
        queuedForTurnApply: true,
        status: 'queued',
        bundleId: bundle.id,
        summary: `Queued workbook change set for turn apply and verification is incomplete: ${bundle.summary}`,
        scope: bundle.scope,
        riskClass: bundle.riskClass,
        estimatedAffectedCells: bundle.estimatedAffectedCells,
        affectedRanges: bundle.affectedRanges,
        mutationReceipt,
      }),
    )
  }
  return textToolResult(
    stringifyJson({
      applied: false,
      staged: true,
      reviewQueued: true,
      queuedForTurnApply: false,
      status: 'staged',
      bundleId: bundle.id,
      summary: `Prepared workbook review item; the workbook is unchanged until this is applied: ${bundle.summary}`,
      scope: bundle.scope,
      riskClass: bundle.riskClass,
      estimatedAffectedCells: bundle.estimatedAffectedCells,
      affectedRanges: bundle.affectedRanges,
      mutationReceipt,
    }),
  )
}
