import type {
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewCellDiff,
  WorkbookAgentWriteCellInput,
} from '@bilig/agent-api'
import { diffWorkbookSemanticSnapshots } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  buildCellNumberFormatCode,
  sanitizeCellStyleRecord,
  type CellRangeRef,
  type CellStylePatch,
  type LiteralInput,
  type WorkbookSnapshot,
} from '@bilig/protocol'
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

const MAX_RECEIPT_READBACK_CELLS = 4000

type WorkbookSemanticStyleRecord = NonNullable<NonNullable<WorkbookSnapshot['workbook']['metadata']>['styles']>[number]

export interface WorkbookAgentMutationProofContext {
  readonly documentId: string
  readonly session?: SessionIdentity
  readonly uiContext: WorkbookAgentUiContext | null
  readonly zeroSyncService: ZeroSyncService
  readonly stageCommand?: unknown
}

export interface WorkbookAuthoritativeReadbackProof {
  readonly requested: boolean
  readonly matched: boolean | null
  readonly ranges: readonly unknown[]
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
}

export interface WorkbookSemanticReadbackProof {
  readonly requested: boolean
  readonly matched: boolean | null
  readonly incompleteReason: string | null
}

export interface WorkbookMutationUndoProof {
  readonly available: boolean
  readonly token: string | null
  readonly reasonUnavailable: string | null
  readonly lookupFailed: boolean
}

function describeUnknownError(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message
  }
  if (typeof error === 'string' && error.trim().length > 0) {
    return error
  }
  return 'unknown error'
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

function patchValueMatches(actual: unknown, expected: unknown): boolean {
  if (expected === null) {
    return actual === null || actual === undefined
  }
  if (isRecord(expected)) {
    if (!isRecord(actual)) {
      return false
    }
    return Object.entries(expected).every(([key, value]) => patchValueMatches(actual[key], value))
  }
  return valuesEqual(actual, expected)
}

function styleMatchesPatch(style: Record<string, unknown>, patch: CellStylePatch): boolean {
  return Object.entries(patch).every(([key, value]) => patchValueMatches(style[key], value))
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

function authoritativeStylesById(readbacks: readonly unknown[]): Map<string, Record<string, unknown>> {
  const styles = new Map<string, Record<string, unknown>>()
  readbacks.forEach((readback) => {
    if (!isRecord(readback) || !Array.isArray(readback['styles'])) {
      return
    }
    readback['styles'].forEach((style) => {
      if (isRecord(style) && typeof style['id'] === 'string') {
        styles.set(style['id'], style)
      }
    })
  })
  return styles
}

function authoritativeNumberFormatsById(readbacks: readonly unknown[]): Map<string, Record<string, unknown>> {
  const numberFormats = new Map<string, Record<string, unknown>>()
  readbacks.forEach((readback) => {
    if (!isRecord(readback) || !Array.isArray(readback['numberFormats'])) {
      return
    }
    readback['numberFormats'].forEach((numberFormat) => {
      if (isRecord(numberFormat) && typeof numberFormat['id'] === 'string') {
        numberFormats.set(numberFormat['id'], numberFormat)
      }
    })
  })
  return numberFormats
}

function asLiteralInput(value: unknown): LiteralInput | undefined {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? value : undefined
}

function semanticStyleFromRecord(style: Record<string, unknown>, id: string): WorkbookSemanticStyleRecord {
  return sanitizeCellStyleRecord(id, style) ?? { id }
}

function readbackSemanticSnapshot(readbacks: readonly unknown[]): WorkbookSnapshot {
  const sheets = new Map<string, WorkbookSnapshot['sheets'][number]>()
  const styles = new Map<string, WorkbookSemanticStyleRecord>()
  readbacks.forEach((readback) => {
    if (!isRecord(readback) || !isRecord(readback['range']) || !Array.isArray(readback['rows'])) {
      return
    }
    const sheetName = typeof readback['range']['sheetName'] === 'string' ? readback['range']['sheetName'] : null
    if (!sheetName) {
      return
    }
    const sheet =
      sheets.get(sheetName) ??
      ({
        name: sheetName,
        order: sheets.size,
        metadata: {
          styleRanges: [],
        },
        cells: [],
      } satisfies WorkbookSnapshot['sheets'][number])
    readback['rows'].forEach((row) => {
      if (!Array.isArray(row)) {
        return
      }
      row.forEach((cell) => {
        if (!isRecord(cell) || typeof cell['address'] !== 'string') {
          return
        }
        const snapshotCell: WorkbookSnapshot['sheets'][number]['cells'][number] = {
          address: cell['address'],
        }
        const value = asLiteralInput(cell['value'] ?? null)
        if (value !== undefined) {
          snapshotCell.value = value
        }
        if (typeof cell['formula'] === 'string' && cell['formula'].length > 0) {
          snapshotCell.formula = cell['formula']
        }
        if (typeof cell['displayFormat'] === 'string' && cell['displayFormat'].length > 0) {
          snapshotCell.format = cell['displayFormat']
        }
        sheet.cells.push(snapshotCell)
        if (typeof cell['styleId'] === 'string') {
          sheet.metadata ??= {}
          sheet.metadata.styleRanges ??= []
          sheet.metadata.styleRanges.push({
            range: {
              sheetName,
              startAddress: cell['address'],
              endAddress: cell['address'],
            },
            styleId: cell['styleId'],
          })
        }
      })
    })
    if (Array.isArray(readback['styles'])) {
      readback['styles'].forEach((style) => {
        if (isRecord(style) && typeof style['id'] === 'string') {
          styles.set(style['id'], semanticStyleFromRecord(style, style['id']))
        }
      })
    }
    sheets.set(sheetName, sheet)
  })
  const semanticSheets = [...sheets.values()]
  semanticSheets.forEach((sheet, order) => {
    sheet.order = order
  })
  return {
    version: 1,
    workbook: {
      name: 'authoritative-readback',
      metadata: {
        styles: [...styles.values()],
      },
    },
    sheets: semanticSheets,
  }
}

function mergeStylePatch(style: WorkbookSemanticStyleRecord, patch: CellStylePatch): WorkbookSemanticStyleRecord {
  const merged = {
    ...style,
    ...(patch.fill ? { fill: { ...style.fill, ...patch.fill } } : {}),
    ...(patch.font ? { font: { ...style.font, ...patch.font } } : {}),
    ...(patch.alignment ? { alignment: { ...style.alignment, ...patch.alignment } } : {}),
    ...(patch.borders ? { borders: { ...style.borders, ...patch.borders } } : {}),
  }
  return sanitizeCellStyleRecord(style.id, merged) ?? { id: style.id }
}

function findSemanticSheet(snapshot: WorkbookSnapshot, sheetName: string): WorkbookSnapshot['sheets'][number] | null {
  return snapshot.sheets.find((sheet) => sheet.name === sheetName) ?? null
}

function findSemanticCell(
  snapshot: WorkbookSnapshot,
  sheetName: string,
  address: string,
): WorkbookSnapshot['sheets'][number]['cells'][number] | null {
  return findSemanticSheet(snapshot, sheetName)?.cells.find((cell) => cell.address === address) ?? null
}

function replaceCellStyleRange(input: {
  readonly snapshot: WorkbookSnapshot
  readonly sheetName: string
  readonly address: string
  readonly styleId: string
}): void {
  const sheet = findSemanticSheet(input.snapshot, input.sheetName)
  if (!sheet) {
    return
  }
  sheet.metadata ??= {}
  sheet.metadata.styleRanges = (sheet.metadata.styleRanges ?? []).filter(
    (styleRange) =>
      !(
        styleRange.range.sheetName === input.sheetName &&
        styleRange.range.startAddress === input.address &&
        styleRange.range.endAddress === input.address
      ),
  )
  sheet.metadata.styleRanges.push({
    range: {
      sheetName: input.sheetName,
      startAddress: input.address,
      endAddress: input.address,
    },
    styleId: input.styleId,
  })
}

function expectedSemanticSnapshot(input: {
  readonly actualSnapshot: WorkbookSnapshot
  readonly previewDiffs: readonly WorkbookAgentPreviewCellDiff[]
  readonly commands: readonly WorkbookAgentCommand[]
}): WorkbookSnapshot {
  const expected = structuredClone(input.actualSnapshot)
  input.previewDiffs.forEach((diff) => {
    const cell = findSemanticCell(expected, diff.sheetName, diff.address)
    if (!cell) {
      return
    }
    if (diff.changeKinds.includes('input')) {
      cell.value = diff.afterInput
      delete cell.formula
    }
    if (diff.changeKinds.includes('formula') && diff.afterFormula !== null) {
      cell.formula = diff.afterFormula
    }
  })
  const styles = new Map((expected.workbook.metadata?.styles ?? []).map((style) => [style.id, style]))
  let expectedStyleIndex = 0
  input.commands.forEach((command) => {
    if (command.kind !== 'formatRange') {
      return
    }
    const start = parseCellAddress(command.range.startAddress, command.range.sheetName)
    const end = parseCellAddress(command.range.endAddress, command.range.sheetName)
    const rowStart = Math.min(start.row, end.row)
    const rowEnd = Math.max(start.row, end.row)
    const colStart = Math.min(start.col, end.col)
    const colEnd = Math.max(start.col, end.col)
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        const address = formatAddress(row, col)
        const cell = findSemanticCell(expected, command.range.sheetName, address)
        if (!cell) {
          continue
        }
        if (command.numberFormat !== undefined) {
          cell.format = buildCellNumberFormatCode(command.numberFormat)
        }
        if (command.patch !== undefined) {
          const currentStyleRange = findSemanticSheet(expected, command.range.sheetName)?.metadata?.styleRanges?.find(
            (styleRange) =>
              styleRange.range.sheetName === command.range.sheetName &&
              styleRange.range.startAddress === address &&
              styleRange.range.endAddress === address,
          )
          const currentStyle = currentStyleRange ? (styles.get(currentStyleRange.styleId) ?? { id: currentStyleRange.styleId }) : { id: '' }
          const styleId = `expected-agent-style-${String(expectedStyleIndex)}`
          expectedStyleIndex += 1
          styles.set(styleId, mergeStylePatch({ ...currentStyle, id: styleId }, command.patch))
          replaceCellStyleRange({
            snapshot: expected,
            sheetName: command.range.sheetName,
            address,
            styleId,
          })
        }
      }
    }
  })
  expected.workbook.metadata ??= {}
  expected.workbook.metadata.styles = [...styles.values()]
  return expected
}

function collectSemanticReadbackMismatches(input: {
  readonly commands: readonly WorkbookAgentCommand[]
  readonly previewDiffs: readonly WorkbookAgentPreviewCellDiff[]
  readonly readbacks: readonly unknown[]
}): {
  readonly matched: boolean | null
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
} {
  if (input.previewDiffs.length === 0 && input.commands.every((command) => command.kind !== 'formatRange')) {
    return { matched: null, mismatches: [], incompleteReason: null }
  }
  const actualSnapshot = readbackSemanticSnapshot(input.readbacks)
  const expectedSnapshot = expectedSemanticSnapshot({
    actualSnapshot,
    previewDiffs: input.previewDiffs,
    commands: input.commands,
  })
  const diffs = diffWorkbookSemanticSnapshots(expectedSnapshot, actualSnapshot)
  return {
    matched: diffs.length === 0,
    mismatches: diffs.map((diff) => ({
      sheetName: input.previewDiffs[0]?.sheetName ?? input.commands[0]?.kind ?? 'workbook',
      address: input.previewDiffs[0]?.address ?? 'semantic',
      field: diff.path,
      expected: diff.left,
      actual: diff.right,
      source: 'authoritative',
    })),
    incompleteReason: diffs.length === 0 ? null : 'Authoritative semantic readback did not match expected workbook state.',
  }
}

function collectFormatCommandMismatches(input: {
  readonly commands: readonly WorkbookAgentCommand[]
  readonly readbacks: readonly unknown[]
}): {
  readonly matched: boolean | null
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
} {
  const formatCommands = input.commands.filter(
    (command): command is Extract<WorkbookAgentCommand, { kind: 'formatRange' }> => command.kind === 'formatRange',
  )
  if (formatCommands.length === 0) {
    return { matched: null, mismatches: [], incompleteReason: null }
  }
  const cells = authoritativeRowsByAddress(input.readbacks)
  const styles = authoritativeStylesById(input.readbacks)
  const numberFormats = authoritativeNumberFormatsById(input.readbacks)
  const mismatches: WorkbookVerificationMismatch[] = []
  let comparableCount = 0
  formatCommands.forEach((command) => {
    const start = parseCellAddress(command.range.startAddress, command.range.sheetName)
    const end = parseCellAddress(command.range.endAddress, command.range.sheetName)
    const rowStart = Math.min(start.row, end.row)
    const rowEnd = Math.max(start.row, end.row)
    const colStart = Math.min(start.col, end.col)
    const colEnd = Math.max(start.col, end.col)
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        const address = formatAddress(row, col)
        const cell = cells.get(`${command.range.sheetName}!${address}`)
        if (!cell) {
          mismatches.push({
            sheetName: command.range.sheetName,
            address,
            field: 'cell',
            expected: 'authoritative cell present',
            actual: null,
            source: 'authoritative',
          })
          continue
        }
        if (command.patch !== undefined) {
          comparableCount += 1
          const styleId = typeof cell['styleId'] === 'string' ? cell['styleId'] : null
          const style = styleId ? (styles.get(styleId) ?? null) : null
          if (!style || !styleMatchesPatch(style, command.patch)) {
            mismatches.push({
              sheetName: command.range.sheetName,
              address,
              field: 'style',
              expected: command.patch,
              actual: style,
              source: 'authoritative',
            })
          }
        }
        if (command.numberFormat !== undefined) {
          comparableCount += 1
          const numberFormatId = typeof cell['numberFormatId'] === 'string' ? cell['numberFormatId'] : null
          const numberFormat = numberFormatId ? (numberFormats.get(numberFormatId) ?? null) : null
          const expectedCode = buildCellNumberFormatCode(command.numberFormat)
          const actualCode = typeof numberFormat?.['code'] === 'string' ? numberFormat['code'] : null
          if (actualCode !== expectedCode) {
            mismatches.push({
              sheetName: command.range.sheetName,
              address,
              field: 'numberFormat',
              expected: expectedCode,
              actual: actualCode,
              source: 'authoritative',
            })
          }
        }
      }
    }
  })
  if (comparableCount === 0) {
    return {
      matched: null,
      mismatches,
      incompleteReason: 'No style or number-format expectations were available for authoritative comparison.',
    }
  }
  return {
    matched: mismatches.length === 0,
    mismatches,
    incompleteReason: mismatches.length === 0 ? null : 'Authoritative readback did not match formatting expectations.',
  }
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

export async function resolveWorkbookMutationUndoStatus(input: {
  readonly context: WorkbookAgentMutationProofContext
  readonly appliedRevision: number | null
}): Promise<WorkbookMutationUndoProof> {
  if (input.appliedRevision === null) {
    return {
      available: false,
      token: null,
      reasonUnavailable: 'Workbook mutation has not been applied yet.',
      lookupFailed: false,
    }
  }
  let changes: Awaited<ReturnType<ZeroSyncService['listWorkbookChanges']>>
  try {
    changes = await input.context.zeroSyncService.listWorkbookChanges(input.context.documentId, 25)
  } catch (error) {
    return {
      available: false,
      token: null,
      reasonUnavailable: `Undo metadata lookup failed for applied revision r${String(input.appliedRevision)}: ${describeUnknownError(error)}`,
      lookupFailed: true,
    }
  }
  const matchingChange = changes.find((change) => change.revision === input.appliedRevision) ?? null
  if (matchingChange?.undoBundle) {
    return {
      available: true,
      token: `revision:${String(input.appliedRevision)}`,
      reasonUnavailable: null,
      lookupFailed: false,
    }
  }
  return {
    available: false,
    token: null,
    reasonUnavailable: 'No persisted undo metadata was returned for the applied revision.',
    lookupFailed: false,
  }
}

export async function buildWorkbookAuthoritativeReadbackProof(input: {
  readonly context: WorkbookAgentMutationProofContext
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
  const previewDiffs =
    input.executionRecord?.preview?.cellDiffs.length === 0
      ? deriveWriteRangePreviewDiffs(input.bundle)
      : (input.executionRecord?.preview?.cellDiffs ?? [])
  const comparison = collectComparablePreviewMismatches({
    previewDiffs,
    readbacks,
  })
  const formatComparison = collectFormatCommandMismatches({
    commands: input.bundle.commands,
    readbacks,
  })
  const semanticComparison = collectSemanticReadbackMismatches({
    commands: input.bundle.commands,
    previewDiffs,
    readbacks,
  })
  const mismatches = [...comparison.mismatches, ...formatComparison.mismatches, ...semanticComparison.mismatches]
  const matched =
    comparison.matched === false || formatComparison.matched === false || semanticComparison.matched === false
      ? false
      : comparison.matched === true || formatComparison.matched === true || semanticComparison.matched === true
        ? true
        : null
  const incompleteReason =
    mismatches.length > 0
      ? comparison.matched === false
        ? comparison.incompleteReason
        : (formatComparison.incompleteReason ?? semanticComparison.incompleteReason ?? comparison.incompleteReason)
      : matched === true
        ? null
        : (formatComparison.incompleteReason ?? semanticComparison.incompleteReason ?? comparison.incompleteReason)
  return {
    requested: true,
    matched,
    ranges: readbacks,
    mismatches,
    incompleteReason,
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

export async function buildWorkbookRenderedReadbackProof(input: {
  readonly context: WorkbookAgentMutationProofContext
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
    minRevision: input.appliedRevision,
    nextChunk,
    ...(authoritativeRows !== undefined ? { authoritativeRows } : {}),
  })
}

export async function buildWorkbookAgentVerificationReport(input: {
  readonly context: WorkbookAgentMutationProofContext
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
        minRevision: input.revision,
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
