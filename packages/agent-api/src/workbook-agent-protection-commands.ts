import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, WorkbookRangeProtectionSnapshot, WorkbookSheetProtectionSnapshot } from '@bilig/protocol'
import type { WorkbookAgentCommand, WorkbookAgentPreviewRange } from './workbook-agent-bundles.js'

export type WorkbookAgentProtectionCommand = Extract<
  WorkbookAgentCommand,
  { kind: 'setSheetProtection' } | { kind: 'clearSheetProtection' } | { kind: 'upsertRangeProtection' } | { kind: 'deleteRangeProtection' }
>

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string'
  )
}

function isSheetProtection(value: unknown): value is WorkbookSheetProtectionSnapshot {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    (value['hideFormulas'] === undefined || typeof value['hideFormulas'] === 'boolean')
  )
}

function isRangeProtection(value: unknown): value is WorkbookRangeProtectionSnapshot {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    value['id'].trim().length > 0 &&
    isCellRangeRef(value['range']) &&
    (value['hideFormulas'] === undefined || typeof value['hideFormulas'] === 'boolean')
  )
}

function normalizeRangeBounds(range: CellRangeRef): CellRangeRef {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(Math.min(start.row, end.row), Math.min(start.col, end.col)),
    endAddress: formatAddress(Math.max(start.row, end.row), Math.max(start.col, end.col)),
  }
}

function countRangeCells(range: CellRangeRef): number {
  const normalized = normalizeRangeBounds(range)
  const start = parseCellAddress(normalized.startAddress, normalized.sheetName)
  const end = parseCellAddress(normalized.endAddress, normalized.sheetName)
  return (end.row - start.row + 1) * (end.col - start.col + 1)
}

function rangeLabel(range: CellRangeRef): string {
  const normalized = normalizeRangeBounds(range)
  return normalized.startAddress === normalized.endAddress
    ? `${normalized.sheetName}!${normalized.startAddress}`
    : `${normalized.sheetName}!${normalized.startAddress}:${normalized.endAddress}`
}

export function isWorkbookAgentProtectionCommandKind(kind: string): kind is WorkbookAgentProtectionCommand['kind'] {
  return (
    kind === 'setSheetProtection' || kind === 'clearSheetProtection' || kind === 'upsertRangeProtection' || kind === 'deleteRangeProtection'
  )
}

export function isWorkbookAgentProtectionCommand(command: WorkbookAgentCommand): command is WorkbookAgentProtectionCommand {
  return isWorkbookAgentProtectionCommandKind(command.kind)
}

export function isWorkbookAgentProtectionCommandValue(value: unknown): value is WorkbookAgentProtectionCommand {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'setSheetProtection':
      return isSheetProtection(value['protection'])
    case 'clearSheetProtection':
      return typeof value['sheetName'] === 'string'
    case 'upsertRangeProtection':
      return isRangeProtection(value['protection'])
    case 'deleteRangeProtection':
      return typeof value['id'] === 'string' && isCellRangeRef(value['range'])
    default:
      return false
  }
}

export function isHighRiskWorkbookAgentProtectionCommand(command: WorkbookAgentProtectionCommand): boolean {
  return command.kind === 'setSheetProtection' || command.kind === 'clearSheetProtection'
}

export function isWorkbookScopeProtectionCommand(command: WorkbookAgentProtectionCommand): boolean {
  return command.kind === 'setSheetProtection' || command.kind === 'clearSheetProtection'
}

export function describeWorkbookAgentProtectionCommand(command: WorkbookAgentProtectionCommand): string {
  switch (command.kind) {
    case 'setSheetProtection':
      return command.protection.hideFormulas
        ? `Protect sheet ${command.protection.sheetName} and hide formulas`
        : `Protect sheet ${command.protection.sheetName}`
    case 'clearSheetProtection':
      return `Unprotect sheet ${command.sheetName}`
    case 'upsertRangeProtection':
      return command.protection.hideFormulas
        ? `Protect ${rangeLabel(command.protection.range)} and hide formulas`
        : `Protect ${rangeLabel(command.protection.range)}`
    case 'deleteRangeProtection':
      return `Unprotect ${rangeLabel(command.range)}`
  }
}

export function estimateWorkbookAgentProtectionCommandAffectedCells(command: WorkbookAgentProtectionCommand): number {
  switch (command.kind) {
    case 'setSheetProtection':
    case 'clearSheetProtection':
      return 0
    case 'upsertRangeProtection':
      return countRangeCells(command.protection.range)
    case 'deleteRangeProtection':
      return countRangeCells(command.range)
  }
}

export function deriveWorkbookAgentProtectionCommandPreviewRanges(command: WorkbookAgentProtectionCommand): WorkbookAgentPreviewRange[] {
  switch (command.kind) {
    case 'setSheetProtection':
    case 'clearSheetProtection':
      return []
    case 'upsertRangeProtection':
      return [{ ...normalizeRangeBounds(command.protection.range), role: 'target' }]
    case 'deleteRangeProtection':
      return [{ ...normalizeRangeBounds(command.range), role: 'target' }]
  }
}

export function applyWorkbookAgentProtectionCommand(engine: SpreadsheetEngine, command: WorkbookAgentProtectionCommand): void {
  switch (command.kind) {
    case 'setSheetProtection':
      engine.setSheetProtection(structuredClone(command.protection))
      return
    case 'clearSheetProtection':
      engine.clearSheetProtection(command.sheetName)
      return
    case 'upsertRangeProtection':
      engine.setRangeProtection(structuredClone(command.protection))
      return
    case 'deleteRangeProtection':
      engine.deleteRangeProtection(command.id)
      return
  }
}
