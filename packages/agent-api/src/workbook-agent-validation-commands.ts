import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type {
  CellRangeRef,
  LiteralInput,
  WorkbookDataValidationRuleSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookValidationListSourceSnapshot,
} from '@bilig/protocol'
import type { WorkbookAgentCommand, WorkbookAgentPreviewRange } from './workbook-agent-bundles.js'

export type WorkbookAgentValidationCommand = Extract<WorkbookAgentCommand, { kind: 'setDataValidation' } | { kind: 'clearDataValidation' }>

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isLiteralInputValue(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string'
  )
}

function isWorkbookValidationListSource(value: unknown): value is WorkbookValidationListSourceSnapshot {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'named-range':
      return typeof value['name'] === 'string' && value['name'].trim().length > 0
    case 'cell-ref':
      return typeof value['sheetName'] === 'string' && typeof value['address'] === 'string' && value['address'].trim().length > 0
    case 'range-ref':
      return isCellRangeRef(value)
    case 'structured-ref':
      return (
        typeof value['tableName'] === 'string' &&
        value['tableName'].trim().length > 0 &&
        typeof value['columnName'] === 'string' &&
        value['columnName'].trim().length > 0
      )
    default:
      return false
  }
}

function isWorkbookDataValidationRule(value: unknown): value is WorkbookDataValidationRuleSnapshot {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'list': {
      const hasValues =
        Array.isArray(value['values']) && value['values'].length > 0 && value['values'].every((entry) => isLiteralInputValue(entry))
      const hasSource = value['source'] !== undefined && isWorkbookValidationListSource(value['source'])
      return (hasValues ? 1 : 0) + (hasSource ? 1 : 0) === 1
    }
    case 'checkbox':
      return (
        (value['checkedValue'] === undefined || isLiteralInputValue(value['checkedValue'])) &&
        (value['uncheckedValue'] === undefined || isLiteralInputValue(value['uncheckedValue']))
      )
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength':
      return (
        typeof value['operator'] === 'string' &&
        Array.isArray(value['values']) &&
        value['values'].length > 0 &&
        value['values'].length <= 2 &&
        value['values'].every((entry) => isLiteralInputValue(entry))
      )
    default:
      return false
  }
}

function isWorkbookDataValidationSnapshot(value: unknown): value is WorkbookDataValidationSnapshot {
  return (
    isRecord(value) &&
    isCellRangeRef(value['range']) &&
    isWorkbookDataValidationRule(value['rule']) &&
    (value['allowBlank'] === undefined || typeof value['allowBlank'] === 'boolean') &&
    (value['showDropdown'] === undefined || typeof value['showDropdown'] === 'boolean') &&
    (value['promptTitle'] === undefined || typeof value['promptTitle'] === 'string') &&
    (value['promptMessage'] === undefined || typeof value['promptMessage'] === 'string') &&
    (value['errorStyle'] === undefined ||
      value['errorStyle'] === 'stop' ||
      value['errorStyle'] === 'warning' ||
      value['errorStyle'] === 'information') &&
    (value['errorTitle'] === undefined || typeof value['errorTitle'] === 'string') &&
    (value['errorMessage'] === undefined || typeof value['errorMessage'] === 'string')
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

function rangeLabel(range: Pick<CellRangeRef, 'sheetName' | 'startAddress' | 'endAddress'>): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

function countRangeCells(range: CellRangeRef): number {
  const normalized = normalizeRangeBounds(range)
  const start = parseCellAddress(normalized.startAddress, normalized.sheetName)
  const end = parseCellAddress(normalized.endAddress, normalized.sheetName)
  return (end.row - start.row + 1) * (end.col - start.col + 1)
}

export function isWorkbookAgentValidationCommandKind(kind: string): kind is WorkbookAgentValidationCommand['kind'] {
  return kind === 'setDataValidation' || kind === 'clearDataValidation'
}

export function isWorkbookAgentValidationCommand(command: WorkbookAgentCommand): command is WorkbookAgentValidationCommand {
  return isWorkbookAgentValidationCommandKind(command.kind)
}

export function isWorkbookAgentValidationCommandValue(value: unknown): value is WorkbookAgentValidationCommand {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'setDataValidation':
      return isWorkbookDataValidationSnapshot(value['validation'])
    case 'clearDataValidation':
      return isCellRangeRef(value['range'])
    default:
      return false
  }
}

export function isHighRiskWorkbookAgentValidationCommand(_command: WorkbookAgentValidationCommand): boolean {
  return false
}

export function isWorkbookScopeValidationCommand(_command: WorkbookAgentValidationCommand): boolean {
  return false
}

export function describeWorkbookAgentValidationCommand(command: WorkbookAgentValidationCommand): string {
  switch (command.kind) {
    case 'setDataValidation':
      return `Set data validation on ${rangeLabel(normalizeRangeBounds(command.validation.range))}`
    case 'clearDataValidation':
      return `Clear data validation on ${rangeLabel(normalizeRangeBounds(command.range))}`
  }
}

export function estimateWorkbookAgentValidationCommandAffectedCells(command: WorkbookAgentValidationCommand): number {
  return countRangeCells(command.kind === 'setDataValidation' ? command.validation.range : command.range)
}

export function deriveWorkbookAgentValidationCommandPreviewRanges(command: WorkbookAgentValidationCommand): WorkbookAgentPreviewRange[] {
  return [
    {
      ...normalizeRangeBounds(command.kind === 'setDataValidation' ? command.validation.range : command.range),
      role: 'target',
    },
  ]
}

export function applyWorkbookAgentValidationCommand(engine: SpreadsheetEngine, command: WorkbookAgentValidationCommand): void {
  switch (command.kind) {
    case 'setDataValidation':
      engine.setDataValidation(structuredClone(command.validation))
      return
    case 'clearDataValidation':
      engine.clearDataValidation(command.range.sheetName, command.range)
      return
  }
}
