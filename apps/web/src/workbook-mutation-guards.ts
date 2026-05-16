import { isCommitOps } from '@bilig/core'
import type { CellNumberFormatInput, CellStyleField, CellStylePatch } from '@bilig/protocol'
import {
  CELL_BORDER_STYLE_VALUES,
  CELL_BORDER_WEIGHT_VALUES,
  CELL_DATE_STYLE_VALUES,
  CELL_HORIZONTAL_ALIGNMENT_VALUES,
  CELL_NUMBER_FORMAT_KIND_VALUES,
  CELL_NUMBER_NEGATIVE_STYLE_VALUES,
  CELL_NUMBER_ZERO_STYLE_VALUES,
  CELL_STYLE_FIELD_VALUES,
  CELL_VERTICAL_ALIGNMENT_VALUES,
  isCellRangeRef,
  isLiteralInput,
} from '@bilig/protocol'
import type { PendingWorkbookMutation, PendingWorkbookMutationInput, WorkbookMutationMethod } from './workbook-sync.js'

type KnownPendingWorkbookMutationMethod = PendingWorkbookMutationInput['method']
type CellStyleBorderSideName = 'top' | 'right' | 'bottom' | 'left'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

function isNonNegativeFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0
}

function isNullableNonNegativeFiniteNumber(value: unknown): value is number | null {
  return value === null || isNonNegativeFiniteNumber(value)
}

function isOptionalNullableString(value: unknown): value is string | null | undefined {
  return value === undefined || value === null || typeof value === 'string'
}

function isOptionalNullableFiniteNumber(value: unknown): value is number | null | undefined {
  return value === undefined || value === null || (typeof value === 'number' && Number.isFinite(value))
}

function isOptionalNullableBoolean(value: unknown): value is boolean | null | undefined {
  return value === undefined || value === null || typeof value === 'boolean'
}

export function isWorkbookSheetName(value: unknown): value is string {
  return isNonEmptyString(value)
}

export function isWorkbookStructuralIndex(value: unknown): value is number {
  return isNonNegativeInteger(value)
}

export function isWorkbookStructuralCount(value: unknown): value is number {
  return isNonNegativeInteger(value)
}

export function isWorkbookStructuralSize(value: unknown): value is number {
  return isNonNegativeFiniteNumber(value)
}

const CELL_STYLE_FIELD_VALUE_SET = new Set<string>(CELL_STYLE_FIELD_VALUES)

export function isCellStyleFieldList(value: unknown): value is CellStyleField[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === 'string' && CELL_STYLE_FIELD_VALUE_SET.has(entry))
}

const HORIZONTAL_ALIGNMENT_VALUES = new Set<string>(CELL_HORIZONTAL_ALIGNMENT_VALUES)
const VERTICAL_ALIGNMENT_VALUES = new Set<string>(CELL_VERTICAL_ALIGNMENT_VALUES)
const BORDER_STYLE_VALUES = new Set<string>(CELL_BORDER_STYLE_VALUES)
const BORDER_WEIGHT_VALUES = new Set<string>(CELL_BORDER_WEIGHT_VALUES)
const BORDER_SIDE_NAMES: readonly CellStyleBorderSideName[] = ['top', 'right', 'bottom', 'left']
const NUMBER_FORMAT_KIND_VALUES = new Set<string>(CELL_NUMBER_FORMAT_KIND_VALUES)
const NUMBER_FORMAT_NEGATIVE_STYLE_VALUES = new Set<string>(CELL_NUMBER_NEGATIVE_STYLE_VALUES)
const NUMBER_FORMAT_ZERO_STYLE_VALUES = new Set<string>(CELL_NUMBER_ZERO_STYLE_VALUES)
const NUMBER_FORMAT_DATE_STYLE_VALUES = new Set<string>(CELL_DATE_STYLE_VALUES)

function isOptionalNullableEnumValue(value: unknown, values: ReadonlySet<string>): value is string | null | undefined {
  return value === undefined || value === null || (typeof value === 'string' && values.has(value))
}

function isCellBorderSidePatchValue(value: unknown): boolean {
  return (
    value === null ||
    (isRecord(value) &&
      isOptionalNullableEnumValue(value['style'], BORDER_STYLE_VALUES) &&
      isOptionalNullableEnumValue(value['weight'], BORDER_WEIGHT_VALUES) &&
      isOptionalNullableString(value['color']))
  )
}

export function isCellStylePatchValue(value: unknown): value is CellStylePatch {
  if (!isRecord(value)) {
    return false
  }

  const fill = value['fill']
  if (fill !== undefined && fill !== null && (!isRecord(fill) || !isOptionalNullableString(fill['backgroundColor']))) {
    return false
  }

  const font = value['font']
  if (
    font !== undefined &&
    font !== null &&
    (!isRecord(font) ||
      !isOptionalNullableString(font['family']) ||
      !isOptionalNullableFiniteNumber(font['size']) ||
      !isOptionalNullableBoolean(font['bold']) ||
      !isOptionalNullableBoolean(font['italic']) ||
      !isOptionalNullableBoolean(font['underline']) ||
      !isOptionalNullableString(font['color']))
  ) {
    return false
  }

  const alignment = value['alignment']
  if (
    alignment !== undefined &&
    alignment !== null &&
    (!isRecord(alignment) ||
      !isOptionalNullableEnumValue(alignment['horizontal'], HORIZONTAL_ALIGNMENT_VALUES) ||
      !isOptionalNullableEnumValue(alignment['vertical'], VERTICAL_ALIGNMENT_VALUES) ||
      !isOptionalNullableBoolean(alignment['wrap']) ||
      !isOptionalNullableFiniteNumber(alignment['indent']) ||
      !isOptionalNullableBoolean(alignment['shrinkToFit']) ||
      !isOptionalNullableFiniteNumber(alignment['readingOrder']) ||
      !isOptionalNullableFiniteNumber(alignment['textRotation']) ||
      !isOptionalNullableBoolean(alignment['justifyLastLine']))
  ) {
    return false
  }

  const borders = value['borders']
  if (borders !== undefined && borders !== null) {
    if (!isRecord(borders)) {
      return false
    }
    for (const side of BORDER_SIDE_NAMES) {
      if (!isCellBorderSidePatchValue(borders[side])) {
        return false
      }
    }
  }

  return true
}

export function isCellNumberFormatInputValue(value: unknown): value is CellNumberFormatInput {
  return (
    typeof value === 'string' ||
    (isRecord(value) &&
      typeof value['kind'] === 'string' &&
      NUMBER_FORMAT_KIND_VALUES.has(value['kind']) &&
      (value['currency'] === undefined || typeof value['currency'] === 'string') &&
      (value['decimals'] === undefined ||
        (typeof value['decimals'] === 'number' && Number.isInteger(value['decimals']) && value['decimals'] >= 0)) &&
      (value['useGrouping'] === undefined || typeof value['useGrouping'] === 'boolean') &&
      (value['negativeStyle'] === undefined ||
        (typeof value['negativeStyle'] === 'string' && NUMBER_FORMAT_NEGATIVE_STYLE_VALUES.has(value['negativeStyle']))) &&
      (value['zeroStyle'] === undefined ||
        (typeof value['zeroStyle'] === 'string' && NUMBER_FORMAT_ZERO_STYLE_VALUES.has(value['zeroStyle']))) &&
      (value['dateStyle'] === undefined ||
        (typeof value['dateStyle'] === 'string' && NUMBER_FORMAT_DATE_STYLE_VALUES.has(value['dateStyle']))))
  )
}

export { isCellRangeRef, isCommitOps, isLiteralInput }

export function isWorkbookMutationMethod(value: unknown): value is WorkbookMutationMethod {
  return (
    value === 'setCellValue' ||
    value === 'setCellFormula' ||
    value === 'clearCell' ||
    value === 'clearRange' ||
    value === 'renderCommit' ||
    value === 'fillRange' ||
    value === 'copyRange' ||
    value === 'moveRange' ||
    value === 'insertRows' ||
    value === 'deleteRows' ||
    value === 'insertColumns' ||
    value === 'deleteColumns' ||
    value === 'updateRowMetadata' ||
    value === 'updateColumnMetadata' ||
    value === 'setFreezePane' ||
    value === 'mergeCells' ||
    value === 'unmergeCells' ||
    value === 'setRangeStyle' ||
    value === 'clearRangeStyle' ||
    value === 'setRangeNumberFormat' ||
    value === 'clearRangeNumberFormat'
  )
}

function isKnownPendingWorkbookMutationMethod(value: unknown): value is KnownPendingWorkbookMutationMethod {
  return value === 'updateColumnWidth' || isWorkbookMutationMethod(value)
}

function isString(value: unknown): value is string {
  return typeof value === 'string'
}

function isNullableWorkbookStructuralSize(value: unknown): value is number | null {
  return value === null || isWorkbookStructuralSize(value)
}

function isNullableBoolean(value: unknown): value is boolean | null {
  return value === null || typeof value === 'boolean'
}

function hasArgCount(args: readonly unknown[], count: number): boolean {
  return args.length === count
}

function hasOptionalArgCount(args: readonly unknown[], min: number, max: number): boolean {
  return args.length >= min && args.length <= max
}

function isPendingWorkbookMutationArgs(method: KnownPendingWorkbookMutationMethod, args: readonly unknown[]): boolean {
  switch (method) {
    case 'setCellValue': {
      const [sheetName, address, value] = args
      return hasArgCount(args, 3) && isWorkbookSheetName(sheetName) && isString(address) && isLiteralInput(value)
    }
    case 'setCellFormula': {
      const [sheetName, address, formula] = args
      return hasArgCount(args, 3) && isWorkbookSheetName(sheetName) && isString(address) && isString(formula)
    }
    case 'clearCell': {
      const [sheetName, address] = args
      return hasArgCount(args, 2) && isWorkbookSheetName(sheetName) && isString(address)
    }
    case 'clearRange':
    case 'mergeCells':
    case 'unmergeCells':
    case 'clearRangeNumberFormat': {
      const [range] = args
      return hasArgCount(args, 1) && isCellRangeRef(range)
    }
    case 'renderCommit': {
      const [ops] = args
      return hasArgCount(args, 1) && isCommitOps(ops)
    }
    case 'fillRange':
    case 'copyRange':
    case 'moveRange': {
      const [source, target] = args
      return hasArgCount(args, 2) && isCellRangeRef(source) && isCellRangeRef(target)
    }
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns': {
      const [sheetName, start, count] = args
      return hasArgCount(args, 3) && isWorkbookSheetName(sheetName) && isWorkbookStructuralIndex(start) && isWorkbookStructuralCount(count)
    }
    case 'updateRowMetadata':
    case 'updateColumnMetadata': {
      const [sheetName, start, count, size, hidden] = args
      return (
        hasArgCount(args, 5) &&
        isWorkbookSheetName(sheetName) &&
        isWorkbookStructuralIndex(start) &&
        isWorkbookStructuralCount(count) &&
        isNullableWorkbookStructuralSize(size) &&
        isNullableBoolean(hidden)
      )
    }
    case 'updateColumnWidth': {
      const [sheetName, columnIndex, width] = args
      return (
        hasArgCount(args, 3) && isWorkbookSheetName(sheetName) && isWorkbookStructuralIndex(columnIndex) && isWorkbookStructuralSize(width)
      )
    }
    case 'setFreezePane': {
      const [sheetName, rows, cols] = args
      return hasArgCount(args, 3) && isWorkbookSheetName(sheetName) && isWorkbookStructuralIndex(rows) && isWorkbookStructuralIndex(cols)
    }
    case 'setRangeStyle': {
      const [range, patch] = args
      return hasArgCount(args, 2) && isCellRangeRef(range) && isCellStylePatchValue(patch)
    }
    case 'clearRangeStyle': {
      const [range, fields] = args
      return hasOptionalArgCount(args, 1, 2) && isCellRangeRef(range) && (fields === undefined || isCellStyleFieldList(fields))
    }
    case 'setRangeNumberFormat': {
      const [range, format] = args
      return hasArgCount(args, 2) && isCellRangeRef(range) && isCellNumberFormatInputValue(format)
    }
  }
}

export function isPendingWorkbookMutationInput(value: unknown): value is PendingWorkbookMutationInput {
  return (
    isRecord(value) &&
    isKnownPendingWorkbookMutationMethod(value['method']) &&
    Array.isArray(value['args']) &&
    isPendingWorkbookMutationArgs(value['method'], value['args'])
  )
}

export function isPendingWorkbookMutation(value: unknown): value is PendingWorkbookMutation {
  return (
    isRecord(value) &&
    isNonEmptyString(value['id']) &&
    isNonNegativeInteger(value['localSeq']) &&
    isNonNegativeInteger(value['baseRevision']) &&
    isNonNegativeFiniteNumber(value['enqueuedAtUnixMs']) &&
    isNullableNonNegativeFiniteNumber(value['submittedAtUnixMs']) &&
    isNullableNonNegativeFiniteNumber(value['lastAttemptedAtUnixMs']) &&
    isNullableNonNegativeFiniteNumber(value['ackedAtUnixMs']) &&
    isNullableNonNegativeFiniteNumber(value['rebasedAtUnixMs']) &&
    isNullableNonNegativeFiniteNumber(value['failedAtUnixMs']) &&
    isNonNegativeInteger(value['attemptCount']) &&
    (value['failureMessage'] === null || typeof value['failureMessage'] === 'string') &&
    (value['status'] === 'local' ||
      value['status'] === 'submitted' ||
      value['status'] === 'acked' ||
      value['status'] === 'rebased' ||
      value['status'] === 'failed') &&
    isPendingWorkbookMutationInput(value)
  )
}

export function isPendingWorkbookMutationList(value: unknown): value is readonly PendingWorkbookMutation[] {
  return Array.isArray(value) && value.every((entry) => isPendingWorkbookMutation(entry))
}
