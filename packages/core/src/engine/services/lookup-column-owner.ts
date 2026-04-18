import { ValueTag, type CellValue } from '@bilig/protocol'
import type { RuntimeColumnOwner } from './runtime-column-store-service.js'

const MAX_COLUMN_OWNER_SPAN = 65_536

const EMPTY_KIND = 0
const NUMERIC_KIND = 1
const BOOLEAN_KIND = 2
const TEXT_KIND = 3
const INVALID_KIND = 4

type ComparableKindCode = typeof EMPTY_KIND | typeof NUMERIC_KIND | typeof BOOLEAN_KIND | typeof TEXT_KIND | typeof INVALID_KIND

function decodeComparableKindCode(raw: number | undefined): ComparableKindCode {
  switch (raw) {
    case undefined:
      return EMPTY_KIND
    case EMPTY_KIND:
      return EMPTY_KIND
    case NUMERIC_KIND:
      return NUMERIC_KIND
    case BOOLEAN_KIND:
      return BOOLEAN_KIND
    case TEXT_KIND:
      return TEXT_KIND
    case INVALID_KIND:
      return INVALID_KIND
    default:
      return EMPTY_KIND
  }
}

export interface LookupColumnOwner {
  readonly sheetName: string
  readonly col: number
  columnVersion: number
  structureVersion: number
  sheetColumnVersions: Uint32Array
  readonly rowStart: number
  readonly rowEnd: number
  readonly length: number
  readonly kindCodes: Uint8Array
  readonly numericValues: Float64Array
  readonly textValues: string[]
  readonly rowLists: Map<string, number[]>
  sortedNumericAscendingBreaks: Uint32Array | undefined
  sortedNumericDescendingBreaks: Uint32Array | undefined
  sortedTextAscendingBreaks: Uint32Array | undefined
  sortedTextDescendingBreaks: Uint32Array | undefined
  incompatibleNumericPrefix: Uint32Array | undefined
  incompatibleTextPrefix: Uint32Array | undefined
  summariesDirty: boolean
}

export interface ExactRangeSummary {
  readonly comparableKind: 'numeric' | 'text' | 'mixed'
  readonly uniformStart: number | undefined
  readonly uniformStep: number | undefined
}

export interface ApproximateRangeSummary {
  readonly comparableKind: 'numeric' | 'text' | undefined
  readonly uniformStart: number | undefined
  readonly uniformStep: number | undefined
  readonly sortedAscending: boolean
  readonly sortedDescending: boolean
}

export function isLookupColumnOwner(value: unknown): value is LookupColumnOwner {
  return (
    value !== null &&
    typeof value === 'object' &&
    'sheetName' in value &&
    typeof value.sheetName === 'string' &&
    'col' in value &&
    typeof value.col === 'number' &&
    'rowStart' in value &&
    typeof value.rowStart === 'number' &&
    'rowEnd' in value &&
    typeof value.rowEnd === 'number' &&
    'length' in value &&
    typeof value.length === 'number' &&
    'kindCodes' in value &&
    value.kindCodes instanceof Uint8Array &&
    'numericValues' in value &&
    value.numericValues instanceof Float64Array &&
    'textValues' in value &&
    Array.isArray(value.textValues) &&
    'rowLists' in value &&
    value.rowLists instanceof Map
  )
}

export interface LookupColumnOwnerWrite {
  readonly row: number
  readonly oldValue: CellValue
  readonly newValue: CellValue
  readonly oldStringId?: number
  readonly newStringId?: number
}

function decodeValueTag(rawTag: number | undefined): ValueTag {
  if (rawTag === undefined) {
    return ValueTag.Empty
  }
  switch (rawTag) {
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    case 0:
    default:
      return ValueTag.Empty
  }
}

function kindCodeForValue(value: CellValue): ComparableKindCode {
  switch (value.tag) {
    case ValueTag.Empty:
      return EMPTY_KIND
    case ValueTag.Number:
      return NUMERIC_KIND
    case ValueTag.Boolean:
      return BOOLEAN_KIND
    case ValueTag.String:
      return TEXT_KIND
    case ValueTag.Error:
      return INVALID_KIND
  }
}

function numericValueForValue(value: CellValue): number {
  switch (value.tag) {
    case ValueTag.Empty:
      return 0
    case ValueTag.Number:
      return Object.is(value.value, -0) ? 0 : value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.String:
    case ValueTag.Error:
      return 0
  }
}

function textValueForValue(value: CellValue, normalizeStringId: (stringId: number) => string, stringId = 0): string {
  if (value.tag === ValueTag.String) {
    return (stringId !== 0 ? normalizeStringId(stringId) : value.value).toUpperCase()
  }
  return ''
}

function exactLookupKeyForValue(value: CellValue, normalizeStringId: (stringId: number) => string, stringId = 0): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'e:'
    case ValueTag.Number:
      return `n:${Object.is(value.value, -0) ? 0 : value.value}`
    case ValueTag.Boolean:
      return value.value ? 'b:1' : 'b:0'
    case ValueTag.String:
      return `s:${(stringId !== 0 ? normalizeStringId(stringId) : value.value).toUpperCase()}`
    case ValueTag.Error:
      return undefined
  }
}

function exactLookupKeyAt(owner: LookupColumnOwner, offset: number): string {
  switch (decodeComparableKindCode(owner.kindCodes[offset])) {
    case EMPTY_KIND:
      return 'e:'
    case NUMERIC_KIND:
      return `n:${owner.numericValues[offset] ?? 0}`
    case BOOLEAN_KIND:
      return (owner.numericValues[offset] ?? 0) !== 0 ? 'b:1' : 'b:0'
    case TEXT_KIND:
      return `s:${owner.textValues[offset] ?? ''}`
    case INVALID_KIND:
    default:
      return 'x:'
  }
}

function insertRowSorted(rows: number[], row: number): void {
  let low = 0
  let high = rows.length
  while (low < high) {
    const mid = (low + high) >> 1
    if (rows[mid]! < row) {
      low = mid + 1
    } else {
      high = mid
    }
  }
  rows.splice(low, 0, row)
}

function lowerBound(rows: readonly number[], target: number): number {
  let low = 0
  let high = rows.length
  while (low < high) {
    const mid = (low + high) >> 1
    if (rows[mid]! < target) {
      low = mid + 1
    } else {
      high = mid
    }
  }
  return low
}

function upperBound(rows: readonly number[], target: number): number {
  let low = 0
  let high = rows.length
  while (low < high) {
    const mid = (low + high) >> 1
    if (rows[mid]! <= target) {
      low = mid + 1
    } else {
      high = mid
    }
  }
  return low
}

function detectUniformNumericStepInOwner(
  owner: LookupColumnOwner,
  start: number,
  end: number,
): { start: number; step: number } | undefined {
  if (end - start < 1) {
    return undefined
  }
  const first = owner.numericValues[start]!
  const step = owner.numericValues[start + 1]! - first
  if (!Number.isFinite(step) || step === 0) {
    return undefined
  }
  for (let offset = start + 2; offset <= end; offset += 1) {
    if (owner.numericValues[offset]! - owner.numericValues[offset - 1]! !== step) {
      return undefined
    }
  }
  return { start: first, step }
}

export function findExactMatchInRange(
  owner: LookupColumnOwner,
  key: string,
  rowStart: number,
  rowEnd: number,
  searchMode: 1 | -1,
): number | undefined {
  const rows = owner.rowLists.get(key)
  if (!rows || rows.length === 0) {
    return undefined
  }
  if (searchMode === 1) {
    const index = lowerBound(rows, rowStart)
    const row = rows[index]
    return row !== undefined && row <= rowEnd ? row : undefined
  }
  const index = upperBound(rows, rowEnd) - 1
  const row = rows[index]
  return row !== undefined && row >= rowStart ? row : undefined
}

export function buildLookupColumnOwner(args: {
  readonly owner: RuntimeColumnOwner
  readonly normalizeStringId: (stringId: number) => string
}): LookupColumnOwner | undefined {
  let minRow = Number.POSITIVE_INFINITY
  let maxRow = Number.NEGATIVE_INFINITY

  args.owner.pages.forEach((page) => {
    for (let localRow = 0; localRow < page.tags.length; localRow += 1) {
      if (decodeValueTag(page.tags[localRow]) === ValueTag.Empty) {
        continue
      }
      const row = page.rowStart + localRow
      minRow = Math.min(minRow, row)
      maxRow = Math.max(maxRow, row)
    }
  })

  if (!Number.isFinite(minRow) || !Number.isFinite(maxRow)) {
    return undefined
  }

  const length = maxRow - minRow + 1
  if (length > MAX_COLUMN_OWNER_SPAN) {
    return undefined
  }

  const kindCodes = new Uint8Array(length)
  const numericValues = new Float64Array(length)
  const textValues = Array.from({ length }, () => '')

  args.owner.pages.forEach((page) => {
    for (let localRow = 0; localRow < page.tags.length; localRow += 1) {
      const tag = decodeValueTag(page.tags[localRow])
      if (tag === ValueTag.Empty) {
        continue
      }
      const absoluteRow = page.rowStart + localRow
      const offset = absoluteRow - minRow
      switch (tag) {
        case ValueTag.Number:
          kindCodes[offset] = NUMERIC_KIND
          numericValues[offset] = Object.is(page.numbers[localRow] ?? 0, -0) ? 0 : (page.numbers[localRow] ?? 0)
          break
        case ValueTag.Boolean:
          kindCodes[offset] = BOOLEAN_KIND
          numericValues[offset] = (page.numbers[localRow] ?? 0) !== 0 ? 1 : 0
          break
        case ValueTag.String:
          kindCodes[offset] = TEXT_KIND
          textValues[offset] = args.normalizeStringId(page.stringIds[localRow] ?? 0)
          break
        case ValueTag.Error:
          kindCodes[offset] = INVALID_KIND
          break
        default:
          break
      }
    }
  })

  const rowLists = new Map<string, number[]>()
  for (let offset = 0; offset < length; offset += 1) {
    const row = minRow + offset
    const key = exactLookupKeyAt(
      {
        sheetName: args.owner.sheetName,
        col: args.owner.col,
        columnVersion: args.owner.columnVersion,
        structureVersion: args.owner.structureVersion,
        sheetColumnVersions: args.owner.sheetColumnVersions,
        rowStart: minRow,
        rowEnd: maxRow,
        length,
        kindCodes,
        numericValues,
        textValues,
        rowLists: new Map(),
        sortedNumericAscendingBreaks: undefined,
        sortedNumericDescendingBreaks: undefined,
        sortedTextAscendingBreaks: undefined,
        sortedTextDescendingBreaks: undefined,
        incompatibleNumericPrefix: undefined,
        incompatibleTextPrefix: undefined,
        summariesDirty: true,
      },
      offset,
    )
    const rows = rowLists.get(key)
    if (rows) {
      rows.push(row)
    } else {
      rowLists.set(key, [row])
    }
  }

  return {
    sheetName: args.owner.sheetName,
    col: args.owner.col,
    columnVersion: args.owner.columnVersion,
    structureVersion: args.owner.structureVersion,
    sheetColumnVersions: args.owner.sheetColumnVersions,
    rowStart: minRow,
    rowEnd: maxRow,
    length,
    kindCodes,
    numericValues,
    textValues,
    rowLists,
    sortedNumericAscendingBreaks: undefined,
    sortedNumericDescendingBreaks: undefined,
    sortedTextAscendingBreaks: undefined,
    sortedTextDescendingBreaks: undefined,
    incompatibleNumericPrefix: undefined,
    incompatibleTextPrefix: undefined,
    summariesDirty: true,
  }
}

export function applyLookupColumnOwnerLiteralWrite(args: {
  readonly owner: LookupColumnOwner
  readonly write: LookupColumnOwnerWrite
  readonly normalizeStringId: (stringId: number) => string
}): boolean {
  if (args.write.row < args.owner.rowStart || args.write.row > args.owner.rowEnd) {
    return false
  }
  const offset = args.write.row - args.owner.rowStart
  const oldKey = exactLookupKeyForValue(args.write.oldValue, args.normalizeStringId, args.write.oldStringId)
  const newKey = exactLookupKeyForValue(args.write.newValue, args.normalizeStringId, args.write.newStringId)

  if (oldKey !== undefined) {
    const rows = args.owner.rowLists.get(oldKey)
    if (!rows) {
      return false
    }
    const rowIndex = rows.indexOf(args.write.row)
    if (rowIndex === -1) {
      return false
    }
    rows.splice(rowIndex, 1)
    if (rows.length === 0) {
      args.owner.rowLists.delete(oldKey)
    }
  }

  if (newKey !== undefined) {
    const rows = args.owner.rowLists.get(newKey)
    if (rows) {
      insertRowSorted(rows, args.write.row)
    } else {
      args.owner.rowLists.set(newKey, [args.write.row])
    }
  }

  args.owner.kindCodes[offset] = kindCodeForValue(args.write.newValue)
  args.owner.numericValues[offset] = numericValueForValue(args.write.newValue)
  args.owner.textValues[offset] = textValueForValue(args.write.newValue, args.normalizeStringId, args.write.newStringId)
  args.owner.summariesDirty = true
  return true
}

export function ensureApproximateLookupSummaries(owner: LookupColumnOwner): void {
  if (!owner.summariesDirty) {
    return
  }

  const incompatibleNumericPrefix = new Uint32Array(owner.length + 1)
  const incompatibleTextPrefix = new Uint32Array(owner.length + 1)
  const sortedNumericAscendingBreaks = new Uint32Array(owner.length + 1)
  const sortedNumericDescendingBreaks = new Uint32Array(owner.length + 1)
  const sortedTextAscendingBreaks = new Uint32Array(owner.length + 1)
  const sortedTextDescendingBreaks = new Uint32Array(owner.length + 1)

  for (let offset = 0; offset < owner.length; offset += 1) {
    const kind = decodeComparableKindCode(owner.kindCodes[offset])
    incompatibleNumericPrefix[offset + 1] = incompatibleNumericPrefix[offset]! + (kind === TEXT_KIND || kind === INVALID_KIND ? 1 : 0)
    incompatibleTextPrefix[offset + 1] =
      incompatibleTextPrefix[offset]! + (kind === NUMERIC_KIND || kind === BOOLEAN_KIND || kind === INVALID_KIND ? 1 : 0)
    sortedNumericAscendingBreaks[offset + 1] = sortedNumericAscendingBreaks[offset]!
    sortedNumericDescendingBreaks[offset + 1] = sortedNumericDescendingBreaks[offset]!
    sortedTextAscendingBreaks[offset + 1] = sortedTextAscendingBreaks[offset]!
    sortedTextDescendingBreaks[offset + 1] = sortedTextDescendingBreaks[offset]!

    if (offset === 0) {
      continue
    }

    if (owner.numericValues[offset - 1]! > owner.numericValues[offset]!) {
      sortedNumericAscendingBreaks[offset + 1] = sortedNumericAscendingBreaks[offset + 1]! + 1
    }
    if (owner.numericValues[offset - 1]! < owner.numericValues[offset]!) {
      sortedNumericDescendingBreaks[offset + 1] = sortedNumericDescendingBreaks[offset + 1]! + 1
    }
    if ((owner.textValues[offset - 1] ?? '') > (owner.textValues[offset] ?? '')) {
      sortedTextAscendingBreaks[offset + 1] = sortedTextAscendingBreaks[offset + 1]! + 1
    }
    if ((owner.textValues[offset - 1] ?? '') < (owner.textValues[offset] ?? '')) {
      sortedTextDescendingBreaks[offset + 1] = sortedTextDescendingBreaks[offset + 1]! + 1
    }
  }

  owner.incompatibleNumericPrefix = incompatibleNumericPrefix
  owner.incompatibleTextPrefix = incompatibleTextPrefix
  owner.sortedNumericAscendingBreaks = sortedNumericAscendingBreaks
  owner.sortedNumericDescendingBreaks = sortedNumericDescendingBreaks
  owner.sortedTextAscendingBreaks = sortedTextAscendingBreaks
  owner.sortedTextDescendingBreaks = sortedTextDescendingBreaks
  owner.summariesDirty = false
}

export function supportsNumericApproximateRange(owner: LookupColumnOwner, rowStart: number, rowEnd: number, matchMode: 1 | -1): boolean {
  ensureApproximateLookupSummaries(owner)
  const start = rowStart - owner.rowStart
  const end = rowEnd - owner.rowStart
  if (start < 0 || end >= owner.length) {
    return false
  }
  if (owner.incompatibleNumericPrefix![end + 1]! - owner.incompatibleNumericPrefix![start]! !== 0) {
    return false
  }
  const breaks =
    matchMode === 1
      ? owner.sortedNumericAscendingBreaks![end + 1]! - owner.sortedNumericAscendingBreaks![start + 1]!
      : owner.sortedNumericDescendingBreaks![end + 1]! - owner.sortedNumericDescendingBreaks![start + 1]!
  return breaks === 0
}

export function supportsTextApproximateRange(owner: LookupColumnOwner, rowStart: number, rowEnd: number, matchMode: 1 | -1): boolean {
  ensureApproximateLookupSummaries(owner)
  const start = rowStart - owner.rowStart
  const end = rowEnd - owner.rowStart
  if (start < 0 || end >= owner.length) {
    return false
  }
  if (owner.incompatibleTextPrefix![end + 1]! - owner.incompatibleTextPrefix![start]! !== 0) {
    return false
  }
  const breaks =
    matchMode === 1
      ? owner.sortedTextAscendingBreaks![end + 1]! - owner.sortedTextAscendingBreaks![start + 1]!
      : owner.sortedTextDescendingBreaks![end + 1]! - owner.sortedTextDescendingBreaks![start + 1]!
  return breaks === 0
}

export function sliceOffsetBounds(owner: LookupColumnOwner, rowStart: number, rowEnd: number): { start: number; end: number } | undefined {
  const start = rowStart - owner.rowStart
  const end = rowEnd - owner.rowStart
  if (start < 0 || end >= owner.length || start > end) {
    return undefined
  }
  return { start, end }
}

export function summarizeExactRange(owner: LookupColumnOwner, rowStart: number, rowEnd: number): ExactRangeSummary | undefined {
  const bounds = sliceOffsetBounds(owner, rowStart, rowEnd)
  if (!bounds) {
    return undefined
  }
  let allNumeric = true
  let allText = true
  for (let offset = bounds.start; offset <= bounds.end; offset += 1) {
    const kind = decodeComparableKindCode(owner.kindCodes[offset])
    allNumeric &&= kind === NUMERIC_KIND
    allText &&= kind === TEXT_KIND
    if (!allNumeric && !allText) {
      return {
        comparableKind: 'mixed',
        uniformStart: undefined,
        uniformStep: undefined,
      }
    }
  }
  if (allNumeric) {
    const uniform = detectUniformNumericStepInOwner(owner, bounds.start, bounds.end)
    return {
      comparableKind: 'numeric',
      uniformStart: uniform?.start,
      uniformStep: uniform?.step,
    }
  }
  if (allText) {
    return {
      comparableKind: 'text',
      uniformStart: undefined,
      uniformStep: undefined,
    }
  }
  return {
    comparableKind: 'mixed',
    uniformStart: undefined,
    uniformStep: undefined,
  }
}

export function summarizeApproximateRange(owner: LookupColumnOwner, rowStart: number, rowEnd: number): ApproximateRangeSummary | undefined {
  const bounds = sliceOffsetBounds(owner, rowStart, rowEnd)
  if (!bounds) {
    return undefined
  }
  const numericAscending = supportsNumericApproximateRange(owner, rowStart, rowEnd, 1)
  const numericDescending = supportsNumericApproximateRange(owner, rowStart, rowEnd, -1)
  if (numericAscending || numericDescending) {
    const uniform = detectUniformNumericStepInOwner(owner, bounds.start, bounds.end)
    return {
      comparableKind: 'numeric',
      uniformStart: uniform?.start,
      uniformStep: uniform?.step,
      sortedAscending: numericAscending,
      sortedDescending: numericDescending,
    }
  }
  const textAscending = supportsTextApproximateRange(owner, rowStart, rowEnd, 1)
  const textDescending = supportsTextApproximateRange(owner, rowStart, rowEnd, -1)
  if (textAscending || textDescending) {
    return {
      comparableKind: 'text',
      uniformStart: undefined,
      uniformStep: undefined,
      sortedAscending: textAscending,
      sortedDescending: textDescending,
    }
  }
  return {
    comparableKind: undefined,
    uniformStart: undefined,
    uniformStep: undefined,
    sortedAscending: false,
    sortedDescending: false,
  }
}
