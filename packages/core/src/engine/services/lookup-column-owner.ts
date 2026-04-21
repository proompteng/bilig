import { ValueTag, type CellValue } from '@bilig/protocol'
import type { RuntimeColumnOwner } from './runtime-column-store-service.js'

const MAX_COLUMN_OWNER_SPAN = 1_048_576

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
  readonly firstPositions: Map<string, number>
  readonly lastPositions: Map<string, number>
  readonly rowLists: Map<string, number[]>
  sortedNumericAscendingBreaks: Uint32Array | undefined
  sortedNumericDescendingBreaks: Uint32Array | undefined
  numericUniformBreakOffsets: Uint32Array | undefined
  sortedTextAscendingBreaks: Uint32Array | undefined
  sortedTextDescendingBreaks: Uint32Array | undefined
  incompatibleNumericOffsets: Uint32Array | undefined
  incompatibleTextOffsets: Uint32Array | undefined
  exactNumericIncompatibleOffsets: Uint32Array | undefined
  exactTextIncompatibleOffsets: Uint32Array | undefined
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

function exactLowerBound(rows: Uint32Array, target: number): number {
  let low = 0
  let high = rows.length
  while (low < high) {
    const mid = (low + high) >> 1
    if ((rows[mid] ?? 0) < target) {
      low = mid + 1
    } else {
      high = mid
    }
  }
  return low
}

function hasOffsetInRange(offsets: Uint32Array | undefined, start: number, end: number): boolean {
  if (!offsets || offsets.length === 0 || start > end) {
    return false
  }
  const index = exactLowerBound(offsets, start)
  return index < offsets.length && (offsets[index] ?? 0) <= end
}

function withOffsetMembership(offsets: Uint32Array | undefined, target: number, present: boolean): Uint32Array {
  const source = offsets ?? new Uint32Array(0)
  const index = exactLowerBound(source, target)
  const exists = index < source.length && (source[index] ?? 0) === target
  if (exists === present) {
    return source
  }
  if (present) {
    const result = new Uint32Array(source.length + 1)
    result.set(source.subarray(0, index), 0)
    result[index] = target
    result.set(source.subarray(index), index + 1)
    return result
  }
  const result = new Uint32Array(source.length - 1)
  result.set(source.subarray(0, index), 0)
  result.set(source.subarray(index + 1), index)
  return result
}

function isNumericApproximateIncompatibleKind(kind: ComparableKindCode): boolean {
  return kind === TEXT_KIND || kind === INVALID_KIND
}

function isTextApproximateIncompatibleKind(kind: ComparableKindCode): boolean {
  return kind === NUMERIC_KIND || kind === BOOLEAN_KIND || kind === INVALID_KIND
}

function initializeApproximateLookupSummaries(owner: LookupColumnOwner): void {
  const numericAscendingBreaks: number[] = []
  const numericDescendingBreaks: number[] = []
  const numericUniformBreakOffsets: number[] = []
  const textAscendingBreaks: number[] = []
  const textDescendingBreaks: number[] = []
  const incompatibleNumericOffsets: number[] = []
  const incompatibleTextOffsets: number[] = []
  const exactNumericIncompatibleOffsets: number[] = []
  const exactTextIncompatibleOffsets: number[] = []

  for (let offset = 0; offset < owner.length; offset += 1) {
    const kind = decodeComparableKindCode(owner.kindCodes[offset])
    if (isNumericApproximateIncompatibleKind(kind)) {
      incompatibleNumericOffsets.push(offset)
    }
    if (isTextApproximateIncompatibleKind(kind)) {
      incompatibleTextOffsets.push(offset)
    }
    if (kind !== NUMERIC_KIND) {
      exactNumericIncompatibleOffsets.push(offset)
    }
    if (kind !== TEXT_KIND) {
      exactTextIncompatibleOffsets.push(offset)
    }
    if (offset === 0) {
      continue
    }
    if (owner.numericValues[offset - 1]! > owner.numericValues[offset]!) {
      numericAscendingBreaks.push(offset)
    }
    if (owner.numericValues[offset - 1]! < owner.numericValues[offset]!) {
      numericDescendingBreaks.push(offset)
    }
    if (
      offset >= 2 &&
      owner.numericValues[offset - 2]! - owner.numericValues[offset - 1]! !==
        owner.numericValues[offset - 1]! - owner.numericValues[offset]!
    ) {
      numericUniformBreakOffsets.push(offset)
    }
    if ((owner.textValues[offset - 1] ?? '') > (owner.textValues[offset] ?? '')) {
      textAscendingBreaks.push(offset)
    }
    if ((owner.textValues[offset - 1] ?? '') < (owner.textValues[offset] ?? '')) {
      textDescendingBreaks.push(offset)
    }
  }

  owner.incompatibleNumericOffsets = Uint32Array.from(incompatibleNumericOffsets)
  owner.incompatibleTextOffsets = Uint32Array.from(incompatibleTextOffsets)
  owner.exactNumericIncompatibleOffsets = Uint32Array.from(exactNumericIncompatibleOffsets)
  owner.exactTextIncompatibleOffsets = Uint32Array.from(exactTextIncompatibleOffsets)
  owner.sortedNumericAscendingBreaks = Uint32Array.from(numericAscendingBreaks)
  owner.sortedNumericDescendingBreaks = Uint32Array.from(numericDescendingBreaks)
  owner.numericUniformBreakOffsets = Uint32Array.from(numericUniformBreakOffsets)
  owner.sortedTextAscendingBreaks = Uint32Array.from(textAscendingBreaks)
  owner.sortedTextDescendingBreaks = Uint32Array.from(textDescendingBreaks)
  owner.summariesDirty = false
}

function refreshApproximateLookupCompatibility(owner: LookupColumnOwner, offset: number): void {
  const kind = decodeComparableKindCode(owner.kindCodes[offset])
  owner.incompatibleNumericOffsets = withOffsetMembership(
    owner.incompatibleNumericOffsets,
    offset,
    isNumericApproximateIncompatibleKind(kind),
  )
  owner.incompatibleTextOffsets = withOffsetMembership(owner.incompatibleTextOffsets, offset, isTextApproximateIncompatibleKind(kind))
  owner.exactNumericIncompatibleOffsets = withOffsetMembership(owner.exactNumericIncompatibleOffsets, offset, kind !== NUMERIC_KIND)
  owner.exactTextIncompatibleOffsets = withOffsetMembership(owner.exactTextIncompatibleOffsets, offset, kind !== TEXT_KIND)
}

function refreshApproximateLookupBreaks(owner: LookupColumnOwner, offset: number): void {
  if (offset <= 0 || offset >= owner.length) {
    return
  }
  owner.sortedNumericAscendingBreaks = withOffsetMembership(
    owner.sortedNumericAscendingBreaks,
    offset,
    owner.numericValues[offset - 1]! > owner.numericValues[offset]!,
  )
  owner.sortedNumericDescendingBreaks = withOffsetMembership(
    owner.sortedNumericDescendingBreaks,
    offset,
    owner.numericValues[offset - 1]! < owner.numericValues[offset]!,
  )
  owner.sortedTextAscendingBreaks = withOffsetMembership(
    owner.sortedTextAscendingBreaks,
    offset,
    (owner.textValues[offset - 1] ?? '') > (owner.textValues[offset] ?? ''),
  )
  owner.sortedTextDescendingBreaks = withOffsetMembership(
    owner.sortedTextDescendingBreaks,
    offset,
    (owner.textValues[offset - 1] ?? '') < (owner.textValues[offset] ?? ''),
  )
}

function refreshNumericUniformBreak(owner: LookupColumnOwner, offset: number): void {
  if (offset <= 1 || offset >= owner.length) {
    return
  }
  owner.numericUniformBreakOffsets = withOffsetMembership(
    owner.numericUniformBreakOffsets,
    offset,
    owner.numericValues[offset - 2]! - owner.numericValues[offset - 1]! !== owner.numericValues[offset - 1]! - owner.numericValues[offset]!,
  )
}

function removeOwnerKeyRow(owner: LookupColumnOwner, key: string, row: number): boolean {
  const rows = owner.rowLists.get(key)
  if (rows) {
    const rowIndex = rows.indexOf(row)
    if (rowIndex === -1) {
      return false
    }
    rows.splice(rowIndex, 1)
    if (rows.length === 0) {
      owner.rowLists.delete(key)
      owner.firstPositions.delete(key)
      owner.lastPositions.delete(key)
      return true
    }
    if (rows.length === 1) {
      const onlyRow = rows[0]!
      owner.rowLists.delete(key)
      owner.firstPositions.set(key, onlyRow)
      owner.lastPositions.set(key, onlyRow)
      return true
    }
    owner.firstPositions.set(key, rows[0]!)
    owner.lastPositions.set(key, rows[rows.length - 1]!)
    return true
  }
  const firstRow = owner.firstPositions.get(key)
  const lastRow = owner.lastPositions.get(key)
  if (firstRow === undefined || lastRow === undefined || firstRow !== row || lastRow !== row) {
    return false
  }
  owner.firstPositions.delete(key)
  owner.lastPositions.delete(key)
  return true
}

function insertOwnerKeyRow(owner: LookupColumnOwner, key: string, row: number): void {
  const rows = owner.rowLists.get(key)
  if (rows) {
    insertRowSorted(rows, row)
    owner.firstPositions.set(key, rows[0]!)
    owner.lastPositions.set(key, rows[rows.length - 1]!)
    return
  }
  const firstRow = owner.firstPositions.get(key)
  if (firstRow !== undefined) {
    const nextRows = firstRow < row ? [firstRow, row] : [row, firstRow]
    owner.rowLists.set(key, nextRows)
    owner.firstPositions.set(key, nextRows[0]!)
    owner.lastPositions.set(key, nextRows[nextRows.length - 1]!)
    return
  }
  owner.firstPositions.set(key, row)
  owner.lastPositions.set(key, row)
}

function detectUniformNumericStepInOwner(
  owner: LookupColumnOwner,
  start: number,
  end: number,
): { start: number; step: number } | undefined {
  ensureApproximateLookupSummaries(owner)
  if (end - start < 1) {
    return undefined
  }
  const first = owner.numericValues[start]!
  const step = owner.numericValues[start + 1]! - first
  if (!Number.isFinite(step) || step === 0) {
    return undefined
  }
  if (hasOffsetInRange(owner.numericUniformBreakOffsets, start + 2, end)) {
    return undefined
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
  if (rows && rows.length > 0) {
    if (searchMode === 1) {
      const index = lowerBound(rows, rowStart)
      const row = rows[index]
      return row !== undefined && row <= rowEnd ? row : undefined
    }
    const index = upperBound(rows, rowEnd) - 1
    const row = rows[index]
    return row !== undefined && row >= rowStart ? row : undefined
  }
  const row = searchMode === 1 ? owner.firstPositions.get(key) : owner.lastPositions.get(key)
  return row !== undefined && row >= rowStart && row <= rowEnd ? row : undefined
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
  const textValues: string[] = []

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

  const owner: LookupColumnOwner = {
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
    firstPositions: new Map(),
    lastPositions: new Map(),
    rowLists: new Map(),
    sortedNumericAscendingBreaks: undefined,
    sortedNumericDescendingBreaks: undefined,
    numericUniformBreakOffsets: undefined,
    sortedTextAscendingBreaks: undefined,
    sortedTextDescendingBreaks: undefined,
    incompatibleNumericOffsets: undefined,
    incompatibleTextOffsets: undefined,
    exactNumericIncompatibleOffsets: undefined,
    exactTextIncompatibleOffsets: undefined,
    summariesDirty: true,
  }

  for (let offset = 0; offset < length; offset += 1) {
    const row = minRow + offset
    const key = exactLookupKeyAt(owner, offset)
    insertOwnerKeyRow(owner, key, row)
  }
  initializeApproximateLookupSummaries(owner)
  return owner
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
    if (!removeOwnerKeyRow(args.owner, oldKey, args.write.row)) {
      return false
    }
  }

  if (newKey !== undefined) {
    insertOwnerKeyRow(args.owner, newKey, args.write.row)
  }

  const previousKind = decodeComparableKindCode(args.owner.kindCodes[offset])
  args.owner.kindCodes[offset] = kindCodeForValue(args.write.newValue)
  args.owner.numericValues[offset] = numericValueForValue(args.write.newValue)
  if (args.write.newValue.tag === ValueTag.String) {
    args.owner.textValues[offset] = textValueForValue(args.write.newValue, args.normalizeStringId, args.write.newStringId)
  } else if (offset < args.owner.textValues.length) {
    args.owner.textValues[offset] = ''
  }
  if (args.owner.summariesDirty) {
    initializeApproximateLookupSummaries(args.owner)
  }
  const nextKind = decodeComparableKindCode(args.owner.kindCodes[offset])
  if (previousKind !== nextKind) {
    refreshApproximateLookupCompatibility(args.owner, offset)
  }
  refreshApproximateLookupBreaks(args.owner, offset)
  refreshApproximateLookupBreaks(args.owner, offset + 1)
  refreshNumericUniformBreak(args.owner, offset)
  refreshNumericUniformBreak(args.owner, offset + 1)
  refreshNumericUniformBreak(args.owner, offset + 2)
  args.owner.summariesDirty = false
  return true
}

export function ensureApproximateLookupSummaries(owner: LookupColumnOwner): void {
  if (
    !owner.summariesDirty &&
    owner.incompatibleNumericOffsets &&
    owner.incompatibleTextOffsets &&
    owner.exactNumericIncompatibleOffsets &&
    owner.exactTextIncompatibleOffsets &&
    owner.sortedNumericAscendingBreaks &&
    owner.sortedNumericDescendingBreaks &&
    owner.numericUniformBreakOffsets &&
    owner.sortedTextAscendingBreaks &&
    owner.sortedTextDescendingBreaks
  ) {
    return
  }
  initializeApproximateLookupSummaries(owner)
}

export function supportsNumericApproximateRange(owner: LookupColumnOwner, rowStart: number, rowEnd: number, matchMode: 1 | -1): boolean {
  ensureApproximateLookupSummaries(owner)
  const start = rowStart - owner.rowStart
  const end = rowEnd - owner.rowStart
  if (start < 0 || end >= owner.length) {
    return false
  }
  if (hasOffsetInRange(owner.incompatibleNumericOffsets, start, end)) {
    return false
  }
  return !hasOffsetInRange(matchMode === 1 ? owner.sortedNumericAscendingBreaks : owner.sortedNumericDescendingBreaks, start + 1, end)
}

export function supportsTextApproximateRange(owner: LookupColumnOwner, rowStart: number, rowEnd: number, matchMode: 1 | -1): boolean {
  ensureApproximateLookupSummaries(owner)
  const start = rowStart - owner.rowStart
  const end = rowEnd - owner.rowStart
  if (start < 0 || end >= owner.length) {
    return false
  }
  if (hasOffsetInRange(owner.incompatibleTextOffsets, start, end)) {
    return false
  }
  return !hasOffsetInRange(matchMode === 1 ? owner.sortedTextAscendingBreaks : owner.sortedTextDescendingBreaks, start + 1, end)
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
  ensureApproximateLookupSummaries(owner)
  const bounds = sliceOffsetBounds(owner, rowStart, rowEnd)
  if (!bounds) {
    return undefined
  }
  const allNumeric = !hasOffsetInRange(owner.exactNumericIncompatibleOffsets, bounds.start, bounds.end)
  if (allNumeric) {
    const uniform = detectUniformNumericStepInOwner(owner, bounds.start, bounds.end)
    return {
      comparableKind: 'numeric',
      uniformStart: uniform?.start,
      uniformStep: uniform?.step,
    }
  }
  const allText = !hasOffsetInRange(owner.exactTextIncompatibleOffsets, bounds.start, bounds.end)
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
