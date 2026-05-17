import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { ArrayValue } from '../runtime-values.js'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup.js'
import {
  approximateLookupAscending,
  approximateMatchAscending,
  approximateMatchDescending,
  exactMatch,
  findReferenceMatchIndex,
  hasLookupWildcardSyntax,
  vectorLength,
  type LookupReferenceMatchMode,
  type LookupReferenceSearchMode,
} from './lookup-reference-search.js'

interface LookupReferenceBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue
  numberResult: (value: number) => CellValue
  arrayResult: (values: CellValue[], rows: number, cols: number) => ArrayValue
  isError: (value: LookupBuiltinArgument | undefined) => value is Extract<CellValue, { tag: ValueTag.Error }>
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument
  toBoolean: (value: CellValue | undefined) => boolean | undefined
  toInteger: (value: CellValue | undefined) => number | undefined
  requireCellVector: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  toCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  compareScalars: (left: CellValue, right: CellValue) => number | undefined
  getRangeValue: (range: RangeBuiltinArgument, row: number, col: number) => CellValue
  resolveIndexedExactMatch?: (lookupValue: CellValue, range: RangeBuiltinArgument) => number | undefined
}

function coerceLookupReturnValue(value: CellValue, deps: LookupReferenceBuiltinDeps): CellValue {
  return value.tag === ValueTag.Empty ? deps.numberResult(0) : value
}

interface XlookupReturnShape {
  rows: number
  cols: number
  getValues: (matchedIndex: number) => CellValue[]
}

function buildXlookupReturnShape(
  lookupRange: RangeBuiltinArgument,
  returnRange: RangeBuiltinArgument,
  lookupLength: number,
  deps: LookupReferenceBuiltinDeps,
): XlookupReturnShape | CellValue {
  if (returnRange.refKind !== 'cells') {
    return deps.errorValue(ErrorCode.Value)
  }

  if (lookupRange.cols === 1) {
    if (returnRange.rows !== lookupLength) {
      return deps.errorValue(ErrorCode.Value)
    }
    return {
      rows: 1,
      cols: returnRange.cols,
      getValues: (matchedIndex) =>
        Array.from({ length: returnRange.cols }, (_, col) =>
          coerceLookupReturnValue(deps.getRangeValue(returnRange, matchedIndex, col), deps),
        ),
    }
  }

  if (returnRange.cols !== lookupLength) {
    return deps.errorValue(ErrorCode.Value)
  }
  return {
    rows: returnRange.rows,
    cols: 1,
    getValues: (matchedIndex) =>
      Array.from({ length: returnRange.rows }, (_, row) =>
        coerceLookupReturnValue(deps.getRangeValue(returnRange, row, matchedIndex), deps),
      ),
  }
}

function findXlookupMatchIndex(
  lookupValue: CellValue,
  lookupRange: RangeBuiltinArgument,
  matchMode: LookupReferenceMatchMode,
  searchMode: LookupReferenceSearchMode,
  deps: LookupReferenceBuiltinDeps,
): number {
  if (deps.isError(lookupValue)) {
    return -1
  }
  return findReferenceMatchIndex(lookupValue, lookupRange, matchMode, searchMode, deps)
}

function xlookupScalarResult(
  lookupValue: CellValue,
  lookupRange: RangeBuiltinArgument,
  returnShape: XlookupReturnShape,
  ifNotFound: CellValue,
  matchMode: LookupReferenceMatchMode,
  searchMode: LookupReferenceSearchMode,
  deps: LookupReferenceBuiltinDeps,
): CellValue | ArrayValue {
  if (deps.isError(lookupValue)) {
    return lookupValue
  }
  const matchedIndex = findXlookupMatchIndex(lookupValue, lookupRange, matchMode, searchMode, deps)
  if (matchedIndex < 0) {
    return ifNotFound
  }
  const values = returnShape.getValues(matchedIndex)
  return returnShape.rows === 1 && returnShape.cols === 1
    ? (values[0] ?? deps.errorValue(ErrorCode.NA))
    : deps.arrayResult(values, returnShape.rows, returnShape.cols)
}

function xlookupArrayResult(
  lookupValues: RangeBuiltinArgument,
  lookupRange: RangeBuiltinArgument,
  returnShape: XlookupReturnShape,
  ifNotFound: CellValue,
  matchMode: LookupReferenceMatchMode,
  searchMode: LookupReferenceSearchMode,
  deps: LookupReferenceBuiltinDeps,
): CellValue | ArrayValue {
  if (lookupValues.refKind !== 'cells') {
    return deps.errorValue(ErrorCode.Value)
  }

  const resultRows = lookupValues.rows * returnShape.rows
  const resultCols = lookupValues.cols * returnShape.cols
  const values: CellValue[] = []
  for (let lookupRow = 0; lookupRow < lookupValues.rows; lookupRow += 1) {
    const rowSlices: CellValue[][] = []
    for (let lookupCol = 0; lookupCol < lookupValues.cols; lookupCol += 1) {
      const value = deps.getRangeValue(lookupValues, lookupRow, lookupCol)
      if (deps.isError(value)) {
        rowSlices.push(Array.from({ length: returnShape.rows * returnShape.cols }, () => value))
        continue
      }
      const matchedIndex = findXlookupMatchIndex(value, lookupRange, matchMode, searchMode, deps)
      rowSlices.push(
        matchedIndex < 0
          ? Array.from({ length: returnShape.rows * returnShape.cols }, () => ifNotFound)
          : returnShape.getValues(matchedIndex),
      )
    }
    for (let sliceRow = 0; sliceRow < returnShape.rows; sliceRow += 1) {
      for (const slice of rowSlices) {
        values.push(...slice.slice(sliceRow * returnShape.cols, (sliceRow + 1) * returnShape.cols))
      }
    }
  }

  return resultRows === 1 && resultCols === 1
    ? (values[0] ?? deps.errorValue(ErrorCode.NA))
    : deps.arrayResult(values, resultRows, resultCols)
}

export function createLookupReferenceBuiltins(deps: LookupReferenceBuiltinDeps): Record<string, LookupBuiltin> {
  return {
    MATCH: (lookupValue, lookupArray, matchTypeValue = { tag: ValueTag.Number, value: 1 }) => {
      if (lookupValue === undefined || lookupArray === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isRangeArg(lookupValue) || deps.isRangeArg(matchTypeValue)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isError(lookupValue)) {
        return lookupValue
      }
      if (deps.isError(matchTypeValue)) {
        return matchTypeValue
      }

      const rangeOrError = deps.requireCellVector(lookupArray)
      if (!deps.isRangeArg(rangeOrError)) {
        return rangeOrError
      }

      const matchType = deps.toInteger(matchTypeValue)
      if (matchType === undefined || ![-1, 0, 1].includes(matchType)) {
        return deps.errorValue(ErrorCode.Value)
      }

      const position =
        matchType === 0
          ? ((hasLookupWildcardSyntax(lookupValue) ? undefined : deps.resolveIndexedExactMatch?.(lookupValue, rangeOrError)) ??
            exactMatch(lookupValue, rangeOrError, deps, { wildcard: lookupValue.tag === ValueTag.String }))
          : matchType === 1
            ? approximateMatchAscending(lookupValue, rangeOrError, deps)
            : approximateMatchDescending(lookupValue, rangeOrError, deps)

      return position === -1 ? deps.errorValue(ErrorCode.NA) : deps.numberResult(position)
    },
    LOOKUP: (lookupValue, lookupVectorArg, resultVectorArg = lookupVectorArg) => {
      if (deps.isRangeArg(lookupValue) || lookupValue === undefined || resultVectorArg === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }

      const existingError = deps.isError(lookupValue)
        ? lookupValue
        : deps.isError(lookupVectorArg)
          ? lookupVectorArg
          : deps.isError(resultVectorArg)
            ? resultVectorArg
            : undefined
      if (existingError) {
        return existingError
      }

      const lookupRangeOrError = deps.toCellRange(lookupVectorArg)
      const resultRangeOrError = deps.toCellRange(resultVectorArg)
      if (!deps.isRangeArg(lookupRangeOrError)) {
        return lookupRangeOrError
      }
      if (!deps.isRangeArg(resultRangeOrError)) {
        return resultRangeOrError
      }

      if (lookupRangeOrError.rows !== 1 && lookupRangeOrError.cols !== 1) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (resultRangeOrError.rows !== 1 && resultRangeOrError.cols !== 1) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (lookupRangeOrError.values.length !== resultRangeOrError.values.length) {
        return deps.errorValue(ErrorCode.Value)
      }

      const exactPosition = exactMatch(lookupValue, lookupRangeOrError, deps)
      const shouldApproximate = exactPosition === -1 && lookupValue.tag === ValueTag.Number
      const position = shouldApproximate ? approximateLookupAscending(lookupValue, lookupRangeOrError, deps) : exactPosition

      if (position === -1) {
        return deps.errorValue(ErrorCode.NA)
      }

      const resultIndex = position - 1
      const result = resultRangeOrError.values[resultIndex]
      return result === undefined ? deps.errorValue(ErrorCode.NA) : coerceLookupReturnValue(result, deps)
    },
    VLOOKUP: (lookupValue, tableArray, colIndexValue, rangeLookupValue = { tag: ValueTag.Boolean, value: true }) => {
      if (lookupValue === undefined || tableArray === undefined || colIndexValue === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isRangeArg(lookupValue)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (!deps.isRangeArg(tableArray) || tableArray.refKind !== 'cells') {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isRangeArg(colIndexValue) || deps.isRangeArg(rangeLookupValue)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isError(lookupValue)) {
        return lookupValue
      }
      if (deps.isError(colIndexValue)) {
        return colIndexValue
      }
      if (deps.isError(rangeLookupValue)) {
        return rangeLookupValue
      }

      const colIndex = deps.toInteger(colIndexValue)
      const rangeLookup = deps.toBoolean(rangeLookupValue)
      if (colIndex === undefined || colIndex < 1 || colIndex > tableArray.cols || rangeLookup === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }

      let matchedRow = -1
      for (let row = 0; row < tableArray.rows; row += 1) {
        const comparison = deps.compareScalars(deps.getRangeValue(tableArray, row, 0), lookupValue)
        if (comparison === undefined) {
          if (!rangeLookup) {
            continue
          }
          return deps.errorValue(ErrorCode.Value)
        }
        if (comparison === 0) {
          matchedRow = row
          break
        }
        if (rangeLookup && comparison < 0) {
          matchedRow = row
          continue
        }
        if (rangeLookup && comparison > 0) {
          break
        }
      }

      if (matchedRow === -1) {
        return deps.errorValue(ErrorCode.NA)
      }
      return coerceLookupReturnValue(deps.getRangeValue(tableArray, matchedRow, colIndex - 1), deps)
    },
    HLOOKUP: (lookupValue, tableArray, rowIndexValue, rangeLookupValue = { tag: ValueTag.Boolean, value: true }) => {
      if (lookupValue === undefined || tableArray === undefined || rowIndexValue === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isRangeArg(lookupValue)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (!deps.isRangeArg(tableArray) || tableArray.refKind !== 'cells') {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isRangeArg(rowIndexValue) || deps.isRangeArg(rangeLookupValue)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isError(lookupValue)) {
        return lookupValue
      }
      if (deps.isError(rowIndexValue)) {
        return rowIndexValue
      }
      if (deps.isError(rangeLookupValue)) {
        return rangeLookupValue
      }

      const rowIndex = deps.toInteger(rowIndexValue)
      const rangeLookup = deps.toBoolean(rangeLookupValue)
      if (rowIndex === undefined || rowIndex < 1 || rowIndex > tableArray.rows || rangeLookup === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }

      let matchedCol = -1
      for (let col = 0; col < tableArray.cols; col += 1) {
        const comparison = deps.compareScalars(deps.getRangeValue(tableArray, 0, col), lookupValue)
        if (comparison === undefined) {
          if (!rangeLookup) {
            continue
          }
          return deps.errorValue(ErrorCode.Value)
        }
        if (comparison === 0) {
          matchedCol = col
          break
        }
        if (rangeLookup && comparison < 0) {
          matchedCol = col
          continue
        }
        if (rangeLookup && comparison > 0) {
          break
        }
      }

      if (matchedCol === -1) {
        return deps.errorValue(ErrorCode.NA)
      }
      return coerceLookupReturnValue(deps.getRangeValue(tableArray, rowIndex - 1, matchedCol), deps)
    },
    XLOOKUP: (
      lookupValue,
      lookupArray,
      returnArray,
      ifNotFound = { tag: ValueTag.Error, code: ErrorCode.NA },
      matchMode = { tag: ValueTag.Number, value: 0 },
      searchMode = { tag: ValueTag.Number, value: 1 },
    ) => {
      if (lookupValue === undefined || lookupArray === undefined || returnArray === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isRangeArg(ifNotFound) || deps.isRangeArg(matchMode) || deps.isRangeArg(searchMode)) {
        return deps.errorValue(ErrorCode.Value)
      }
      const lookupRange = deps.requireCellVector(lookupArray)
      const returnRange = deps.toCellRange(returnArray)
      if (!deps.isRangeArg(lookupRange)) {
        return lookupRange
      }
      if (!deps.isRangeArg(returnRange)) {
        return returnRange
      }
      if (deps.isError(matchMode)) {
        return matchMode
      }
      if (deps.isError(searchMode)) {
        return searchMode
      }

      const matchModeNumber = deps.toInteger(matchMode)
      const searchModeNumber = deps.toInteger(searchMode)
      if (matchModeNumber === undefined || searchModeNumber === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (
        !isLookupReferenceMatchMode(matchModeNumber) ||
        !isLookupReferenceSearchMode(searchModeNumber) ||
        (matchModeNumber === 2 && (searchModeNumber === 2 || searchModeNumber === -2))
      ) {
        return deps.errorValue(ErrorCode.Value)
      }

      const lookupLength = vectorLength(lookupRange)
      if (lookupLength === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      const returnShape = buildXlookupReturnShape(lookupRange, returnRange, lookupLength, deps)
      if (!('getValues' in returnShape)) {
        return returnShape
      }
      if (lookupValue === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }

      return deps.isRangeArg(lookupValue)
        ? xlookupArrayResult(lookupValue, lookupRange, returnShape, ifNotFound, matchModeNumber, searchModeNumber, deps)
        : xlookupScalarResult(lookupValue, lookupRange, returnShape, ifNotFound, matchModeNumber, searchModeNumber, deps)
    },
    XMATCH: (
      lookupValue,
      lookupArray,
      matchModeValue = { tag: ValueTag.Number, value: 0 },
      searchModeValue = { tag: ValueTag.Number, value: 1 },
    ) => {
      if (
        lookupValue === undefined ||
        deps.isRangeArg(lookupValue) ||
        deps.isRangeArg(matchModeValue) ||
        deps.isRangeArg(searchModeValue)
      ) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isError(lookupValue)) {
        return lookupValue
      }
      if (deps.isError(matchModeValue)) {
        return matchModeValue
      }
      if (deps.isError(searchModeValue)) {
        return searchModeValue
      }
      const rangeOrError = deps.requireCellVector(lookupArray)
      if (!deps.isRangeArg(rangeOrError)) {
        return rangeOrError
      }
      const matchMode = deps.toInteger(matchModeValue)
      const searchMode = deps.toInteger(searchModeValue)
      if (matchMode === undefined || searchMode === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (
        !isLookupReferenceMatchMode(matchMode) ||
        !isLookupReferenceSearchMode(searchMode) ||
        (matchMode === 2 && (searchMode === 2 || searchMode === -2))
      ) {
        return deps.errorValue(ErrorCode.Value)
      }

      const position = xmatchPosition(lookupValue, rangeOrError, matchMode, searchMode, deps)
      if (position === -1) {
        return deps.errorValue(ErrorCode.NA)
      }
      return deps.numberResult(position)
    },
  }
}

function xmatchPosition(
  lookupValue: CellValue,
  range: RangeBuiltinArgument,
  matchMode: LookupReferenceMatchMode,
  searchMode: LookupReferenceSearchMode,
  deps: LookupReferenceBuiltinDeps,
): number {
  if (searchMode === 2 || searchMode === -2 || matchMode === 0 || matchMode === 2) {
    const index = findXlookupMatchIndex(lookupValue, range, matchMode, searchMode, deps)
    return index === -1 ? -1 : index + 1
  }

  const values = searchMode === -1 ? range.values.toReversed() : range.values
  const probe = searchMode === -1 ? { ...range, values } : range
  const position =
    matchMode === 1 ? approximateMatchAscending(lookupValue, probe, deps) : approximateMatchDescending(lookupValue, probe, deps)
  return position === -1 || searchMode !== -1 ? position : range.values.length - position + 1
}

function isLookupReferenceMatchMode(value: number): value is LookupReferenceMatchMode {
  return value === 0 || value === -1 || value === 1 || value === 2
}

function isLookupReferenceSearchMode(value: number): value is LookupReferenceSearchMode {
  return value === 1 || value === -1 || value === 2 || value === -2
}
