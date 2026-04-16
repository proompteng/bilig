import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup.js'

interface LookupReferenceBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue
  numberResult: (value: number) => CellValue
  isError: (value: LookupBuiltinArgument | undefined) => value is Extract<CellValue, { tag: ValueTag.Error }>
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument
  toBoolean: (value: CellValue) => boolean | undefined
  toInteger: (value: CellValue) => number | undefined
  requireCellVector: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  toCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  compareScalars: (left: CellValue, right: CellValue) => number | undefined
  getRangeValue: (range: RangeBuiltinArgument, row: number, col: number) => CellValue
  resolveIndexedExactMatch?: (lookupValue: CellValue, range: RangeBuiltinArgument) => number | undefined
}

function exactMatch(lookupValue: CellValue, range: RangeBuiltinArgument, deps: LookupReferenceBuiltinDeps): number {
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = deps.compareScalars(range.values[index]!, lookupValue)
    if (comparison === 0) {
      return index + 1
    }
  }
  return -1
}

function approximateMatchAscending(lookupValue: CellValue, range: RangeBuiltinArgument, deps: LookupReferenceBuiltinDeps): number {
  let best = -1
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = deps.compareScalars(range.values[index]!, lookupValue)
    if (comparison === undefined) {
      return -1
    }
    if (comparison <= 0) {
      best = index + 1
    } else {
      break
    }
  }
  return best
}

function approximateMatchDescending(lookupValue: CellValue, range: RangeBuiltinArgument, deps: LookupReferenceBuiltinDeps): number {
  let best = -1
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = deps.compareScalars(range.values[index]!, lookupValue)
    if (comparison === undefined) {
      return -1
    }
    if (comparison >= 0) {
      best = index + 1
      continue
    }
    break
  }
  return best
}

export function createLookupReferenceBuiltins(deps: LookupReferenceBuiltinDeps): Record<string, LookupBuiltin> {
  return {
    MATCH: (lookupValue, lookupArray, matchTypeValue = { tag: ValueTag.Number, value: 1 }) => {
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
          ? (deps.resolveIndexedExactMatch?.(lookupValue, rangeOrError) ?? exactMatch(lookupValue, rangeOrError, deps))
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
      const position = shouldApproximate ? approximateMatchAscending(lookupValue, lookupRangeOrError, deps) : exactPosition

      if (position === -1) {
        return deps.errorValue(ErrorCode.NA)
      }

      const resultIndex = position - 1
      return resultRangeOrError.values[resultIndex] ?? deps.errorValue(ErrorCode.NA)
    },
    VLOOKUP: (lookupValue, tableArray, colIndexValue, rangeLookupValue = { tag: ValueTag.Boolean, value: true }) => {
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
      return deps.getRangeValue(tableArray, matchedRow, colIndex - 1)
    },
    HLOOKUP: (lookupValue, tableArray, rowIndexValue, rangeLookupValue = { tag: ValueTag.Boolean, value: true }) => {
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
      return deps.getRangeValue(tableArray, rowIndex - 1, matchedCol)
    },
    XLOOKUP: (
      lookupValue,
      lookupArray,
      returnArray,
      ifNotFound = { tag: ValueTag.Error, code: ErrorCode.NA },
      matchMode = { tag: ValueTag.Number, value: 0 },
      searchMode = { tag: ValueTag.Number, value: 1 },
    ) => {
      if (deps.isRangeArg(lookupValue) || deps.isRangeArg(ifNotFound) || deps.isRangeArg(matchMode) || deps.isRangeArg(searchMode)) {
        return deps.errorValue(ErrorCode.Value)
      }
      const lookupRange = deps.requireCellVector(lookupArray)
      const returnRange = deps.requireCellVector(returnArray)
      if (!deps.isRangeArg(lookupRange)) {
        return lookupRange
      }
      if (!deps.isRangeArg(returnRange)) {
        return returnRange
      }
      if (lookupRange.values.length !== returnRange.values.length) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isError(lookupValue)) {
        return lookupValue
      }
      if (deps.isError(matchMode)) {
        return matchMode
      }
      if (deps.isError(searchMode)) {
        return searchMode
      }

      const matchModeNumber = deps.toInteger(matchMode)
      const searchModeNumber = deps.toInteger(searchMode)
      if ((matchModeNumber ?? 0) !== 0 || (searchModeNumber !== 1 && searchModeNumber !== -1)) {
        return deps.errorValue(ErrorCode.Value)
      }

      if (searchModeNumber === -1) {
        for (let index = lookupRange.values.length - 1; index >= 0; index -= 1) {
          if (deps.compareScalars(lookupRange.values[index]!, lookupValue) === 0) {
            return returnRange.values[index] ?? deps.errorValue(ErrorCode.NA)
          }
        }
        return ifNotFound
      }

      for (let index = 0; index < lookupRange.values.length; index += 1) {
        if (deps.compareScalars(lookupRange.values[index]!, lookupValue) === 0) {
          return returnRange.values[index] ?? deps.errorValue(ErrorCode.NA)
        }
      }
      return ifNotFound
    },
    XMATCH: (
      lookupValue,
      lookupArray,
      matchModeValue = { tag: ValueTag.Number, value: 0 },
      searchModeValue = { tag: ValueTag.Number, value: 1 },
    ) => {
      if (deps.isRangeArg(lookupValue) || deps.isRangeArg(matchModeValue) || deps.isRangeArg(searchModeValue)) {
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
      if (![0, -1, 1].includes(matchMode) || ![1, -1].includes(searchMode)) {
        return deps.errorValue(ErrorCode.Value)
      }

      const values = searchMode === -1 ? rangeOrError.values.toReversed() : rangeOrError.values
      const probe = searchMode === -1 ? { ...rangeOrError, values } : rangeOrError
      const position =
        matchMode === 0
          ? exactMatch(lookupValue, probe, deps)
          : matchMode === 1
            ? approximateMatchAscending(lookupValue, probe, deps)
            : approximateMatchDescending(lookupValue, probe, deps)
      if (position === -1) {
        return deps.errorValue(ErrorCode.NA)
      }
      const normalizedPosition = searchMode === -1 ? rangeOrError.values.length - position + 1 : position
      return deps.numberResult(normalizedPosition)
    },
  }
}
