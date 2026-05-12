import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { ArrayValue } from '../runtime-values.js'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup.js'

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

function approximateLookupAscending(lookupValue: CellValue, range: RangeBuiltinArgument, deps: LookupReferenceBuiltinDeps): number {
  let best = -1
  for (let index = 0; index < range.values.length; index += 1) {
    const value = range.values[index]!
    if (deps.isError(value)) {
      continue
    }
    const comparison = deps.compareScalars(value, lookupValue)
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

function vectorLength(range: RangeBuiltinArgument): number | undefined {
  if (range.refKind !== 'cells' || (range.rows !== 1 && range.cols !== 1)) {
    return undefined
  }
  return range.rows === 1 ? range.cols : range.rows
}

function getVectorValue(range: RangeBuiltinArgument, index: number, deps: LookupReferenceBuiltinDeps): CellValue {
  return range.rows === 1 ? deps.getRangeValue(range, 0, index) : deps.getRangeValue(range, index, 0)
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
  matchMode: number,
  searchMode: number,
  deps: LookupReferenceBuiltinDeps,
): number {
  if (deps.isError(lookupValue)) {
    return -1
  }

  const length = vectorLength(lookupRange)
  if (length === undefined) {
    return -1
  }
  const first = searchMode === -1 ? length - 1 : 0
  const last = searchMode === -1 ? -1 : length
  const step = searchMode === -1 ? -1 : 1

  for (let index = first; index !== last; index += step) {
    const comparison = deps.compareScalars(getVectorValue(lookupRange, index, deps), lookupValue)
    if (comparison === 0) {
      return index
    }
  }
  if (matchMode === 0) {
    return -1
  }

  let bestIndex = -1
  let bestValue: CellValue | undefined
  for (let index = first; index !== last; index += step) {
    const candidate = getVectorValue(lookupRange, index, deps)
    const comparison = deps.compareScalars(candidate, lookupValue)
    if (comparison === undefined) {
      continue
    }
    const qualifies = matchMode === -1 ? comparison < 0 : comparison > 0
    if (!qualifies) {
      continue
    }
    if (bestValue === undefined) {
      bestIndex = index
      bestValue = candidate
      continue
    }
    const bestComparison = deps.compareScalars(candidate, bestValue)
    if (bestComparison === undefined) {
      continue
    }
    if ((matchMode === -1 && bestComparison > 0) || (matchMode === 1 && bestComparison < 0)) {
      bestIndex = index
      bestValue = candidate
    }
  }
  return bestIndex
}

function xlookupScalarResult(
  lookupValue: CellValue,
  lookupRange: RangeBuiltinArgument,
  returnShape: XlookupReturnShape,
  ifNotFound: CellValue,
  matchMode: number,
  searchMode: number,
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
  matchMode: number,
  searchMode: number,
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
      if (![0, -1, 1].includes(matchModeNumber) || (searchModeNumber !== 1 && searchModeNumber !== -1)) {
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
