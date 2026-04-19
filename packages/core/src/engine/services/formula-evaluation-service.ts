import { Effect } from 'effect'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import {
  createLookupBuiltinResolver,
  evaluatePlanResult,
  formatAddress,
  isArrayValue,
  lowerToPlan,
  type EvaluationContext,
  type EvaluationResult,
  type FormulaNode,
  type RangeBuiltinArgument,
  parseCellAddress,
  parseRangeAddress,
} from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import { definedNameValueToCellValue } from '../../engine-metadata-utils.js'
import { emptyValue, errorValue } from '../../engine-value-utils.js'
import type {
  EngineRuntimeState,
  PreparedApproximateVectorLookup,
  PreparedExactVectorLookup,
  RuntimeDirectScalarOperand,
  RuntimeFormula,
  SpillMaterialization,
} from '../runtime-state.js'
import { EngineFormulaEvaluationError } from '../errors.js'
import type { CriterionRangeCacheService } from './criterion-range-cache-service.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnSlice } from './runtime-column-store-service.js'
import type { RangeAggregateCacheService } from './range-aggregate-cache-service.js'
import type { SortedColumnSearchService } from './sorted-column-search-service.js'

function decodeErrorCode(rawCode: number | undefined): ErrorCode {
  return rawCode ?? ErrorCode.None
}

export interface EngineFormulaEvaluationService {
  readonly evaluateDirectLookupFormula: (cellIndex: number) => Effect.Effect<number[] | undefined, EngineFormulaEvaluationError>
  readonly evaluateDirectLookupFormulaNow: (cellIndex: number) => number[] | undefined
  readonly evaluateUnsupportedFormula: (cellIndex: number) => Effect.Effect<number[], EngineFormulaEvaluationError>
  readonly resolveStructuredReference: (
    tableName: string,
    columnName: string,
  ) => Effect.Effect<FormulaNode | undefined, EngineFormulaEvaluationError>
  readonly resolveSpillReference: (
    currentSheetName: string,
    sheetName: string | undefined,
    address: string,
  ) => Effect.Effect<FormulaNode | undefined, EngineFormulaEvaluationError>
  readonly resolveMultipleOperations: (request: {
    formulaSheetName: string
    formulaAddress: string
    rowCellSheetName: string
    rowCellAddress: string
    rowReplacementSheetName: string
    rowReplacementAddress: string
    columnCellSheetName?: string
    columnCellAddress?: string
    columnReplacementSheetName?: string
    columnReplacementAddress?: string
  }) => Effect.Effect<CellValue, EngineFormulaEvaluationError>
}

function evaluationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function decodeRuntimeTag(rawTag: number | undefined): ValueTag {
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

function referenceReplacementKey(sheetName: string, address: string): string {
  return `${sheetName.trim().toUpperCase()}!${address.trim().toUpperCase()}`
}

function cellValuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false
  }
  switch (left.tag) {
    case ValueTag.Empty:
      return true
    case ValueTag.Number:
      return right.tag === ValueTag.Number && left.value === right.value
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code
  }
}

export function createEngineFormulaEvaluationService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings' | 'formulas' | 'getUseColumnIndex'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly criterionCache: CriterionRangeCacheService
  readonly aggregateCache: RangeAggregateCacheService
  readonly exactLookup: Pick<ExactColumnIndexService, 'findVectorMatch' | 'prepareVectorLookup' | 'findPreparedVectorMatch'>
  readonly sortedLookup: Pick<SortedColumnSearchService, 'findVectorMatch' | 'prepareVectorLookup' | 'findPreparedVectorMatch'>
  readonly materializeSpill: (cellIndex: number, arrayValue: { values: CellValue[]; rows: number; cols: number }) => SpillMaterialization
  readonly clearOwnedSpill: (cellIndex: number) => number[]
  readonly resolvePivotData: (
    sheetName: string,
    address: string,
    dataField: string,
    filters: ReadonlyArray<{ field: string; item: CellValue }>,
  ) => CellValue
}): EngineFormulaEvaluationService {
  const emptyChangedCellIndices: number[] = []
  const reusableDirectNumberResult: CellValue = { tag: ValueTag.Number, value: 0 }
  const reusableDirectErrorResult: CellValue = { tag: ValueTag.Error, code: ErrorCode.None }

  const directNumberResult = (value: number): CellValue => {
    reusableDirectNumberResult.value = value
    return reusableDirectNumberResult
  }

  const directErrorResult = (code: ErrorCode): CellValue => {
    reusableDirectErrorResult.code = code
    return reusableDirectErrorResult
  }

  const readCellValue = (sheetName: string, address: string): CellValue => {
    const parsed = parseCellAddress(address, sheetName)
    return args.runtimeColumnStore.readCellValue(sheetName, parsed.row, parsed.col)
  }

  const readCellValueByIndex = (cellIndex: number): CellValue => {
    return args.state.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : args.state.strings.get(stringId)))
  }

  const numericLikeValueAt = (slice: RuntimeColumnSlice, offset: number): number | undefined => {
    const tag = decodeRuntimeTag(slice.tags[offset])
    switch (tag) {
      case ValueTag.Number:
        return slice.numbers[offset] ?? 0
      case ValueTag.Boolean:
        return (slice.numbers[offset] ?? 0) !== 0 ? 1 : 0
      case ValueTag.Empty:
        return 0
      case ValueTag.String:
      case ValueTag.Error:
      default:
        return undefined
    }
  }

  const strictNumericAggregateCandidateAt = (slice: RuntimeColumnSlice, offset: number): number | undefined => {
    return slice.tags[offset] === ValueTag.Number ? (slice.numbers[offset] ?? 0) : undefined
  }

  const refreshDirectExactLookup = (
    directLookup: Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'exact' }>,
  ): PreparedExactVectorLookup => {
    const prepared = directLookup.prepared
    const lookupSheet = args.state.workbook.getSheet(prepared.sheetName)
    const currentColumnVersion = lookupSheet?.columnVersions[prepared.col] ?? 0
    const currentStructureVersion = lookupSheet?.structureVersion ?? 0
    if (currentColumnVersion === prepared.columnVersion && currentStructureVersion === prepared.structureVersion) {
      return prepared
    }
    const refreshed = args.exactLookup.prepareVectorLookup({
      sheetName: prepared.sheetName,
      rowStart: prepared.rowStart,
      rowEnd: prepared.rowEnd,
      col: prepared.col,
    })
    directLookup.prepared = refreshed
    return refreshed
  }

  const refreshDirectApproximateLookup = (
    directLookup: Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'approximate' }>,
  ): PreparedApproximateVectorLookup => {
    const prepared = directLookup.prepared
    const lookupSheet = args.state.workbook.getSheet(prepared.sheetName)
    const currentColumnVersion = lookupSheet?.columnVersions[prepared.col] ?? 0
    const currentStructureVersion = lookupSheet?.structureVersion ?? 0
    if (currentColumnVersion === prepared.columnVersion && currentStructureVersion === prepared.structureVersion) {
      return prepared
    }
    const refreshed = args.sortedLookup.prepareVectorLookup({
      sheetName: prepared.sheetName,
      rowStart: prepared.rowStart,
      rowEnd: prepared.rowEnd,
      col: prepared.col,
    })
    directLookup.prepared = refreshed
    return refreshed
  }

  const refreshDirectExactUniformLookup = (
    formula: RuntimeFormula,
    directLookup: Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'exact-uniform-numeric' }>,
  ):
    | Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'exact-uniform-numeric' }>
    | Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'exact' }> => {
    const lookupSheet = args.state.workbook.getSheet(directLookup.sheetName)
    const currentColumnVersion = lookupSheet?.columnVersions[directLookup.col] ?? 0
    const currentStructureVersion = lookupSheet?.structureVersion ?? 0
    if (currentColumnVersion === directLookup.columnVersion && currentStructureVersion === directLookup.structureVersion) {
      return directLookup
    }
    const refreshed = args.exactLookup.prepareVectorLookup({
      sheetName: directLookup.sheetName,
      rowStart: directLookup.rowStart,
      rowEnd: directLookup.rowEnd,
      col: directLookup.col,
    })
    if (refreshed.comparableKind === 'numeric' && refreshed.uniformStart !== undefined && refreshed.uniformStep !== undefined) {
      directLookup.length = refreshed.length
      directLookup.columnVersion = refreshed.columnVersion
      directLookup.structureVersion = refreshed.structureVersion
      directLookup.sheetColumnVersions = refreshed.sheetColumnVersions
      directLookup.start = refreshed.uniformStart
      directLookup.step = refreshed.uniformStep
      return directLookup
    }
    const fallback = {
      kind: 'exact' as const,
      operandCellIndex: directLookup.operandCellIndex,
      prepared: refreshed,
      searchMode: directLookup.searchMode,
    }
    formula.directLookup = fallback
    return fallback
  }

  const refreshDirectApproximateUniformLookup = (
    formula: RuntimeFormula,
    directLookup: Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'approximate-uniform-numeric' }>,
  ):
    | Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'approximate-uniform-numeric' }>
    | Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'approximate' }> => {
    const lookupSheet = args.state.workbook.getSheet(directLookup.sheetName)
    const currentColumnVersion = lookupSheet?.columnVersions[directLookup.col] ?? 0
    const currentStructureVersion = lookupSheet?.structureVersion ?? 0
    if (currentColumnVersion === directLookup.columnVersion && currentStructureVersion === directLookup.structureVersion) {
      return directLookup
    }
    const refreshed = args.sortedLookup.prepareVectorLookup({
      sheetName: directLookup.sheetName,
      rowStart: directLookup.rowStart,
      rowEnd: directLookup.rowEnd,
      col: directLookup.col,
    })
    if (refreshed.comparableKind === 'numeric' && refreshed.uniformStart !== undefined && refreshed.uniformStep !== undefined) {
      directLookup.length = refreshed.length
      directLookup.columnVersion = refreshed.columnVersion
      directLookup.structureVersion = refreshed.structureVersion
      directLookup.sheetColumnVersions = refreshed.sheetColumnVersions
      directLookup.start = refreshed.uniformStart
      directLookup.step = refreshed.uniformStep
      return directLookup
    }
    const fallback = {
      kind: 'approximate' as const,
      operandCellIndex: directLookup.operandCellIndex,
      prepared: refreshed,
      matchMode: directLookup.matchMode,
    }
    formula.directLookup = fallback
    return fallback
  }

  const readRangeValues = (
    sheetName: string,
    start: string,
    end: string,
    refKind: 'cells' | 'rows' | 'cols',
    replacements?: ReadonlyMap<string, { sheetName: string; address: string }>,
    visiting?: Set<string>,
  ): CellValue[] => {
    if (refKind !== 'cells') {
      return []
    }
    const range = parseRangeAddress(`${start}:${end}`, sheetName)
    if (range.kind !== 'cells') {
      return []
    }
    const values: CellValue[] = []
    if (!replacements || !visiting) {
      return args.runtimeColumnStore.readRangeValues({
        sheetName,
        rowStart: range.start.row,
        rowEnd: range.end.row,
        colStart: range.start.col,
        colEnd: range.end.col,
      })
    }
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let col = range.start.col; col <= range.end.col; col += 1) {
        values.push(evaluateCellWithReferenceReplacements(sheetName, formatAddress(row, col), replacements, visiting))
      }
    }
    return values
  }

  const resolveIndexedExactMatch = (lookupValue: CellValue, range: RangeBuiltinArgument): number | undefined => {
    if (!args.state.getUseColumnIndex() || range.refKind !== 'cells' || range.cols !== 1) {
      return undefined
    }
    if (!range.sheetName || !range.start || !range.end) {
      return undefined
    }
    const result = args.exactLookup.findVectorMatch({
      lookupValue,
      sheetName: range.sheetName,
      start: range.start,
      end: range.end,
      searchMode: 1,
    })
    return result.handled ? result.position : undefined
  }

  const lookupBuiltinResolver = createLookupBuiltinResolver({
    resolveIndexedExactMatch,
  })

  const resolveExactVectorMatch = (
    _formula: RuntimeFormula,
    request: {
      lookupValue: CellValue
      sheetName: string
      start: string
      end: string
      startRow: number
      endRow: number
      startCol: number
      endCol: number
      searchMode: 1 | -1
    },
  ) => {
    return args.exactLookup.findVectorMatch(request)
  }

  const resolveApproximateVectorMatch = (
    _formula: RuntimeFormula,
    request: {
      lookupValue: CellValue
      sheetName: string
      start: string
      end: string
      startRow: number
      endRow: number
      startCol: number
      endCol: number
      matchMode: 1 | -1
    },
  ) => {
    return args.sortedLookup.findVectorMatch(request)
  }

  const tryEvaluateDirectVectorLookup = (formula: RuntimeFormula): CellValue | undefined => {
    const directLookup = formula.directLookup
    if (!directLookup) {
      return undefined
    }
    const cellStore = args.state.workbook.cellStore
    if (directLookup.kind === 'exact-uniform-numeric') {
      const refreshed = refreshDirectExactUniformLookup(formula, directLookup)
      if (refreshed.kind !== 'exact-uniform-numeric') {
        return tryEvaluateDirectVectorLookup(formula)
      }
      const tag = cellStore.tags[refreshed.operandCellIndex]
      if (tag === ValueTag.Error) {
        return undefined
      }
      if (tag !== ValueTag.Number) {
        return directErrorResult(ErrorCode.NA)
      }
      const numericValue = Object.is(cellStore.numbers[refreshed.operandCellIndex] ?? 0, -0)
        ? 0
        : (cellStore.numbers[refreshed.operandCellIndex] ?? 0)
      if (refreshed.step === 1) {
        if (!Number.isInteger(numericValue)) {
          return directErrorResult(ErrorCode.NA)
        }
        const position = numericValue - refreshed.start + 1
        return position >= 1 && position <= refreshed.length ? directNumberResult(position) : directErrorResult(ErrorCode.NA)
      }
      if (refreshed.step === -1) {
        if (!Number.isInteger(numericValue)) {
          return directErrorResult(ErrorCode.NA)
        }
        const position = refreshed.start - numericValue + 1
        return position >= 1 && position <= refreshed.length ? directNumberResult(position) : directErrorResult(ErrorCode.NA)
      }
      const relative = (numericValue - refreshed.start) / refreshed.step
      return Number.isInteger(relative) && relative >= 0 && relative < refreshed.length
        ? directNumberResult(relative + 1)
        : directErrorResult(ErrorCode.NA)
    }
    if (directLookup.kind === 'exact') {
      const prepared = refreshDirectExactLookup(directLookup)
      const cellIndex = directLookup.operandCellIndex
      const result = args.exactLookup.findPreparedVectorMatch({
        lookupValue: readCellValueByIndex(cellIndex),
        prepared,
        searchMode: directLookup.searchMode,
      })
      if (!result.handled) {
        return undefined
      }
      return result.position === undefined ? directErrorResult(ErrorCode.NA) : directNumberResult(result.position)
    }
    if (directLookup.kind === 'approximate-uniform-numeric') {
      const refreshed = refreshDirectApproximateUniformLookup(formula, directLookup)
      if (refreshed.kind !== 'approximate-uniform-numeric') {
        return tryEvaluateDirectVectorLookup(formula)
      }
      const tag = cellStore.tags[refreshed.operandCellIndex]
      let lookupValue = 0
      switch (tag) {
        case undefined:
        case ValueTag.Empty:
          lookupValue = 0
          break
        case ValueTag.Number:
          lookupValue = Object.is(cellStore.numbers[refreshed.operandCellIndex] ?? 0, -0)
            ? 0
            : (cellStore.numbers[refreshed.operandCellIndex] ?? 0)
          break
        case ValueTag.Boolean:
          lookupValue = (cellStore.numbers[refreshed.operandCellIndex] ?? 0) !== 0 ? 1 : 0
          break
        case ValueTag.Error:
        case ValueTag.String:
          return undefined
      }
      const lastValue = refreshed.start + refreshed.step * (refreshed.length - 1)
      if (refreshed.matchMode === 1 && refreshed.step > 0) {
        if (lookupValue < refreshed.start) {
          return directErrorResult(ErrorCode.NA)
        }
        if (lookupValue >= lastValue) {
          return directNumberResult(refreshed.length)
        }
        const position = Math.floor((lookupValue - refreshed.start) / refreshed.step) + 1
        return directNumberResult(Math.min(refreshed.length, Math.max(1, position)))
      }
      if (refreshed.matchMode === -1 && refreshed.step < 0) {
        if (lookupValue > refreshed.start) {
          return directErrorResult(ErrorCode.NA)
        }
        if (lookupValue <= lastValue) {
          return directNumberResult(refreshed.length)
        }
        const position = Math.floor((refreshed.start - lookupValue) / -refreshed.step) + 1
        return directNumberResult(Math.min(refreshed.length, Math.max(1, position)))
      }
      return undefined
    }
    const prepared = refreshDirectApproximateLookup(directLookup)
    const cellIndex = directLookup.operandCellIndex
    const result = args.sortedLookup.findPreparedVectorMatch({
      lookupValue: readCellValueByIndex(cellIndex),
      prepared,
      matchMode: directLookup.matchMode,
    })
    if (!result.handled) {
      return undefined
    }
    return result.position === undefined ? directErrorResult(ErrorCode.NA) : directNumberResult(result.position)
  }

  const tryEvaluateDirectCriteriaAggregate = (formula: RuntimeFormula): CellValue | undefined => {
    const directCriteria = formula.directCriteria
    if (!directCriteria) {
      return undefined
    }

    const resolvedPairs = directCriteria.criteriaPairs.map((pair) => ({
      range: pair.range,
      criteria: pair.criterion.kind === 'literal' ? pair.criterion.value : readCellValueByIndex(pair.criterion.cellIndex),
    }))
    const criterionError = resolvedPairs.find((pair) => pair.criteria.tag === ValueTag.Error)?.criteria
    if (criterionError) {
      return criterionError
    }

    const matches = args.criterionCache.getOrBuildMatchingRows({
      criteriaPairs: resolvedPairs,
    })
    if ('tag' in matches) {
      return matches
    }

    if (directCriteria.aggregateKind === 'count') {
      return directNumberResult(matches.length)
    }

    const aggregateRange = directCriteria.aggregateRange
    if (!aggregateRange) {
      return undefined
    }
    const aggregateSlice = args.runtimeColumnStore.getColumnSlice({
      sheetName: aggregateRange.sheetName,
      rowStart: aggregateRange.rowStart,
      rowEnd: aggregateRange.rowEnd,
      col: aggregateRange.col,
    })

    if (directCriteria.aggregateKind === 'sum') {
      let sum = 0
      for (let index = 0; index < matches.length; index += 1) {
        sum += numericLikeValueAt(aggregateSlice, matches.rows[index]!) ?? 0
      }
      return directNumberResult(sum)
    }

    if (directCriteria.aggregateKind === 'average') {
      let count = 0
      let sum = 0
      for (let index = 0; index < matches.length; index += 1) {
        const numeric = numericLikeValueAt(aggregateSlice, matches.rows[index]!)
        if (numeric === undefined) {
          continue
        }
        count += 1
        sum += numeric
      }
      return count === 0 ? directErrorResult(ErrorCode.Div0) : directNumberResult(sum / count)
    }

    if (directCriteria.aggregateKind === 'min') {
      let minimum = Number.POSITIVE_INFINITY
      for (let index = 0; index < matches.length; index += 1) {
        const numeric = strictNumericAggregateCandidateAt(aggregateSlice, matches.rows[index]!)
        if (numeric === undefined) {
          continue
        }
        minimum = Math.min(minimum, numeric)
      }
      return directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum)
    }

    let maximum = Number.NEGATIVE_INFINITY
    for (let index = 0; index < matches.length; index += 1) {
      const numeric = strictNumericAggregateCandidateAt(aggregateSlice, matches.rows[index]!)
      if (numeric === undefined) {
        continue
      }
      maximum = Math.max(maximum, numeric)
    }
    return directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum)
  }

  const tryEvaluateDirectAggregate = (formula: RuntimeFormula): CellValue | undefined => {
    const directAggregate = formula.directAggregate
    if (!directAggregate) {
      return undefined
    }
    if (formula.dependencyIndices.length > 0) {
      const aggregateSheet = args.state.workbook.getSheet(directAggregate.sheetName)
      if (!aggregateSheet) {
        return undefined
      }
      let sum = 0
      let count = 0
      let averageCount = 0
      let minimum = Number.POSITIVE_INFINITY
      let maximum = Number.NEGATIVE_INFINITY
      for (let row = directAggregate.rowStart; row <= directAggregate.rowEnd; row += 1) {
        const memberCellIndex = aggregateSheet.grid.get(row, directAggregate.col)
        const value: CellValue = memberCellIndex === -1 ? { tag: ValueTag.Empty } : readCellValueByIndex(memberCellIndex)
        switch (value.tag) {
          case ValueTag.Number:
            sum += value.value
            count += 1
            averageCount += 1
            minimum = Math.min(minimum, value.value)
            maximum = Math.max(maximum, value.value)
            break
          case ValueTag.Boolean: {
            const booleanNumber = value.value ? 1 : 0
            sum += booleanNumber
            count += 1
            averageCount += 1
            minimum = Math.min(minimum, booleanNumber)
            maximum = Math.max(maximum, booleanNumber)
            break
          }
          case ValueTag.Empty:
            averageCount += 1
            minimum = Math.min(minimum, 0)
            maximum = Math.max(maximum, 0)
            break
          case ValueTag.Error:
            if (directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average') {
              return directErrorResult(value.code)
            }
            break
          case ValueTag.String:
            break
        }
      }
      if (directAggregate.aggregateKind === 'sum') {
        return directNumberResult(sum)
      }
      if (directAggregate.aggregateKind === 'count') {
        return directNumberResult(count)
      }
      if (directAggregate.aggregateKind === 'average') {
        return averageCount === 0 ? directErrorResult(ErrorCode.Div0) : directNumberResult(sum / averageCount)
      }
      if (directAggregate.aggregateKind === 'min') {
        return directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum)
      }
      return directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum)
    }
    // SUM/AVERAGE ranges should reuse any compatible lower-start prefix to
    // avoid rescanning shifted windows, while still allowing narrower anchors
    // when no compatible reusable prefix exists.
    const sharedPrefixStart = directAggregate.aggregateKind === 'count' ? 0 : directAggregate.rowStart
    const prefix = args.aggregateCache.getOrBuildPrefix({
      sheetName: directAggregate.sheetName,
      rowStart: sharedPrefixStart,
      rowEnd: directAggregate.rowEnd,
      col: directAggregate.col,
    })
    const endOffset = directAggregate.rowEnd - prefix.rowStart
    const startOffset = directAggregate.rowStart - prefix.rowStart - 1
    const errorCode = prefix.prefixErrorCodes[endOffset]
    const prefixSum = prefix.prefixSums[endOffset] ?? 0
    const prefixCount = prefix.prefixCount[endOffset] ?? 0
    const prefixAverageCount = prefix.prefixAverageCount[endOffset] ?? 0
    const prefixErrorCount = prefix.prefixErrorCounts[endOffset] ?? 0
    const sum = startOffset >= 0 ? prefixSum - (prefix.prefixSums[startOffset] ?? 0) : prefixSum
    const count = startOffset >= 0 ? prefixCount - (prefix.prefixCount[startOffset] ?? 0) : prefixCount
    const averageCount = startOffset >= 0 ? prefixAverageCount - (prefix.prefixAverageCount[startOffset] ?? 0) : prefixAverageCount
    const errorCount = startOffset >= 0 ? prefixErrorCount - (prefix.prefixErrorCounts[startOffset] ?? 0) : prefixErrorCount
    if (
      errorCode !== ErrorCode.None &&
      errorCount > 0 &&
      (directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average')
    ) {
      return startOffset >= 0 ? undefined : directErrorResult(decodeErrorCode(errorCode))
    }
    if (directAggregate.aggregateKind === 'sum') {
      return directNumberResult(sum)
    }
    if (directAggregate.aggregateKind === 'count') {
      return directNumberResult(count)
    }
    if (directAggregate.aggregateKind === 'min') {
      const minimum = prefix.prefixMinimums[endOffset] ?? Number.POSITIVE_INFINITY
      return directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum)
    }
    if (directAggregate.aggregateKind === 'max') {
      const maximum = prefix.prefixMaximums[endOffset] ?? Number.NEGATIVE_INFINITY
      return directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum)
    }
    const denominator = averageCount
    return denominator === 0 ? directNumberResult(0) : directNumberResult(sum / denominator)
  }

  const resolveStructuredReferenceNow = (tableName: string, columnName: string): FormulaNode | undefined => {
    const table = args.state.workbook.getTable(tableName)
    if (!table) {
      return undefined
    }
    const columnIndex = table.columnNames.findIndex((name) => name.trim().toUpperCase() === columnName.trim().toUpperCase())
    if (columnIndex === -1) {
      return undefined
    }
    const start = parseCellAddress(table.startAddress, table.sheetName)
    const end = parseCellAddress(table.endAddress, table.sheetName)
    const startRow = start.row + (table.headerRow ? 1 : 0)
    const endRow = end.row - (table.totalsRow ? 1 : 0)
    if (endRow < startRow) {
      return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
    }
    const column = start.col + columnIndex
    return {
      kind: 'RangeRef',
      refKind: 'cells',
      sheetName: table.sheetName,
      start: formatAddress(startRow, column),
      end: formatAddress(endRow, column),
    }
  }

  const resolveSpillReferenceNow = (currentSheetName: string, sheetName: string | undefined, address: string): FormulaNode | undefined => {
    const targetSheetName = sheetName ?? currentSheetName
    const spill = args.state.workbook.getSpill(targetSheetName, address)
    if (!spill) {
      return undefined
    }
    const owner = parseCellAddress(address, targetSheetName)
    return {
      kind: 'RangeRef',
      refKind: 'cells',
      sheetName: targetSheetName,
      start: owner.text,
      end: formatAddress(owner.row + spill.rows - 1, owner.col + spill.cols - 1),
    }
  }

  const evaluateCellWithReferenceReplacements = (
    sheetName: string,
    address: string,
    replacements: ReadonlyMap<string, { sheetName: string; address: string }>,
    visiting: Set<string>,
  ): CellValue => {
    const replacementKey = referenceReplacementKey(sheetName, address)
    const replacement = replacements.get(replacementKey)
    if (replacement) {
      return evaluateCellWithReferenceReplacements(replacement.sheetName, replacement.address, replacements, visiting)
    }

    const visitKey = referenceReplacementKey(sheetName, address)
    if (visiting.has(visitKey)) {
      return errorValue(ErrorCode.Cycle)
    }

    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return emptyValue()
    }

    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      const parsedCell = parseCellAddress(address, sheetName)
      return args.runtimeColumnStore.readCellValue(sheetName, parsedCell.row, parsedCell.col)
    }

    visiting.add(visitKey)
    const evaluationContext: EvaluationContext = {
      sheetName,
      currentAddress: address,
      resolveCell: (targetSheetName, targetAddress) =>
        evaluateCellWithReferenceReplacements(targetSheetName, targetAddress, replacements, visiting),
      resolveRange: (targetSheetName, start, end, refKind) => readRangeValues(targetSheetName, start, end, refKind, replacements, visiting),
      resolveName: (name: string) => {
        const definedName = args.state.workbook.getDefinedName(name)
        if (!definedName) {
          return errorValue(ErrorCode.Name)
        }
        return definedNameValueToCellValue(definedName.value, args.state.strings)
      },
      resolveFormula: (targetSheetName: string, targetAddress: string) => {
        const targetCellIndex = args.state.workbook.getCellIndex(targetSheetName, targetAddress)
        return targetCellIndex === undefined ? undefined : args.state.formulas.get(targetCellIndex)?.source
      },
      resolvePivotData: ({
        dataField,
        sheetName: pivotSheetName,
        address: pivotAddress,
        filters,
      }: {
        dataField: string
        sheetName: string
        address: string
        filters: ReadonlyArray<{ field: string; item: CellValue }>
      }) => args.resolvePivotData(pivotSheetName, pivotAddress, dataField, filters),
      resolveMultipleOperations: (nested: {
        formulaSheetName: string
        formulaAddress: string
        rowCellSheetName: string
        rowCellAddress: string
        rowReplacementSheetName: string
        rowReplacementAddress: string
        columnCellSheetName?: string
        columnCellAddress?: string
        columnReplacementSheetName?: string
        columnReplacementAddress?: string
      }) => resolveMultipleOperationsNow(nested),
      listSheetNames: () =>
        [...args.state.workbook.sheetsByName.values()].toSorted((left, right) => left.order - right.order).map((sheet) => sheet.name),
    }
    const jsPlan = formula.compiled.jsPlan.length > 0 ? formula.compiled.jsPlan : lowerToPlan(formula.compiled.optimizedAst)
    const result = evaluatePlanResult(jsPlan, evaluationContext)
    visiting.delete(visitKey)
    return isArrayValue(result) ? (result.values[0] ?? emptyValue()) : result
  }

  const resolveMultipleOperationsNow = (request: {
    formulaSheetName: string
    formulaAddress: string
    rowCellSheetName: string
    rowCellAddress: string
    rowReplacementSheetName: string
    rowReplacementAddress: string
    columnCellSheetName?: string
    columnCellAddress?: string
    columnReplacementSheetName?: string
    columnReplacementAddress?: string
  }): CellValue => {
    const replacements = new Map<string, { sheetName: string; address: string }>()
    replacements.set(referenceReplacementKey(request.rowCellSheetName, request.rowCellAddress), {
      sheetName: request.rowReplacementSheetName,
      address: request.rowReplacementAddress,
    })
    if (
      request.columnCellSheetName &&
      request.columnCellAddress &&
      request.columnReplacementSheetName &&
      request.columnReplacementAddress
    ) {
      replacements.set(referenceReplacementKey(request.columnCellSheetName, request.columnCellAddress), {
        sheetName: request.columnReplacementSheetName,
        address: request.columnReplacementAddress,
      })
    }
    return evaluateCellWithReferenceReplacements(request.formulaSheetName, request.formulaAddress, replacements, new Set<string>())
  }

  const storeFormulaResult = (cellIndex: number, formula: RuntimeFormula, result: EvaluationResult): number[] => {
    const beforeValue = args.state.workbook.cellStore.getValue(cellIndex, (id) => (id === 0 ? '' : args.state.strings.get(id)))
    const materialization = isArrayValue(result)
      ? args.materializeSpill(cellIndex, result)
      : formula.compiled.producesSpill
        ? {
            changedCellIndices: args.clearOwnedSpill(cellIndex),
            ownerValue: result,
          }
        : {
            changedCellIndices: emptyChangedCellIndices,
            ownerValue: result,
          }

    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    args.state.workbook.cellStore.setValue(
      cellIndex,
      materialization.ownerValue,
      materialization.ownerValue.tag === ValueTag.String ? args.state.strings.intern(materialization.ownerValue.value) : 0,
    )
    if (!cellValuesEqual(beforeValue, materialization.ownerValue)) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
    for (let index = 0; index < materialization.changedCellIndices.length; index += 1) {
      args.state.workbook.notifyCellValueWritten(materialization.changedCellIndices[index]!)
    }
    return materialization.changedCellIndices
  }

  const storeDirectScalarResult = (cellIndex: number, result: CellValue): number[] => {
    const beforeValue = args.state.workbook.cellStore.getValue(cellIndex, (id) => (id === 0 ? '' : args.state.strings.get(id)))
    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    args.state.workbook.cellStore.setValue(cellIndex, result, result.tag === ValueTag.String ? args.state.strings.intern(result.value) : 0)
    if (!cellValuesEqual(beforeValue, result)) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
    return emptyChangedCellIndices
  }

  const readDirectScalarOperand = (operand: RuntimeDirectScalarOperand): number | undefined => {
    if (operand.kind === 'literal-number') {
      return operand.value
    }
    const value = readCellValueByIndex(operand.cellIndex)
    switch (value.tag) {
      case ValueTag.Number:
        return value.value
      case ValueTag.Boolean:
        return value.value ? 1 : 0
      case ValueTag.Empty:
        return 0
      case ValueTag.String:
      case ValueTag.Error:
        return undefined
      default:
        return undefined
    }
  }

  const tryEvaluateDirectScalar = (formula: RuntimeFormula): CellValue | undefined => {
    const directScalar = formula.directScalar
    if (!directScalar) {
      return undefined
    }
    if (directScalar.kind === 'abs') {
      const operand = readDirectScalarOperand(directScalar.operand)
      return operand === undefined ? undefined : directNumberResult(Math.abs(operand))
    }
    const left = readDirectScalarOperand(directScalar.left)
    const right = readDirectScalarOperand(directScalar.right)
    if (left === undefined || right === undefined) {
      return undefined
    }
    switch (directScalar.operator) {
      case '+':
        return directNumberResult(left + right)
      case '-':
        return directNumberResult(left - right)
      case '*':
        return directNumberResult(left * right)
      case '/':
        return right === 0 ? directErrorResult(ErrorCode.Div0) : directNumberResult(left / right)
    }
  }

  const evaluateDirectLookupFormulaNow = (cellIndex: number): number[] | undefined => {
    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return undefined
    }
    const directResult =
      tryEvaluateDirectVectorLookup(formula) ??
      tryEvaluateDirectScalar(formula) ??
      tryEvaluateDirectAggregate(formula) ??
      tryEvaluateDirectCriteriaAggregate(formula)
    return directResult === undefined
      ? undefined
      : formula.compiled.producesSpill
        ? storeFormulaResult(cellIndex, formula, directResult)
        : storeDirectScalarResult(cellIndex, directResult)
  }

  const evaluateUnsupportedFormulaNow = (cellIndex: number): number[] => {
    const formula = args.state.formulas.get(cellIndex)
    const sheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
    if (!formula || !sheetName) {
      return []
    }

    const directResult =
      tryEvaluateDirectVectorLookup(formula) ??
      tryEvaluateDirectScalar(formula) ??
      tryEvaluateDirectAggregate(formula) ??
      tryEvaluateDirectCriteriaAggregate(formula)
    if (directResult !== undefined) {
      return storeFormulaResult(cellIndex, formula, directResult)
    }

    const evaluationContext: EvaluationContext = {
      sheetName,
      currentAddress: args.state.workbook.getAddress(cellIndex),
      resolveCell: (targetSheetName: string, address: string) => readCellValue(targetSheetName, address),
      resolveRange: (targetSheetName: string, start: string, end: string, refKind: 'cells' | 'rows' | 'cols') =>
        readRangeValues(targetSheetName, start, end, refKind),
      resolveName: (name: string) => {
        const definedName = args.state.workbook.getDefinedName(name)
        if (!definedName) {
          return errorValue(ErrorCode.Name)
        }
        return definedNameValueToCellValue(definedName.value, args.state.strings)
      },
      resolveFormula: (targetSheetName: string, address: string) => {
        const targetCellIndex = args.state.workbook.getCellIndex(targetSheetName, address)
        return targetCellIndex === undefined ? undefined : args.state.formulas.get(targetCellIndex)?.source
      },
      resolvePivotData: ({
        dataField,
        sheetName: pivotSheetName,
        address,
        filters,
      }: {
        dataField: string
        sheetName: string
        address: string
        filters: ReadonlyArray<{ field: string; item: CellValue }>
      }) => args.resolvePivotData(pivotSheetName, address, dataField, filters),
      resolveMultipleOperations: (request: {
        formulaSheetName: string
        formulaAddress: string
        rowCellSheetName: string
        rowCellAddress: string
        rowReplacementSheetName: string
        rowReplacementAddress: string
        columnCellSheetName?: string
        columnCellAddress?: string
        columnReplacementSheetName?: string
        columnReplacementAddress?: string
      }) => resolveMultipleOperationsNow(request),
      listSheetNames: () =>
        [...args.state.workbook.sheetsByName.values()].toSorted((left, right) => left.order - right.order).map((sheet) => sheet.name),
      resolveExactVectorMatch: (request) => {
        if (
          request.startRow === undefined ||
          request.endRow === undefined ||
          request.startCol === undefined ||
          request.endCol === undefined
        ) {
          return args.exactLookup.findVectorMatch(request)
        }
        return resolveExactVectorMatch(formula, request)
      },
      resolveApproximateVectorMatch: (request) => {
        if (
          request.startRow === undefined ||
          request.endRow === undefined ||
          request.startCol === undefined ||
          request.endCol === undefined
        ) {
          return args.sortedLookup.findVectorMatch(request)
        }
        return resolveApproximateVectorMatch(formula, request)
      },
      resolveLookupBuiltin: lookupBuiltinResolver,
    }
    const result = evaluatePlanResult(formula.compiled.jsPlan, evaluationContext)
    return storeFormulaResult(cellIndex, formula, result)
  }

  return {
    evaluateDirectLookupFormulaNow: evaluateDirectLookupFormulaNow,
    evaluateDirectLookupFormula(cellIndex) {
      return Effect.try({
        try: () => {
          return evaluateDirectLookupFormulaNow(cellIndex)
        },
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage(`Failed to evaluate direct lookup formula ${cellIndex}`, cause),
            cause,
          }),
      })
    },
    evaluateUnsupportedFormula(cellIndex) {
      return Effect.try({
        try: () => evaluateUnsupportedFormulaNow(cellIndex),
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage(`Failed to evaluate formula ${cellIndex}`, cause),
            cause,
          }),
      })
    },
    resolveStructuredReference(tableName, columnName) {
      return Effect.try({
        try: () => resolveStructuredReferenceNow(tableName, columnName),
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage(`Failed to resolve structured reference ${tableName}[${columnName}]`, cause),
            cause,
          }),
      })
    },
    resolveSpillReference(currentSheetName, sheetName, address) {
      return Effect.try({
        try: () => resolveSpillReferenceNow(currentSheetName, sheetName, address),
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage(`Failed to resolve spill reference ${address}#`, cause),
            cause,
          }),
      })
    },
    resolveMultipleOperations(request) {
      return Effect.try({
        try: () => resolveMultipleOperationsNow(request),
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage('Failed to resolve MULTIPLE.OPERATIONS', cause),
            cause,
          }),
      })
    },
  }
}
