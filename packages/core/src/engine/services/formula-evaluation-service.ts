import { Effect } from 'effect'
import { ErrorCode, MAX_COLS, MAX_ROWS, ValueTag, type CellValue } from '@bilig/protocol'
import {
  createLookupBuiltinResolver,
  evaluatePlanResult,
  formatAddress,
  lowerToPlan,
  isArrayValue,
  type EvaluationContext,
  type EvaluationResult,
  type FormulaNode,
  type RangeBuiltinArgument,
  parseCellAddress,
  parseRangeAddress,
} from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import { definedNameValueToCellValue, definedNameValueToReferenceOperand } from '../../engine-metadata-utils.js'
import { emptyValue, errorValue } from '../../engine-value-utils.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeDirectCriteriaOperand, RuntimeFormula, SpillMaterialization } from '../runtime-state.js'
import { EngineFormulaEvaluationError } from '../errors.js'
import type { CriterionRangeCacheService } from './criterion-range-cache-service.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { EngineRuntimeColumnStoreService } from './runtime-column-store-service.js'
import type { RangeAggregateCacheService } from './range-aggregate-cache-service.js'
import type { SortedColumnSearchService } from './sorted-column-search-service.js'
import { directCriteriaRangeVersionKey, rememberDirectCriteriaResult } from './formula-evaluation-direct-criteria-cache.js'
import {
  applyDirectCriteriaResultTransforms,
  numericLikeValueInView,
  strictNumericAggregateCandidateInView,
  tryEvaluateDirectCriteriaTransformShortCircuit,
} from './formula-evaluation-direct-criteria-transforms.js'
import { tryEvaluateDirectVectorLookup } from './formula-evaluation-direct-lookup.js'
import { tryEvaluateDirectIndexExactMatch, tryEvaluateDirectIndexOffset } from './formula-evaluation-direct-index.js'
import { tryEvaluateDirectScalar } from './formula-evaluation-direct-scalar.js'
import {
  cellValueCriteriaString,
  cellValuesEqual,
  decodeErrorCode,
  directCriteriaCacheValueKey,
  directErrorResult,
  directNumberResult,
  evaluationErrorMessage,
  offsetDirectAggregateResult,
  referenceReplacementKey,
} from './formula-evaluation-helpers.js'
import { readRuntimeDirectCriteriaOperandValue } from './direct-criteria-operands.js'
import type { EngineFormulaEvaluationService } from './formula-evaluation-service-types.js'
export type { EngineFormulaEvaluationService } from './formula-evaluation-service-types.js'
const DIRECT_AGGREGATE_SCAN_MAX_LENGTH = 64
const DIRECT_AGGREGATE_PREFIX_MIN_LENGTH = 16
export function createEngineFormulaEvaluationService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings' | 'formulas' | 'counters' | 'getUseColumnIndex'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly criterionCache: CriterionRangeCacheService
  readonly aggregateCache: RangeAggregateCacheService
  readonly exactLookup: Pick<ExactColumnIndexService, 'findVectorMatch' | 'prepareVectorLookup' | 'findPreparedVectorMatch'>
  readonly sortedLookup: Pick<SortedColumnSearchService, 'findVectorMatch' | 'prepareVectorLookup' | 'findPreparedVectorMatch'>
  readonly materializeSpill: (cellIndex: number, arrayValue: { values: CellValue[]; rows: number; cols: number }) => SpillMaterialization
  readonly clearOwnedSpill: (cellIndex: number) => number[]
  readonly checkEvaluationBudget: (stepCost?: number) => void
  readonly resolvePivotData: (
    sheetName: string,
    address: string,
    dataField: string,
    filters: ReadonlyArray<{ field: string; item: CellValue }>,
  ) => CellValue
}): EngineFormulaEvaluationService {
  const emptyChangedCellIndices: number[] = []
  const directCriteriaAggregateCache = new Map<string, CellValue>()
  const readCellValue = (sheetName: string, address: string): CellValue => {
    if (!args.state.workbook.getSheet(sheetName)) {
      return errorValue(ErrorCode.Ref)
    }
    const parsed = parseCellAddress(address, sheetName)
    return readCellValueAt(sheetName, parsed.row, parsed.col)
  }
  const readCellValueByIndex = (cellIndex: number | undefined): CellValue => {
    if (cellIndex === undefined) {
      return emptyValue()
    }
    return args.state.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : args.state.strings.get(stringId)))
  }
  const readCellValueAt = (sheetName: string, row: number, col: number): CellValue => {
    const sheet = args.state.workbook.getSheet(sheetName)
    return sheet ? readCellValueByIndex(sheet.logical.getVisibleCell(row, col)) : errorValue(ErrorCode.Ref)
  }
  const workbookDateSystem = () => args.state.workbook.getCalculationSettings().dateSystem ?? '1900'

  const directVectorLookupContext = {
    state: args.state,
    exactLookup: args.exactLookup,
    sortedLookup: args.sortedLookup,
    readCellValueByIndex,
  }
  const readRectangularRangeValues = (
    sheetName: string,
    bounds: {
      rowStart: number
      rowEnd: number
      colStart: number
      colEnd: number
    },
    replacements?: ReadonlyMap<string, { sheetName: string; address: string }>,
    visiting?: Set<string>,
  ): CellValue[] => {
    if (bounds.rowEnd < bounds.rowStart || bounds.colEnd < bounds.colStart) {
      return []
    }
    const cellCount = (bounds.rowEnd - bounds.rowStart + 1) * (bounds.colEnd - bounds.colStart + 1)
    args.checkEvaluationBudget(cellCount)
    if (!replacements || !visiting) {
      if (cellCount <= 64) {
        const values: CellValue[] = []
        for (let row = bounds.rowStart; row <= bounds.rowEnd; row += 1) {
          for (let col = bounds.colStart; col <= bounds.colEnd; col += 1) {
            values.push(readCellValueAt(sheetName, row, col))
          }
        }
        return values
      }
      const rangeValues = args.runtimeColumnStore.readRangeValues({
        sheetName,
        rowStart: bounds.rowStart,
        rowEnd: bounds.rowEnd,
        colStart: bounds.colStart,
        colEnd: bounds.colEnd,
      })
      args.checkEvaluationBudget(rangeValues.length)
      return rangeValues
    }

    const values: CellValue[] = []
    for (let row = bounds.rowStart; row <= bounds.rowEnd; row += 1) {
      for (let col = bounds.colStart; col <= bounds.colEnd; col += 1) {
        args.checkEvaluationBudget()
        values.push(evaluateCellWithReferenceReplacements(sheetName, formatAddress(row, col), replacements, visiting))
      }
    }
    return values
  }

  const readRangeValues = (
    sheetName: string,
    start: string,
    end: string,
    refKind: 'cells' | 'rows' | 'cols',
    replacements?: ReadonlyMap<string, { sheetName: string; address: string }>,
    visiting?: Set<string>,
  ): CellValue[] => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return [errorValue(ErrorCode.Ref)]
    }
    const range = parseRangeAddress(`${start}:${end}`, sheetName)
    if (range.kind === 'cells' && refKind === 'cells') {
      return readRectangularRangeValues(
        sheetName,
        {
          rowStart: range.start.row,
          rowEnd: range.end.row,
          colStart: range.start.col,
          colEnd: range.end.col,
        },
        replacements,
        visiting,
      )
    }
    if (range.kind === 'rows' && refKind === 'rows') {
      const rowStart = Math.max(0, range.start.row)
      const rowEnd = Math.min(MAX_ROWS - 1, range.end.row)
      if (rowEnd < rowStart) {
        return []
      }
      let maxResidentCol = -1
      sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
        if (row >= rowStart && row <= rowEnd && col >= 0 && col < MAX_COLS && col > maxResidentCol) {
          maxResidentCol = col
        }
      })
      return maxResidentCol < 0
        ? []
        : readRectangularRangeValues(
            sheetName,
            {
              rowStart,
              rowEnd,
              colStart: 0,
              colEnd: maxResidentCol,
            },
            replacements,
            visiting,
          )
    }
    if (range.kind === 'cols' && refKind === 'cols') {
      const colStart = Math.max(0, range.start.col)
      const colEnd = Math.min(MAX_COLS - 1, range.end.col)
      if (colEnd < colStart) {
        return []
      }
      let maxResidentRow = -1
      sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
        if (col >= colStart && col <= colEnd && row >= 0 && row < MAX_ROWS && row > maxResidentRow) {
          maxResidentRow = row
        }
      })
      return maxResidentRow < 0
        ? []
        : readRectangularRangeValues(
            sheetName,
            {
              rowStart: 0,
              rowEnd: maxResidentRow,
              colStart,
              colEnd,
            },
            replacements,
            visiting,
          )
    }
    return []
  }
  const createRowHiddenResolver = (): ((sheetName: string, rowIndex: number) => boolean) => {
    const hiddenRowsBySheet = new Map<string, Set<number>>()
    return (sheetName, rowIndex) => {
      if (!Number.isInteger(rowIndex) || rowIndex < 0) {
        return false
      }
      let hiddenRows = hiddenRowsBySheet.get(sheetName)
      if (hiddenRows === undefined) {
        hiddenRows = new Set<number>()
        for (const entry of args.state.workbook.listRowAxisEntries(sheetName)) {
          if (entry.hidden === true) {
            hiddenRows.add(entry.index)
          }
        }
        for (const record of args.state.workbook.listRowMetadata(sheetName)) {
          if (record.hidden !== true) {
            continue
          }
          for (let row = record.start; row < record.start + record.count; row += 1) {
            hiddenRows.add(row)
          }
        }
        hiddenRowsBySheet.set(sheetName, hiddenRows)
      }
      return hiddenRows.has(rowIndex)
    }
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
  const readDirectCriteriaOperandValue = (operand: RuntimeDirectCriteriaOperand): CellValue => {
    return readRuntimeDirectCriteriaOperandValue({
      operand,
      readCellValueByIndex,
      stringifyCriteriaValue: cellValueCriteriaString,
    })
  }
  const tryEvaluateDirectCriteriaAggregate = (formula: RuntimeFormula): CellValue | undefined => {
    const directCriteria = formula.directCriteria
    if (!directCriteria) return undefined
    const transformShortCircuit = tryEvaluateDirectCriteriaTransformShortCircuit(readCellValueByIndex, formula)
    if (transformShortCircuit) {
      return transformShortCircuit
    }
    const directIndexOffsetResult = tryEvaluateDirectIndexOffset({
      directCriteria,
      runtimeColumnStore: args.runtimeColumnStore,
      readCellValueByIndex,
    })
    if (directIndexOffsetResult !== undefined) {
      return applyDirectCriteriaResultTransforms(readCellValueByIndex, formula, directIndexOffsetResult)
    }
    const resolvedPairs = directCriteria.criteriaPairs.map((pair) => ({
      range: pair.range,
      criteria: readDirectCriteriaOperandValue(pair.criterion),
    }))
    const criterionError = resolvedPairs.find((pair) => pair.criteria.tag === ValueTag.Error)?.criteria
    if (criterionError) {
      return criterionError
    }
    const directIndexExactMatchResult =
      resolvedPairs.length === 1
        ? tryEvaluateDirectIndexExactMatch({
            directCriteria,
            exactLookup: args.exactLookup,
            runtimeColumnStore: args.runtimeColumnStore,
            lookupValue: resolvedPairs[0]!.criteria,
          })
        : undefined
    if (directIndexExactMatchResult !== undefined) {
      return applyDirectCriteriaResultTransforms(readCellValueByIndex, formula, directIndexExactMatchResult)
    }

    const matches = args.criterionCache.getOrBuildMatchingRows({
      criteriaPairs: resolvedPairs,
    })
    if ('tag' in matches) {
      return matches
    }

    if (directCriteria.aggregateKind === 'count') {
      return applyDirectCriteriaResultTransforms(readCellValueByIndex, formula, directNumberResult(matches.length))
    }

    const aggregateRange = directCriteria.aggregateRange
    if (!aggregateRange) {
      return undefined
    }
    const aggregateCacheKey = [
      directCriteria.aggregateKind,
      directCriteriaRangeVersionKey(args.state, aggregateRange),
      resolvedPairs
        .map((pair) => `${directCriteriaRangeVersionKey(args.state, pair.range)}:${directCriteriaCacheValueKey(pair.criteria)}`)
        .join('|'),
    ].join('\u0000')
    const cachedAggregate = directCriteriaAggregateCache.get(aggregateCacheKey)
    if (cachedAggregate) {
      return applyDirectCriteriaResultTransforms(readCellValueByIndex, formula, cachedAggregate)
    }

    if (directCriteria.aggregateKind === 'first') {
      const firstOffset = matches.rows[0]
      const result =
        firstOffset === undefined
          ? directErrorResult(ErrorCode.NA)
          : args.runtimeColumnStore
              .getColumnView({
                sheetName: aggregateRange.sheetName,
                rowStart: aggregateRange.rowStart,
                rowEnd: aggregateRange.rowEnd,
                col: aggregateRange.col,
              })
              .readCellValueAt(firstOffset)
      return applyDirectCriteriaResultTransforms(
        readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(directCriteriaAggregateCache, aggregateCacheKey, result),
      )
    }

    const aggregateView = args.runtimeColumnStore.getColumnView({
      sheetName: aggregateRange.sheetName,
      rowStart: aggregateRange.rowStart,
      rowEnd: aggregateRange.rowEnd,
      col: aggregateRange.col,
    })

    if (directCriteria.aggregateKind === 'sum') {
      let sum = 0
      for (let index = 0; index < matches.length; index += 1) {
        sum += numericLikeValueInView(aggregateView, matches.rows[index]!) ?? 0
      }
      return applyDirectCriteriaResultTransforms(
        readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(directCriteriaAggregateCache, aggregateCacheKey, directNumberResult(sum)),
      )
    }

    if (directCriteria.aggregateKind === 'average') {
      let count = 0
      let sum = 0
      for (let index = 0; index < matches.length; index += 1) {
        const numeric = numericLikeValueInView(aggregateView, matches.rows[index]!)
        if (numeric === undefined) {
          continue
        }
        count += 1
        sum += numeric
      }
      return applyDirectCriteriaResultTransforms(
        readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(
          directCriteriaAggregateCache,
          aggregateCacheKey,
          count === 0 ? directErrorResult(ErrorCode.Div0) : directNumberResult(sum / count),
        ),
      )
    }

    if (directCriteria.aggregateKind === 'min') {
      let minimum = Number.POSITIVE_INFINITY
      for (let index = 0; index < matches.length; index += 1) {
        const numeric = strictNumericAggregateCandidateInView(aggregateView, matches.rows[index]!)
        if (numeric === undefined) {
          continue
        }
        minimum = Math.min(minimum, numeric)
      }
      return applyDirectCriteriaResultTransforms(
        readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(
          directCriteriaAggregateCache,
          aggregateCacheKey,
          directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum),
        ),
      )
    }

    let maximum = Number.NEGATIVE_INFINITY
    for (let index = 0; index < matches.length; index += 1) {
      const numeric = strictNumericAggregateCandidateInView(aggregateView, matches.rows[index]!)
      if (numeric === undefined) {
        continue
      }
      maximum = Math.max(maximum, numeric)
    }
    return applyDirectCriteriaResultTransforms(
      readCellValueByIndex,
      formula,
      rememberDirectCriteriaResult(
        directCriteriaAggregateCache,
        aggregateCacheKey,
        directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum),
      ),
    )
  }

  const tryEvaluateDirectAggregate = (formula: RuntimeFormula): CellValue | undefined => {
    const directAggregate = formula.directAggregate
    if (!directAggregate) {
      return undefined
    }
    const columnCount = directAggregate.colEnd - directAggregate.col + 1
    const canUseSlidingPrefix =
      formula.dependencyIndices.length === 0 &&
      (directAggregate.aggregateKind === 'sum' ||
        directAggregate.aggregateKind === 'average' ||
        directAggregate.aggregateKind === 'count') &&
      directAggregate.length > DIRECT_AGGREGATE_PREFIX_MIN_LENGTH
    const canUseExistingLargePrefix = formula.dependencyIndices.length === 0 && directAggregate.length > DIRECT_AGGREGATE_SCAN_MAX_LENGTH
    if (!canUseSlidingPrefix && !canUseExistingLargePrefix) {
      addEngineCounter(args.state.counters, 'directAggregateScanEvaluations')
      addEngineCounter(args.state.counters, 'directAggregateScanCells', directAggregate.length)
      const aggregateSheet = args.state.workbook.getSheet(directAggregate.sheetName)
      if (!aggregateSheet) {
        return undefined
      }
      let sum = 0
      let count = 0
      let averageCount = 0
      let minimum = Number.POSITIVE_INFINITY
      let maximum = Number.NEGATIVE_INFINITY
      for (let col = directAggregate.col; col <= directAggregate.colEnd; col += 1) {
        for (let row = directAggregate.rowStart; row <= directAggregate.rowEnd; row += 1) {
          const memberCellIndex =
            aggregateSheet.structureVersion === 1 ? aggregateSheet.grid.getPhysical(row, col) : aggregateSheet.grid.get(row, col)
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
      }
      if (directAggregate.aggregateKind === 'sum') {
        return offsetDirectAggregateResult(directAggregate, directNumberResult(sum))
      }
      if (directAggregate.aggregateKind === 'count') {
        return offsetDirectAggregateResult(directAggregate, directNumberResult(count))
      }
      if (directAggregate.aggregateKind === 'average') {
        return averageCount === 0
          ? directErrorResult(ErrorCode.Div0)
          : offsetDirectAggregateResult(directAggregate, directNumberResult(sum / averageCount))
      }
      if (directAggregate.aggregateKind === 'min') {
        return offsetDirectAggregateResult(directAggregate, directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum))
      }
      return offsetDirectAggregateResult(directAggregate, directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum))
    }
    addEngineCounter(args.state.counters, 'directAggregatePrefixEvaluations')
    // SUM/AVERAGE ranges should reuse any compatible lower-start prefix to
    // avoid rescanning shifted windows, while still allowing narrower anchors
    // when no compatible reusable prefix exists.
    const sharedPrefixStart =
      directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average' || directAggregate.aggregateKind === 'count'
        ? 0
        : directAggregate.rowStart
    let errorCode = ErrorCode.None
    let sum = 0
    let count = 0
    let averageCount = 0
    let errorCount = 0
    let minimum = Number.POSITIVE_INFINITY
    let maximum = Number.NEGATIVE_INFINITY
    let hasShiftedPrefixStart = false
    for (let colOffset = 0; colOffset < columnCount; colOffset += 1) {
      const prefix = args.aggregateCache.getOrBuildColumnPrefix(
        {
          sheetName: directAggregate.sheetName,
          rowStart: sharedPrefixStart,
          rowEnd: directAggregate.rowEnd,
          col: directAggregate.col + colOffset,
        },
        directAggregate.aggregateKind,
      )
      const endOffset = directAggregate.rowEnd - prefix.rowStart
      const startOffset = directAggregate.rowStart - prefix.rowStart - 1
      hasShiftedPrefixStart ||= startOffset >= 0
      const prefixSum = prefix.prefixSums[endOffset] ?? 0
      const prefixCount = prefix.prefixCount[endOffset] ?? 0
      const prefixAverageCount = prefix.prefixAverageCount[endOffset] ?? 0
      const prefixErrorCount = prefix.prefixErrorCounts[endOffset] ?? 0
      sum += startOffset >= 0 ? prefixSum - (prefix.prefixSums[startOffset] ?? 0) : prefixSum
      count += startOffset >= 0 ? prefixCount - (prefix.prefixCount[startOffset] ?? 0) : prefixCount
      averageCount += startOffset >= 0 ? prefixAverageCount - (prefix.prefixAverageCount[startOffset] ?? 0) : prefixAverageCount
      errorCount += startOffset >= 0 ? prefixErrorCount - (prefix.prefixErrorCounts[startOffset] ?? 0) : prefixErrorCount
      const nextErrorCode = prefix.prefixErrorCodes[endOffset]
      if (errorCode === ErrorCode.None && nextErrorCode !== undefined && nextErrorCode !== Number(ErrorCode.None)) {
        errorCode = decodeErrorCode(nextErrorCode)
      }
      minimum = Math.min(minimum, prefix.prefixMinimums[endOffset] ?? Number.POSITIVE_INFINITY)
      maximum = Math.max(maximum, prefix.prefixMaximums[endOffset] ?? Number.NEGATIVE_INFINITY)
    }
    if (
      errorCode !== ErrorCode.None &&
      errorCount > 0 &&
      (directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average')
    ) {
      return hasShiftedPrefixStart ? undefined : directErrorResult(decodeErrorCode(errorCode))
    }
    if (directAggregate.aggregateKind === 'sum') {
      return offsetDirectAggregateResult(directAggregate, directNumberResult(sum))
    }
    if (directAggregate.aggregateKind === 'count') {
      return offsetDirectAggregateResult(directAggregate, directNumberResult(count))
    }
    if (directAggregate.aggregateKind === 'min') {
      return offsetDirectAggregateResult(directAggregate, directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum))
    }
    if (directAggregate.aggregateKind === 'max') {
      return offsetDirectAggregateResult(directAggregate, directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum))
    }
    const denominator = averageCount
    return denominator === 0
      ? directErrorResult(ErrorCode.Div0)
      : offsetDirectAggregateResult(directAggregate, directNumberResult(sum / denominator))
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

    if (!args.state.workbook.getSheet(sheetName)) {
      return errorValue(ErrorCode.Ref)
    }

    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return emptyValue()
    }

    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return readCellValueByIndex(cellIndex)
    }

    visiting.add(visitKey)
    const isRowHidden = createRowHiddenResolver()
    const evaluationContext: EvaluationContext = {
      sheetName,
      workbookName: args.state.workbook.workbookName,
      currentAddress: address,
      dateSystem: workbookDateSystem(),
      resolveCell: (targetSheetName, targetAddress) =>
        evaluateCellWithReferenceReplacements(targetSheetName, targetAddress, replacements, visiting),
      resolveRange: (targetSheetName, start, end, refKind) => readRangeValues(targetSheetName, start, end, refKind, replacements, visiting),
      resolveName: (name: string, scopeSheetName?: string) => {
        const definedName = args.state.workbook.getDefinedName(name, scopeSheetName ?? sheetName)
        if (!definedName) {
          return errorValue(ErrorCode.Name)
        }
        return definedNameValueToCellValue(definedName.value, args.state.strings)
      },
      resolveNameReference: (name: string, scopeSheetName?: string) => {
        const definedName = args.state.workbook.getDefinedName(name, scopeSheetName ?? sheetName)
        return definedName ? definedNameValueToReferenceOperand(definedName.value) : undefined
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
      isRowHidden,
      checkEvaluationBudget: (stepCost) => args.checkEvaluationBudget(stepCost),
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
    const cellStore = args.state.workbook.cellStore
    const beforeTag = cellStore.tags[cellIndex]
    const beforeNumber = cellStore.numbers[cellIndex] ?? 0
    const beforeStringId = cellStore.stringIds[cellIndex] ?? 0
    const beforeError = cellStore.errors[cellIndex] ?? ErrorCode.None
    const nextStringId = result.tag === ValueTag.String ? args.state.strings.intern(result.value) : 0
    const changed =
      beforeTag !== result.tag ||
      (result.tag === ValueTag.Number && !Object.is(beforeNumber, result.value)) ||
      (result.tag === ValueTag.Boolean && beforeNumber !== (result.value ? 1 : 0)) ||
      (result.tag === ValueTag.String && beforeStringId !== nextStringId) ||
      (result.tag === ValueTag.Error && (beforeError as ErrorCode) !== result.code)
    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    args.state.workbook.cellStore.setValue(cellIndex, result, nextStringId)
    if (changed) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
    return emptyChangedCellIndices
  }

  const evaluateDirectLookupFormulaNow = (cellIndex: number): number[] | undefined => {
    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return undefined
    }
    const directResult =
      tryEvaluateDirectVectorLookup(directVectorLookupContext, formula) ??
      tryEvaluateDirectScalar(formula, readCellValueByIndex) ??
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
      tryEvaluateDirectVectorLookup(directVectorLookupContext, formula) ??
      tryEvaluateDirectScalar(formula, readCellValueByIndex) ??
      tryEvaluateDirectAggregate(formula) ??
      tryEvaluateDirectCriteriaAggregate(formula)
    if (directResult !== undefined) {
      return storeFormulaResult(cellIndex, formula, directResult)
    }

    const isRowHidden = createRowHiddenResolver()
    const evaluationContext: EvaluationContext = {
      sheetName,
      workbookName: args.state.workbook.workbookName,
      currentAddress: args.state.workbook.getAddress(cellIndex),
      dateSystem: workbookDateSystem(),
      resolveCell: (targetSheetName: string, address: string) => readCellValue(targetSheetName, address),
      resolveRange: (targetSheetName: string, start: string, end: string, refKind: 'cells' | 'rows' | 'cols') =>
        readRangeValues(targetSheetName, start, end, refKind),
      resolveName: (name: string, scopeSheetName?: string) => {
        const definedName = args.state.workbook.getDefinedName(name, scopeSheetName ?? sheetName)
        if (!definedName) {
          return errorValue(ErrorCode.Name)
        }
        return definedNameValueToCellValue(definedName.value, args.state.strings)
      },
      resolveNameReference: (name: string, scopeSheetName?: string) => {
        const definedName = args.state.workbook.getDefinedName(name, scopeSheetName ?? sheetName)
        return definedName ? definedNameValueToReferenceOperand(definedName.value) : undefined
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
      isRowHidden,
      checkEvaluationBudget: (stepCost) => args.checkEvaluationBudget(stepCost),
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
    const result = evaluatePlanResult(
      formula.compiled.jsPlan.length > 0 ? formula.compiled.jsPlan : lowerToPlan(formula.compiled.ast),
      evaluationContext,
    )
    return storeFormulaResult(cellIndex, formula, result)
  }

  return {
    evaluateDirectLookupFormulaNow: evaluateDirectLookupFormulaNow,
    evaluateDirectLookupFormula(cellIndex) {
      return Effect.try({
        try: () => evaluateDirectLookupFormulaNow(cellIndex),
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
