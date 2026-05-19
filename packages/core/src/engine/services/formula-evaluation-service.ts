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
import type { CriterionRangeCacheService, CriterionRangeMatch } from './criterion-range-cache-service.js'
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
  directCriteriaCacheValueKey,
  directErrorResult,
  directNumberResult,
  evaluationErrorMessage,
  referenceReplacementKey,
} from './formula-evaluation-helpers.js'
import { tryEvaluateDirectAggregate } from './formula-evaluation-direct-aggregate.js'
import { readRuntimeDirectCriteriaOperandValue } from './direct-criteria-operands.js'
import type { EngineFormulaEvaluationService } from './formula-evaluation-service-types.js'
export type { EngineFormulaEvaluationService } from './formula-evaluation-service-types.js'

const DIRECT_CRITERIA_MATCH_CACHE_LIMIT = 16_384
const INDEXED_WHOLE_AXIS_BOUND_LIMIT = 4096

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
  const directCriteriaMatchCache = new Map<string, CriterionRangeMatch>()
  const rememberDirectCriteriaMatch = (key: string, value: CriterionRangeMatch): CriterionRangeMatch => {
    if (directCriteriaMatchCache.size >= DIRECT_CRITERIA_MATCH_CACHE_LIMIT) {
      const firstKey = directCriteriaMatchCache.keys().next().value
      if (firstKey !== undefined) {
        directCriteriaMatchCache.delete(firstKey)
      }
    }
    directCriteriaMatchCache.set(key, value)
    return value
  }
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
      if (rowEnd - rowStart + 1 <= INDEXED_WHOLE_AXIS_BOUND_LIMIT) {
        for (let row = rowStart; row <= rowEnd; row += 1) {
          sheet.logical.forEachVisibleRowCellEntry(row, (_cellIndex, col) => {
            if (col >= 0 && col < MAX_COLS && col > maxResidentCol) {
              maxResidentCol = col
            }
          })
        }
      } else {
        sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
          if (row >= rowStart && row <= rowEnd && col >= 0 && col < MAX_COLS && col > maxResidentCol) {
            maxResidentCol = col
          }
        })
      }
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
      if (colEnd - colStart + 1 <= INDEXED_WHOLE_AXIS_BOUND_LIMIT) {
        maxResidentRow = args.runtimeColumnStore.findMaxResidentRowInColumns({
          sheetName,
          rowStart: 0,
          rowEnd: MAX_ROWS - 1,
          colStart,
          colEnd,
        })
      } else {
        sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
          if (col >= colStart && col <= colEnd && row >= 0 && row < MAX_ROWS && row > maxResidentRow) {
            maxResidentRow = row
          }
        })
      }
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
    const aggregateRange = directCriteria.aggregateRange
    const criteriaVersionKey = resolvedPairs
      .map((pair) => `${directCriteriaRangeVersionKey(args.state, pair.range)}:${directCriteriaCacheValueKey(pair.criteria)}`)
      .join('|')
    const aggregateCacheKey =
      aggregateRange === undefined
        ? undefined
        : [directCriteria.aggregateKind, directCriteriaRangeVersionKey(args.state, aggregateRange), criteriaVersionKey].join('\u0000')
    const cachedAggregate = aggregateCacheKey === undefined ? undefined : directCriteriaAggregateCache.get(aggregateCacheKey)
    if (cachedAggregate) {
      addEngineCounter(args.state.counters, 'directCriteriaAggregateCacheHits')
      return applyDirectCriteriaResultTransforms(readCellValueByIndex, formula, cachedAggregate)
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
    const exactAggregateResult =
      resolvedPairs.length === 1 && (directCriteria.aggregateKind === 'count' || directCriteria.aggregateKind === 'sum')
        ? args.criterionCache.getOrBuildExactAggregate({
            criteriaPair: resolvedPairs[0]!,
            ...(aggregateRange === undefined ? {} : { aggregateRange }),
            aggregateKind: directCriteria.aggregateKind,
          })
        : undefined
    if (exactAggregateResult !== undefined) {
      const cachedResult =
        aggregateCacheKey === undefined
          ? exactAggregateResult
          : rememberDirectCriteriaResult(directCriteriaAggregateCache, aggregateCacheKey, exactAggregateResult)
      return applyDirectCriteriaResultTransforms(readCellValueByIndex, formula, cachedResult)
    }

    const cachedMatches = directCriteriaMatchCache.get(criteriaVersionKey)
    if (cachedMatches !== undefined) {
      addEngineCounter(args.state.counters, 'directCriteriaMatchCacheHits')
    }
    const matches =
      cachedMatches ??
      args.criterionCache.getOrBuildMatchingRows({
        criteriaPairs: resolvedPairs,
      })
    if ('tag' in matches) {
      return matches
    }
    if (cachedMatches === undefined) {
      rememberDirectCriteriaMatch(criteriaVersionKey, matches)
    }

    if (directCriteria.aggregateKind === 'count') {
      return applyDirectCriteriaResultTransforms(readCellValueByIndex, formula, directNumberResult(matches.length))
    }

    if (!aggregateRange) {
      return undefined
    }
    const concreteAggregateCacheKey = aggregateCacheKey
    if (concreteAggregateCacheKey === undefined) {
      return undefined
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
        rememberDirectCriteriaResult(directCriteriaAggregateCache, concreteAggregateCacheKey, result),
      )
    }

    const aggregateView = args.runtimeColumnStore.getColumnView({
      sheetName: aggregateRange.sheetName,
      rowStart: aggregateRange.rowStart,
      rowEnd: aggregateRange.rowEnd,
      col: aggregateRange.col,
    })
    const matchedAggregateError = firstMatchedAggregateError(aggregateView, matches.rows, matches.length)
    if (matchedAggregateError) {
      return applyDirectCriteriaResultTransforms(
        readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(directCriteriaAggregateCache, concreteAggregateCacheKey, matchedAggregateError),
      )
    }

    if (directCriteria.aggregateKind === 'sum') {
      let sum = 0
      for (let index = 0; index < matches.length; index += 1) {
        sum += numericLikeValueInView(aggregateView, matches.rows[index]!) ?? 0
      }
      return applyDirectCriteriaResultTransforms(
        readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(directCriteriaAggregateCache, concreteAggregateCacheKey, directNumberResult(sum)),
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
          concreteAggregateCacheKey,
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
          concreteAggregateCacheKey,
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
        concreteAggregateCacheKey,
        directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum),
      ),
    )
  }

  const firstMatchedAggregateError = (
    view: ReturnType<EngineRuntimeColumnStoreService['getColumnView']>,
    rows: ArrayLike<number>,
    length: number,
  ): CellValue | undefined => {
    for (let index = 0; index < length; index += 1) {
      const row = rows[index]!
      if ((view.readTagAt(row) as ValueTag) === ValueTag.Error) {
        return directErrorResult(view.readErrorAt(row) as ErrorCode)
      }
    }
    return undefined
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
      tryEvaluateDirectAggregate({
        formula,
        workbook: args.state.workbook,
        counters: args.state.counters,
        aggregateCache: args.aggregateCache,
        readCellValueByIndex,
      }) ??
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
      tryEvaluateDirectAggregate({
        formula,
        workbook: args.state.workbook,
        counters: args.state.counters,
        aggregateCache: args.aggregateCache,
        readCellValueByIndex,
      }) ??
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
    evaluateUnsupportedFormulaNow,
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
