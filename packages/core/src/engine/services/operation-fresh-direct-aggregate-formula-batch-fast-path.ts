import { ErrorCode, ValueTag, type EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { formatAddress } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { batchOpOrder, markBatchApplied, type OpOrder } from '../../replica-state.js'
import type { SheetRecord } from '../../workbook-store.js'
import type {
  EngineRuntimeState,
  RuntimeDirectAggregateDescriptor,
  RuntimeDirectCriteriaDescriptor,
  RuntimeFormula,
  U32,
} from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)
const FRESH_DIRECT_AGGREGATE_FORMULA_BATCH_MIN_SIZE = 32
const FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT = 4096

type FastPathState = Pick<
  EngineRuntimeState,
  | 'workbook'
  | 'ranges'
  | 'strings'
  | 'events'
  | 'formulas'
  | 'counters'
  | 'replicaState'
  | 'getLastMetrics'
  | 'setLastMetrics'
  | 'getSyncClientConnection'
  | 'trackReplicaVersions'
>

interface FreshDirectAggregateFormulaEntry {
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: NonNullable<CreateEngineOperationServiceArgs['compileTemplateFormula']> extends (...args: never[]) => infer Result
    ? Result extends { readonly compiled: infer Compiled }
      ? Compiled
      : never
    : never
  readonly templateId: number
  readonly result: DirectScalarCurrentOperand
}

interface ContiguousSingleColumnFormulaBatch {
  readonly rowStart: number
  readonly col: number
}

interface FreshDirectAggregateMatrixBatch {
  readonly sheet: NonNullable<ReturnType<FastPathState['workbook']['getSheetById']>>
  readonly sheetId: number
  readonly rowStart: number
  readonly rowCount: number
  readonly colStart: number
  readonly inputColCount: number
  readonly formulaCol: number
  readonly values: Float64Array
  readonly formulaEntries: readonly FreshDirectAggregateFormulaEntry[]
}

type FreshFormulaCellAttacher = (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void

export interface OperationFreshDirectAggregateFormulaBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly hasRegionFormulaSubscriptionsOverlappingRange?: (
    sheetId: number,
    rowStart: number,
    rowEnd: number,
    colStart: number,
    colEnd: number,
  ) => boolean
  readonly getRegionFormulaSubscriptionCount?: () => number
  readonly bindPreparedFormula: NonNullable<CreateEngineOperationServiceArgs['bindPreparedFormula']>
  readonly compileTemplateFormula: NonNullable<CreateEngineOperationServiceArgs['compileTemplateFormula']>
  readonly materializeDeferredStructuralFormulaSources: () => void
  readonly beginMutationCollection: () => void
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly getChangedInputBuffer: () => U32
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, result: DirectScalarCurrentOperand) => boolean
}

export function createOperationFreshDirectAggregateFormulaBatchFastPath(args: OperationFreshDirectAggregateFormulaBatchFastPathArgs): {
  readonly tryApplyFreshDirectAggregateFormulaBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyFreshDirectAggregateFormulaMatrixBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => boolean
} {
  const tryApplyFreshDirectAggregateFormulaMatrixBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (
      source !== 'local' ||
      firstRef === undefined ||
      refs.length < FRESH_DIRECT_AGGREGATE_FORMULA_BATCH_MIN_SIZE ||
      potentialNewCells !== refs.length ||
      args.state.workbook.metadata.definedNames.size > 0
    ) {
      return false
    }
    if (args.state.trackReplicaVersions && batch === null) {
      return false
    }
    const matrix = collectFreshDirectAggregateMatrixBatch(args, refs, firstRef)
    if (matrix === null || freshMatrixOverlapsFormulaDependencies(args, matrix)) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + 1)
    args.resetMaterializedCellScratch(refs.length)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const ensureRowId = args.state.workbook.createLogicalAxisIdEnsurer(matrix.sheetId, 'row')
    const ensureColId = args.state.workbook.createLogicalAxisIdEnsurer(matrix.sheetId, 'column')
    const totalColCount = matrix.inputColCount + 1
    const columnIds = Array.from({ length: totalColCount }, (_, offset) => ensureColId(matrix.colStart + offset))
    const attachFreshCell = createFreshFormulaCellAttacher(matrix.sheet)
    const firstCellIndex = args.state.workbook.cellStore.allocateDenseRowMajorAtReserved(
      matrix.sheetId,
      matrix.rowStart,
      matrix.rowCount,
      matrix.colStart,
      totalColCount,
    )
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        let valueIndex = 0
        for (let rowOffset = 0; rowOffset < matrix.rowCount; rowOffset += 1) {
          const row = matrix.rowStart + rowOffset
          const rowId = ensureRowId(row)
          for (let colOffset = 0; colOffset < matrix.inputColCount; colOffset += 1) {
            const col = matrix.colStart + colOffset
            const cellIndex = firstCellIndex + rowOffset * totalColCount + colOffset
            attachFreshCell(row, col, cellIndex, rowId, columnIds[colOffset]!)
            writeFreshNumericLiteralToCellStore(args, cellIndex, matrix.values[valueIndex]!)
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            if (requiresChangedSet) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            }
            if (args.state.trackReplicaVersions && batch !== null) {
              args.setCellEntityVersion(matrix.sheet.name, formatAddress(row, col), batchOpOrder(batch, valueIndex))
            }
            valueIndex += 1
          }
        }
        for (let rowOffset = 0; rowOffset < matrix.rowCount; rowOffset += 1) {
          const entry = matrix.formulaEntries[rowOffset]!
          const cellIndex = firstCellIndex + rowOffset * totalColCount + matrix.inputColCount
          attachFreshCell(entry.row, entry.col, cellIndex, ensureRowId(entry.row), columnIds[matrix.inputColCount]!)
          args.bindPreparedFormula(cellIndex, matrix.sheet.name, entry.source, entry.compiled, entry.templateId, {
            assumeFreshFormula: true,
          })
          args.applyDirectFormulaCurrentResult(cellIndex, entry.result)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          }
          if (args.state.trackReplicaVersions && batch !== null) {
            args.setCellEntityVersion(
              matrix.sheet.name,
              formatAddress(entry.row, entry.col),
              batchOpOrder(batch, matrix.values.length + rowOffset),
            )
          }
        }
        const writtenColumns = new Uint32Array(matrix.inputColCount)
        for (let index = 0; index < matrix.inputColCount; index += 1) {
          writtenColumns[index] = matrix.colStart + index
        }
        args.state.workbook.notifyColumnsWritten(matrix.sheetId, writtenColumns)
      })
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
    addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed: requiresChangedSet ? changedInputArray : EMPTY_CHANGED_CELLS,
      changedInputCount,
      explicitChangedCount,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch: args.emitBatch,
    })
    return true
  }

  const tryApplyFreshDirectAggregateFormulaBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (
      source !== 'local' ||
      firstRef === undefined ||
      refs.length < FRESH_DIRECT_AGGREGATE_FORMULA_BATCH_MIN_SIZE ||
      potentialNewCells !== refs.length ||
      args.state.workbook.metadata.definedNames.size > 0
    ) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (!sheet) {
      return false
    }
    const entries = collectFreshDirectAggregateFormulaEntries(args, refs, firstRef, sheet.name)
    if (entries === null) {
      return false
    }
    if (args.state.trackReplicaVersions && batch === null) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + 1)
    args.resetMaterializedCellScratch(refs.length)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const ensureRowId = args.state.workbook.createLogicalAxisIdEnsurer(firstRef.sheetId, 'row')
    const ensureColId = args.state.workbook.createLogicalAxisIdEnsurer(firstRef.sheetId, 'column')
    const colIdsByColumn = new Map<number, string>()
    const colId = (col: number): string => {
      let cached = colIdsByColumn.get(col)
      if (cached === undefined) {
        cached = ensureColId(col)
        colIdsByColumn.set(col, cached)
      }
      return cached
    }
    const contiguousBatch = getContiguousSingleColumnFormulaBatch(entries)
    const firstContiguousCellIndex =
      contiguousBatch === undefined
        ? undefined
        : args.state.workbook.cellStore.allocateDenseSingleColumnReserved(
            firstRef.sheetId,
            contiguousBatch.rowStart,
            entries.length,
            contiguousBatch.col,
          )
    if (firstContiguousCellIndex === undefined) {
      args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + refs.length)
    }
    const attachFreshCell = createFreshFormulaCellAttacher(sheet)
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let index = 0; index < entries.length; index += 1) {
          const entry = entries[index]!
          const cellIndex =
            firstContiguousCellIndex === undefined
              ? args.state.workbook.cellStore.allocateReserved(firstRef.sheetId, entry.row, entry.col)
              : firstContiguousCellIndex + index
          attachFreshCell(entry.row, entry.col, cellIndex, ensureRowId(entry.row), colId(entry.col))
          args.bindPreparedFormula(cellIndex, sheet.name, entry.source, entry.compiled, entry.templateId, {
            assumeFreshFormula: true,
          })
          args.applyDirectFormulaCurrentResult(cellIndex, entry.result)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          }
          if (args.state.trackReplicaVersions && batch !== null) {
            args.setCellEntityVersion(sheet.name, formatAddress(entry.row, entry.col), batchOpOrder(batch, index))
          }
        }
      })
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed: requiresChangedSet ? changedInputArray : EMPTY_CHANGED_CELLS,
      changedInputCount,
      explicitChangedCount,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch: args.emitBatch,
    })
    return true
  }

  return { tryApplyFreshDirectAggregateFormulaBatch, tryApplyFreshDirectAggregateFormulaMatrixBatch }
}

function getContiguousSingleColumnFormulaBatch(
  entries: readonly FreshDirectAggregateFormulaEntry[],
): ContiguousSingleColumnFormulaBatch | undefined {
  const first = entries[0]
  if (first === undefined) {
    return undefined
  }
  for (let index = 1; index < entries.length; index += 1) {
    const entry = entries[index]!
    if (entry.col !== first.col || entry.row !== first.row + index) {
      return undefined
    }
  }
  return { rowStart: first.row, col: first.col }
}

function collectFreshDirectAggregateMatrixBatch(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
  firstRef: EngineCellMutationRef,
): FreshDirectAggregateMatrixBatch | null {
  const firstMutation = firstRef.mutation
  if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
    return null
  }
  const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
  if (!sheet) {
    return null
  }

  let firstFormulaRefIndex = -1
  for (let index = 0; index < refs.length; index += 1) {
    if (refs[index]!.mutation.kind === 'setCellFormula') {
      firstFormulaRefIndex = index
      break
    }
  }
  if (firstFormulaRefIndex <= 0) {
    return null
  }

  const values = new Float64Array(firstFormulaRefIndex)
  let currentRow = firstMutation.row
  let currentWidth = 0
  let inputColCount = 0
  let rowCount = 1
  for (let refIndex = 0; refIndex < firstFormulaRefIndex; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    if (
      ref.sheetId !== firstRef.sheetId ||
      ref.cellIndex !== undefined ||
      mutation.kind !== 'setCellValue' ||
      typeof mutation.value !== 'number' ||
      Object.is(mutation.value, -0)
    ) {
      return null
    }
    if (mutation.row === currentRow) {
      if (mutation.col !== firstMutation.col + currentWidth) {
        return null
      }
      currentWidth += 1
    } else {
      if (mutation.row !== currentRow + 1 || mutation.col !== firstMutation.col || currentWidth === 0) {
        return null
      }
      if (inputColCount === 0) {
        inputColCount = currentWidth
      } else if (currentWidth !== inputColCount) {
        return null
      }
      currentRow = mutation.row
      currentWidth = 1
      rowCount += 1
    }
    if (
      sheet.grid.getPhysical(mutation.row, mutation.col) !== -1 ||
      sheet.logical.getVisibleCell(mutation.row, mutation.col) !== undefined
    ) {
      return null
    }
    values[refIndex] = mutation.value
  }
  if (inputColCount === 0) {
    inputColCount = currentWidth
  } else if (currentWidth !== inputColCount) {
    return null
  }
  if (rowCount < 2 || inputColCount < 2 || rowCount * inputColCount !== firstFormulaRefIndex) {
    return null
  }
  if (refs.length - firstFormulaRefIndex !== rowCount) {
    return null
  }
  for (let col = firstMutation.col; col < firstMutation.col + inputColCount; col += 1) {
    if (args.hasTrackedExactLookupDependents(firstRef.sheetId, col) || args.hasTrackedSortedLookupDependents(firstRef.sheetId, col)) {
      return null
    }
  }

  const formulaCol = firstMutation.col + inputColCount
  const formulaEntries: FreshDirectAggregateFormulaEntry[] = []
  for (let refIndex = firstFormulaRefIndex; refIndex < refs.length; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    const rowOffset = refIndex - firstFormulaRefIndex
    if (
      ref.sheetId !== firstRef.sheetId ||
      ref.cellIndex !== undefined ||
      mutation.kind !== 'setCellFormula' ||
      mutation.row !== firstMutation.row + rowOffset ||
      mutation.col !== formulaCol
    ) {
      return null
    }
    if (
      sheet.grid.getPhysical(mutation.row, mutation.col) !== -1 ||
      sheet.logical.getVisibleCell(mutation.row, mutation.col) !== undefined ||
      args.hasTrackedExactLookupDependents(ref.sheetId, mutation.col) ||
      args.hasTrackedSortedLookupDependents(ref.sheetId, mutation.col) ||
      args.hasTrackedDirectRangeDependents(ref.sheetId, mutation.col)
    ) {
      return null
    }
    let template: ReturnType<OperationFreshDirectAggregateFormulaBatchFastPathArgs['compileTemplateFormula']>
    try {
      template = args.compileTemplateFormula(mutation.formula, mutation.row, mutation.col)
    } catch {
      return null
    }
    const compiled = template.compiled
    if (
      compiled.volatile ||
      compiled.producesSpill ||
      compiled.symbolicNames.length !== 0 ||
      compiled.symbolicTables.length !== 0 ||
      compiled.symbolicSpills.length !== 0
    ) {
      return null
    }
    const aggregate = compiled.directAggregateCandidate
    const range = aggregate === undefined ? undefined : compiled.parsedSymbolicRanges?.[aggregate.symbolicRangeIndex]
    if (
      aggregate === undefined ||
      range === undefined ||
      range.refKind !== 'cells' ||
      (range.sheetName ?? sheet.name) !== sheet.name ||
      range.startRow !== range.endRow ||
      range.startRow !== mutation.row ||
      range.startCol < firstMutation.col ||
      range.endCol >= formulaCol ||
      range.startCol > range.endCol ||
      range.endCol - range.startCol + 1 > FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT
    ) {
      return null
    }
    const result = evaluateFreshDirectAggregateMatrixRow({
      aggregateKind: aggregate.aggregateKind,
      colEnd: range.endCol,
      colStart: range.startCol,
      inputColCount,
      matrixColStart: firstMutation.col,
      resultOffset: aggregate.resultOffset,
      rowOffset,
      values,
    })
    formulaEntries.push({
      row: mutation.row,
      col: mutation.col,
      source: mutation.formula,
      compiled,
      templateId: template.templateId,
      result,
    })
  }

  return {
    sheet,
    sheetId: firstRef.sheetId,
    rowStart: firstMutation.row,
    rowCount,
    colStart: firstMutation.col,
    inputColCount,
    formulaCol,
    values,
    formulaEntries,
  }
}

function createFreshFormulaCellAttacher(sheet: SheetRecord): FreshFormulaCellAttacher {
  const attachFreshVisibleCellIdentity = sheet.logical.setFreshVisibleCellIdentityWithAxisIdsDeferred.bind(sheet.logical)
  sheet.logical.deferVisibleCellPageRebuild()
  const setGridCell = sheet.grid.createRowMajorSetter()
  return (row, col, cellIndex, rowId, colId) => {
    attachFreshVisibleCellIdentity(cellIndex, rowId, colId)
    setGridCell(row, col, cellIndex)
  }
}

function writeFreshNumericLiteralToCellStore(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  cellIndex: number,
  value: number,
): void {
  const cellStore = args.state.workbook.cellStore
  cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  cellStore.stringIds[cellIndex] = 0
  cellStore.tags[cellIndex] = ValueTag.Number
  cellStore.numbers[cellIndex] = value
  cellStore.errors[cellIndex] = ErrorCode.None
}

function evaluateFreshDirectAggregateMatrixRow(input: {
  readonly aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
  readonly colEnd: number
  readonly colStart: number
  readonly inputColCount: number
  readonly matrixColStart: number
  readonly resultOffset: number | undefined
  readonly rowOffset: number
  readonly values: Float64Array
}): DirectScalarCurrentOperand {
  let sum = 0
  let count = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  const rowBase = input.rowOffset * input.inputColCount
  for (let col = input.colStart; col <= input.colEnd; col += 1) {
    const value = input.values[rowBase + col - input.matrixColStart]!
    sum += value
    count += 1
    minimum = Math.min(minimum, value)
    maximum = Math.max(maximum, value)
  }
  const result =
    input.aggregateKind === 'sum'
      ? sum
      : input.aggregateKind === 'count'
        ? count
        : input.aggregateKind === 'average'
          ? count === 0
            ? undefined
            : sum / count
          : input.aggregateKind === 'min'
            ? minimum === Number.POSITIVE_INFINITY
              ? 0
              : minimum
            : maximum === Number.NEGATIVE_INFINITY
              ? 0
              : maximum
  if (result === undefined) {
    return { kind: 'error', code: ErrorCode.Div0 }
  }
  return { kind: 'number', value: result + (input.resultOffset ?? 0) }
}

function freshMatrixOverlapsFormulaDependencies(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  matrix: FreshDirectAggregateMatrixBatch,
): boolean {
  const rowEnd = matrix.rowStart + matrix.rowCount - 1
  const colEnd = matrix.formulaCol
  if (
    args.getRegionFormulaSubscriptionCount?.() === args.state.formulas.size &&
    args.hasRegionFormulaSubscriptionsOverlappingRange?.(matrix.sheetId, matrix.rowStart, rowEnd, matrix.colStart, colEnd) === false
  ) {
    return false
  }
  let overlaps = false
  args.state.formulas.forEach((formula) => {
    if (overlaps) {
      return
    }
    overlaps =
      directAggregateOverlapsFreshMatrix(formula.directAggregate, matrix.sheet.name, matrix.rowStart, rowEnd, matrix.colStart, colEnd) ||
      directCriteriaOverlapsFreshMatrix(formula.directCriteria, matrix.sheet.name, matrix.rowStart, rowEnd, matrix.colStart, colEnd) ||
      rangeDependenciesOverlapFreshMatrix(args, formula, matrix.sheetId, matrix.rowStart, rowEnd, matrix.colStart, colEnd)
  })
  return overlaps
}

function directAggregateOverlapsFreshMatrix(
  aggregate: RuntimeDirectAggregateDescriptor | undefined,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  return (
    aggregate !== undefined &&
    aggregate.sheetName === sheetName &&
    aggregate.rowEnd >= rowStart &&
    aggregate.rowStart <= rowEnd &&
    aggregate.colEnd >= colStart &&
    aggregate.col <= colEnd
  )
}

function directCriteriaOverlapsFreshMatrix(
  criteria: RuntimeDirectCriteriaDescriptor | undefined,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  if (criteria === undefined) {
    return false
  }
  const rangeOverlaps = (range: {
    readonly sheetName: string
    readonly rowStart: number
    readonly rowEnd: number
    readonly col: number
  }): boolean =>
    range.sheetName === sheetName && range.rowEnd >= rowStart && range.rowStart <= rowEnd && range.col >= colStart && range.col <= colEnd
  return (
    (criteria.aggregateRange !== undefined && rangeOverlaps(criteria.aggregateRange)) ||
    criteria.criteriaPairs.some((pair) => rangeOverlaps(pair.range))
  )
}

function rangeDependenciesOverlapFreshMatrix(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  formula: RuntimeFormula,
  sheetId: number,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  for (let index = 0; index < formula.rangeDependencies.length; index += 1) {
    if (rangeDescriptorOverlapsFreshMatrix(args, formula.rangeDependencies[index]!, sheetId, rowStart, rowEnd, colStart, colEnd)) {
      return true
    }
  }
  for (let index = 0; index < formula.graphRangeDependencies.length; index += 1) {
    if (rangeDescriptorOverlapsFreshMatrix(args, formula.graphRangeDependencies[index]!, sheetId, rowStart, rowEnd, colStart, colEnd)) {
      return true
    }
  }
  return false
}

function rangeDescriptorOverlapsFreshMatrix(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  rangeIndex: number,
  sheetId: number,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  const descriptor = args.state.ranges.getDescriptor(rangeIndex)
  return (
    descriptor.sheetId === sheetId &&
    descriptor.row2 >= rowStart &&
    descriptor.row1 <= rowEnd &&
    descriptor.col2 >= colStart &&
    descriptor.col1 <= colEnd
  )
}

function collectFreshDirectAggregateFormulaEntries(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
  firstRef: EngineCellMutationRef,
  ownerSheetName: string,
): readonly FreshDirectAggregateFormulaEntry[] | null {
  const entries: FreshDirectAggregateFormulaEntry[] = []
  const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
  if (!sheet) {
    return null
  }
  for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    if (ref.sheetId !== firstRef.sheetId || ref.cellIndex !== undefined || mutation.kind !== 'setCellFormula') {
      return null
    }
    if (
      sheet.grid.getPhysical(mutation.row, mutation.col) !== -1 ||
      sheet.logical.getVisibleCell(mutation.row, mutation.col) !== undefined ||
      args.hasTrackedExactLookupDependents(ref.sheetId, mutation.col) ||
      args.hasTrackedSortedLookupDependents(ref.sheetId, mutation.col) ||
      args.hasTrackedDirectRangeDependents(ref.sheetId, mutation.col)
    ) {
      return null
    }
    let template: ReturnType<OperationFreshDirectAggregateFormulaBatchFastPathArgs['compileTemplateFormula']>
    try {
      template = args.compileTemplateFormula(mutation.formula, mutation.row, mutation.col)
    } catch {
      return null
    }
    const compiled = template.compiled
    if (
      compiled.volatile ||
      compiled.producesSpill ||
      compiled.symbolicNames.length !== 0 ||
      compiled.symbolicTables.length !== 0 ||
      compiled.symbolicSpills.length !== 0
    ) {
      return null
    }
    const aggregate = compiled.directAggregateCandidate
    const range = aggregate === undefined ? undefined : compiled.parsedSymbolicRanges?.[aggregate.symbolicRangeIndex]
    if (
      aggregate === undefined ||
      range === undefined ||
      range.refKind !== 'cells' ||
      (range.sheetName ?? ownerSheetName) !== ownerSheetName ||
      range.startRow !== range.endRow ||
      range.startRow !== mutation.row ||
      range.startCol > range.endCol ||
      (mutation.col >= range.startCol && mutation.col <= range.endCol) ||
      range.endCol - range.startCol + 1 > FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT
    ) {
      return null
    }
    const result = evaluateFreshDirectAggregateRow(args, {
      sheetId: ref.sheetId,
      row: range.startRow,
      colStart: range.startCol,
      colEnd: range.endCol,
      aggregateKind: aggregate.aggregateKind,
      resultOffset: aggregate.resultOffset,
    })
    if (result === undefined) {
      return null
    }
    entries.push({
      row: mutation.row,
      col: mutation.col,
      source: mutation.formula,
      compiled,
      templateId: template.templateId,
      result,
    })
  }
  return entries
}

function evaluateFreshDirectAggregateRow(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  request: {
    readonly sheetId: number
    readonly row: number
    readonly colStart: number
    readonly colEnd: number
    readonly aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
    readonly resultOffset: number | undefined
  },
): DirectScalarCurrentOperand | undefined {
  const sheet = args.state.workbook.getSheetById(request.sheetId)
  if (!sheet) {
    return undefined
  }
  const cellStore = args.state.workbook.cellStore
  let sum = 0
  let count = 0
  let averageCount = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  for (let col = request.colStart; col <= request.colEnd; col += 1) {
    const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(request.row, col) : sheet.grid.get(request.row, col)
    if (memberCellIndex === -1) {
      continue
    }
    if (args.state.formulas.has(memberCellIndex)) {
      return undefined
    }
    const tag = (cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    switch (tag) {
      case ValueTag.Number: {
        const value = cellStore.numbers[memberCellIndex] ?? 0
        sum += value
        count += 1
        averageCount += 1
        minimum = Math.min(minimum, value)
        maximum = Math.max(maximum, value)
        break
      }
      case ValueTag.Boolean: {
        const value = (cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
        sum += value
        count += 1
        averageCount += 1
        minimum = Math.min(minimum, value)
        maximum = Math.max(maximum, value)
        break
      }
      case ValueTag.Error:
        if (request.aggregateKind === 'sum' || request.aggregateKind === 'average') {
          return { kind: 'error', code: (cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
        }
        break
      case ValueTag.Empty:
      case ValueTag.String:
        break
    }
  }
  const result =
    request.aggregateKind === 'sum'
      ? sum
      : request.aggregateKind === 'count'
        ? count
        : request.aggregateKind === 'average'
          ? averageCount === 0
            ? undefined
            : sum / averageCount
          : request.aggregateKind === 'min'
            ? minimum === Number.POSITIVE_INFINITY
              ? 0
              : minimum
            : maximum === Number.NEGATIVE_INFINITY
              ? 0
              : maximum
  if (result === undefined) {
    return { kind: 'error', code: ErrorCode.Div0 }
  }
  return { kind: 'number', value: result + (request.resultOffset ?? 0) }
}
