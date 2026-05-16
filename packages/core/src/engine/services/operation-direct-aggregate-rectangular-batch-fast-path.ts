import { ErrorCode, ValueTag, type EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markBatchApplied } from '../../replica-state.js'
import type { EngineRuntimeState, RuntimeDirectAggregateDescriptor, RuntimeDirectCriteriaDescriptor, U32 } from '../runtime-state.js'
import { tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

type FastPathState = Pick<
  EngineRuntimeState,
  'workbook' | 'events' | 'formulas' | 'counters' | 'replicaState' | 'getLastMetrics' | 'setLastMetrics' | 'getSyncClientConnection'
>

export interface OperationDirectAggregateRectangularBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly canUseDirectFormulaPostRecalc: (cellIndex: number) => boolean
  readonly canSkipFormulaColumnVersion: (cellIndex: number) => boolean
  readonly writeNumericLiteralToCellStore: (cellIndex: number, value: number) => void
  readonly applyTerminalDirectFormulaNumericResult: (cellIndex: number, value: number) => void
  readonly getSingleEntityDependent: (entityId: number) => number
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
  readonly composeDisjointEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
}

type DenseAggregateBatchMutation = Extract<EngineCellMutationRef['mutation'], { kind: 'setCellValue' | 'clearCell' }>

export function createOperationDirectAggregateRectangularBatchFastPath(args: OperationDirectAggregateRectangularBatchFastPathArgs): {
  readonly tryApplyDenseRectangularDirectAggregateLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
} {
  const tryApplyDenseRectangularDirectAggregateLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32 || (potentialNewCells ?? 0) !== 0) {
      return false
    }
    const rectangle = collectDenseNumericRectangle(args, refs, firstRef)
    if (rectangle === null) {
      return false
    }
    const { sheet, sheetId, firstRow, rowCount, firstCol, colCount, inputCellIndices, inputNumericValues, rowSums } = rectangle
    for (let col = firstCol; col < firstCol + colCount; col += 1) {
      if (args.hasTrackedExactLookupDependents(sheetId, col) || args.hasTrackedSortedLookupDependents(sheetId, col)) {
        return false
      }
    }
    const formulas = collectRowAggregateFormulas(args, {
      sheetName: sheet.name,
      sheetId,
      firstRow,
      rowCount,
      firstCol,
      colCount,
    })
    if (formulas === null) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + rowCount + 1)
    args.resetMaterializedCellScratch(0)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < inputCellIndices.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        if (rectangle.emptyInputFlags[index] === 1) {
          writeEmptyLiteralToCellStore(args.state.workbook.cellStore, cellIndex)
        } else {
          args.writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
        }
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      const writtenColumns = new Uint32Array(colCount)
      for (let index = 0; index < colCount; index += 1) {
        writtenColumns[index] = firstCol + index
      }
      args.state.workbook.notifyColumnsWritten(sheetId, writtenColumns)
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)

    const formulaChanged = requiresChangedSet ? new Uint32Array(rowCount) : EMPTY_CHANGED_CELLS
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
      const formulaCellIndex = formulas.cellIndices[rowOffset]!
      args.applyTerminalDirectFormulaNumericResult(formulaCellIndex, rowSums[rowOffset]! + formulas.resultOffsets[rowOffset]!)
      if (requiresChangedSet) {
        formulaChanged[rowOffset] = formulaCellIndex
      }
    }
    addEngineCounter(args.state.counters, 'directAggregateDeltaApplications', rowCount)
    addEngineCounter(args.state.counters, 'directAggregateDeltaOnlyRecalcSkips')

    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaChanged, explicitChangedCount) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, sheetId, explicitChangedCount)
    }
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed,
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

  return { tryApplyDenseRectangularDirectAggregateLiteralBatch }
}

function collectDenseNumericRectangle(
  args: OperationDirectAggregateRectangularBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
  firstRef: EngineCellMutationRef,
): {
  readonly sheet: NonNullable<ReturnType<FastPathState['workbook']['getSheetById']>>
  readonly sheetId: number
  readonly firstRow: number
  readonly rowCount: number
  readonly firstCol: number
  readonly colCount: number
  readonly inputCellIndices: Uint32Array
  readonly inputNumericValues: Float64Array
  readonly emptyInputFlags: Uint8Array
  readonly rowSums: Float64Array
} | null {
  const firstMutation = firstRef.mutation
  if (!isSupportedAggregateBatchMutation(firstMutation)) {
    return null
  }
  const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
  if (!sheet || sheet.structureVersion !== 1) {
    return null
  }
  const inputCellIndices = new Uint32Array(refs.length)
  const inputNumericValues = new Float64Array(refs.length)
  const emptyInputFlags = new Uint8Array(refs.length)
  const rowSumsScratch = new Float64Array(refs.length)
  const cellStore = args.state.workbook.cellStore
  const firstRow = firstMutation.row
  const firstCol = firstMutation.col
  let currentRow = firstRow
  let currentWidth = 0
  let colCount = 0
  let rowOffset = 0
  for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    if (ref.sheetId !== firstRef.sheetId || !isSupportedAggregateBatchMutation(mutation)) {
      return null
    }
    if (mutation.row === currentRow) {
      if (mutation.col !== firstCol + currentWidth) {
        return null
      }
      currentWidth += 1
    } else {
      if (mutation.row !== currentRow + 1 || mutation.col !== firstCol || currentWidth === 0) {
        return null
      }
      if (colCount === 0) {
        colCount = currentWidth
      } else if (currentWidth !== colCount) {
        return null
      }
      currentRow = mutation.row
      currentWidth = 1
      rowOffset += 1
    }
    const existingIndex =
      ref.cellIndex !== undefined &&
      cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
      cellStore.rows[ref.cellIndex] === mutation.row &&
      cellStore.cols[ref.cellIndex] === mutation.col
        ? ref.cellIndex
        : sheet.grid.getPhysical(mutation.row, mutation.col)
    if (
      existingIndex === -1 ||
      !args.canFastPathLiteralOverwrite(existingIndex) ||
      args.getSingleEntityDependent(makeCellEntity(existingIndex)) !== -1
    ) {
      return null
    }
    inputCellIndices[refIndex] = existingIndex
    if (mutation.kind === 'clearCell' || mutation.value === null) {
      emptyInputFlags[refIndex] = 1
      inputNumericValues[refIndex] = 0
    } else {
      const numericValue = mutation.value
      if (typeof numericValue !== 'number') {
        return null
      }
      inputNumericValues[refIndex] = numericValue
      rowSumsScratch[rowOffset] = (rowSumsScratch[rowOffset] ?? 0) + numericValue
    }
  }
  if (colCount === 0) {
    colCount = currentWidth
  } else if (currentWidth !== colCount) {
    return null
  }
  const rowCount = rowOffset + 1
  if (rowCount < 2 || colCount < 2 || rowCount * colCount !== refs.length) {
    return null
  }
  return {
    sheet,
    sheetId: firstRef.sheetId,
    firstRow,
    rowCount,
    firstCol,
    colCount,
    inputCellIndices,
    inputNumericValues,
    emptyInputFlags,
    rowSums: rowSumsScratch.subarray(0, rowCount),
  }
}

function isSupportedAggregateBatchMutation(mutation: EngineCellMutationRef['mutation']): mutation is DenseAggregateBatchMutation {
  if (mutation.kind === 'clearCell') {
    return true
  }
  return (
    mutation.kind === 'setCellValue' && (mutation.value === null || (typeof mutation.value === 'number' && !Object.is(mutation.value, -0)))
  )
}

function writeEmptyLiteralToCellStore(cellStore: FastPathState['workbook']['cellStore'], cellIndex: number): void {
  const flags = cellStore.flags[cellIndex] ?? 0
  cellStore.flags[cellIndex] = flags & ~(CellFlags.AuthoredBlank | CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  cellStore.stringIds[cellIndex] = 0
  cellStore.tags[cellIndex] = ValueTag.Empty
  cellStore.numbers[cellIndex] = 0
  cellStore.errors[cellIndex] = ErrorCode.None
}

function collectRowAggregateFormulas(
  args: OperationDirectAggregateRectangularBatchFastPathArgs,
  rectangle: {
    readonly sheetName: string
    readonly sheetId: number
    readonly firstRow: number
    readonly rowCount: number
    readonly firstCol: number
    readonly colCount: number
  },
): { readonly cellIndices: Uint32Array; readonly resultOffsets: Float64Array } | null {
  const cellIndices = new Int32Array(rectangle.rowCount)
  cellIndices.fill(-1)
  const resultOffsets = new Float64Array(rectangle.rowCount)
  let invalid = false
  const rowEnd = rectangle.firstRow + rectangle.rowCount - 1
  const colEnd = rectangle.firstCol + rectangle.colCount - 1
  args.state.formulas.forEach((formula, cellIndex) => {
    if (invalid) {
      return
    }
    if (
      formula.directCriteria !== undefined &&
      directCriteriaOverlapsRectangle(formula.directCriteria, rectangle.sheetName, rectangle.firstRow, rowEnd, rectangle.firstCol, colEnd)
    ) {
      invalid = true
      return
    }
    const aggregate = formula.directAggregate
    if (
      aggregate === undefined ||
      !directAggregateOverlapsRectangle(aggregate, rectangle.sheetName, rectangle.firstRow, rowEnd, rectangle.firstCol, colEnd)
    ) {
      return
    }
    if (
      aggregate.aggregateKind !== 'sum' ||
      aggregate.sheetName !== rectangle.sheetName ||
      aggregate.rowStart !== aggregate.rowEnd ||
      aggregate.rowStart < rectangle.firstRow ||
      aggregate.rowStart > rowEnd ||
      aggregate.col !== rectangle.firstCol ||
      aggregate.colEnd !== colEnd ||
      formula.dependencyIndices.length !== 0 ||
      formula.directCriteria !== undefined ||
      formula.directLookup !== undefined ||
      formula.directScalar !== undefined ||
      !args.canUseDirectFormulaPostRecalc(cellIndex) ||
      !args.canSkipFormulaColumnVersion(cellIndex) ||
      ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0
    ) {
      invalid = true
      return
    }
    const formulaSheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const formulaRow = args.state.workbook.cellStore.rows[cellIndex]
    const formulaCol = args.state.workbook.cellStore.cols[cellIndex]
    if (
      formulaSheetId !== rectangle.sheetId ||
      formulaRow !== aggregate.rowStart ||
      formulaCol === undefined ||
      (formulaCol >= rectangle.firstCol && formulaCol <= colEnd)
    ) {
      invalid = true
      return
    }
    const rowOffset = aggregate.rowStart - rectangle.firstRow
    if (cellIndices[rowOffset] !== -1) {
      invalid = true
      return
    }
    cellIndices[rowOffset] = cellIndex
    resultOffsets[rowOffset] = aggregate.resultOffset ?? 0
  })
  if (invalid) {
    return null
  }
  const unsignedCellIndices = new Uint32Array(rectangle.rowCount)
  for (let index = 0; index < rectangle.rowCount; index += 1) {
    const cellIndex = cellIndices[index]!
    if (cellIndex < 0) {
      return null
    }
    unsignedCellIndices[index] = cellIndex
  }
  return { cellIndices: unsignedCellIndices, resultOffsets }
}

function directAggregateOverlapsRectangle(
  aggregate: RuntimeDirectAggregateDescriptor,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  return (
    aggregate.sheetName === sheetName &&
    aggregate.rowEnd >= rowStart &&
    aggregate.rowStart <= rowEnd &&
    aggregate.colEnd >= colStart &&
    aggregate.col <= colEnd
  )
}

function directCriteriaOverlapsRectangle(
  criteria: RuntimeDirectCriteriaDescriptor,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  const rangeOverlaps = (range: {
    readonly sheetName: string
    readonly rowStart: number
    readonly rowEnd: number
    readonly col: number
  }): boolean =>
    range.sheetName === sheetName && range.rowEnd >= rowStart && range.rowStart <= rowEnd && range.col >= colStart && range.col <= colEnd
  if (criteria.aggregateRange && rangeOverlaps(criteria.aggregateRange)) {
    return true
  }
  return criteria.criteriaPairs.some((pair) => rangeOverlaps(pair.range))
}
