import type { EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { writeLiteralToCellStore } from '../../engine-value-utils.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markBatchApplied } from '../../replica-state.js'
import type { EngineRuntimeState, RuntimeDirectScalarDescriptor, U32 } from '../runtime-state.js'
import { directScalarLiteralNumericValue } from './direct-scalar-helpers.js'
import { PendingNumericCellValues, type DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import {
  type ExactLookupBatchSkipPlan,
  exactLookupBatchOldNumeric,
  exactLookupBatchWriteHandled,
  prepareExactLookupBatchSkipPlan,
} from './operation-exact-lookup-batch-skip-plan.js'
import { reverseUint32Array, tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'
import { createOperationDirectAggregateRectangularBatchFastPath } from './operation-direct-aggregate-rectangular-batch-fast-path.js'
import { createOperationFreshRectangularLiteralBatchFastPath } from './operation-fresh-rectangular-literal-batch-fast-path.js'
import { DirectScalarPhysicalSliceTracker } from './operation-direct-scalar-physical-slice-tracker.js'
import { createOperationDirectScalarRowPairBatchFastPaths } from './operation-direct-scalar-row-pair-batch-fast-paths.js'
import { createOperationDirectScalarSingleColumnBatchFastPaths } from './operation-direct-scalar-single-column-batch-fast-paths.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

type MutationSource = 'local' | 'restore' | 'undo' | 'redo'

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
>

export interface OperationDirectScalarBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly hasVolatileFormulas: (() => boolean) | undefined
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly canUseDirectFormulaPostRecalc: (cellIndex: number) => boolean
  readonly canSkipFormulaColumnVersion: (cellIndex: number) => boolean
  readonly directScalarCellNumericValue: (cellIndex: number) => number | undefined
  readonly writeNumericLiteralToCellStore: (cellIndex: number, value: number) => void
  readonly applyTerminalDirectFormulaNumericResult: (cellIndex: number, value: number) => void
  readonly applyDirectFormulaNumericResult: (cellIndex: number, value: number) => void
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, value: DirectScalarCurrentOperand) => boolean
  readonly tryEvaluateDirectScalarWithPendingNumbers: (
    directScalar: RuntimeDirectScalarDescriptor,
    pendingNumbers: PendingNumericCellValues,
  ) => DirectScalarCurrentOperand | undefined
  readonly tryEvaluateDirectScalarNumericWithPendingNumbers: (
    directScalar: RuntimeDirectScalarDescriptor,
    pendingNumbers: PendingNumericCellValues,
  ) => number | undefined
  readonly planExactLookupNumericColumnWrite: (
    sheetId: number,
    col: number,
    row: number,
    oldValue: number,
    newValue: number,
  ) => { readonly handled: boolean }
  readonly planApproximateLookupNumericColumnWrite: (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldValue: number,
    newValue: number,
  ) => { readonly handled: boolean }
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly getEntityDependents: (entityId: number) => Uint32Array
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
  readonly invalidateExactLookupColumn: (request: { readonly sheetName: string; readonly col: number }) => void
  readonly invalidateSortedLookupColumn: (request: { readonly sheetName: string; readonly col: number }) => void
}

export function createOperationDirectScalarBatchFastPaths(args: OperationDirectScalarBatchFastPathArgs): {
  readonly tryApplyCoalescedDirectScalarLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: MutationSource,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyDenseRowPairDirectScalarLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyLookupOnlyNumericColumnLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyDenseRectangularDirectAggregateLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyFreshDenseRectangularNumericLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
} {
  const {
    emitBatch,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    canFastPathLiteralOverwrite,
    canUseDirectFormulaPostRecalc,
    canSkipFormulaColumnVersion,
    directScalarCellNumericValue,
    writeNumericLiteralToCellStore,
    applyTerminalDirectFormulaNumericResult,
    applyDirectFormulaNumericResult,
    applyDirectFormulaCurrentResult,
    tryEvaluateDirectScalarWithPendingNumbers,
    tryEvaluateDirectScalarNumericWithPendingNumbers,
    planExactLookupNumericColumnWrite,
    planApproximateLookupNumericColumnWrite,
  } = args

  const { tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch, tryApplyDenseSingleColumnDirectScalarLiteralBatch } =
    createOperationDirectScalarSingleColumnBatchFastPaths(args)

  const { tryApplyDenseRowPairSimpleDirectScalarLiteralBatch, tryApplyDenseRowPairDirectScalarLiteralBatch } =
    createOperationDirectScalarRowPairBatchFastPaths({
      state: args.state,
      emitBatch,
      hasTrackedExactLookupDependents,
      hasTrackedSortedLookupDependents,
      hasTrackedDirectRangeDependents,
      canFastPathLiteralOverwrite,
      canUseDirectFormulaPostRecalc,
      canSkipFormulaColumnVersion,
      writeNumericLiteralToCellStore,
      applyTerminalDirectFormulaNumericResult,
      getEntityDependents: args.getEntityDependents,
      materializeDeferredStructuralFormulaSources: args.materializeDeferredStructuralFormulaSources,
      beginMutationCollection: args.beginMutationCollection,
      ensureRecalcScratchCapacity: args.ensureRecalcScratchCapacity,
      resetMaterializedCellScratch: args.resetMaterializedCellScratch,
      getBatchMutationDepth: args.getBatchMutationDepth,
      setBatchMutationDepth: args.setBatchMutationDepth,
      markInputChanged: args.markInputChanged,
      markExplicitChanged: args.markExplicitChanged,
      getChangedInputBuffer: args.getChangedInputBuffer,
      deferKernelSync: args.deferKernelSync,
      composeDisjointEventChanges: args.composeDisjointEventChanges,
      captureChangedCells: args.captureChangedCells,
    })

  const { tryApplyDenseRectangularDirectAggregateLiteralBatch } = createOperationDirectAggregateRectangularBatchFastPath({
    state: args.state,
    emitBatch,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    canFastPathLiteralOverwrite,
    canUseDirectFormulaPostRecalc,
    canSkipFormulaColumnVersion,
    writeNumericLiteralToCellStore,
    applyTerminalDirectFormulaNumericResult,
    getSingleEntityDependent: args.getSingleEntityDependent,
    materializeDeferredStructuralFormulaSources: args.materializeDeferredStructuralFormulaSources,
    beginMutationCollection: args.beginMutationCollection,
    ensureRecalcScratchCapacity: args.ensureRecalcScratchCapacity,
    resetMaterializedCellScratch: args.resetMaterializedCellScratch,
    getBatchMutationDepth: args.getBatchMutationDepth,
    setBatchMutationDepth: args.setBatchMutationDepth,
    markInputChanged: args.markInputChanged,
    markExplicitChanged: args.markExplicitChanged,
    getChangedInputBuffer: args.getChangedInputBuffer,
    deferKernelSync: args.deferKernelSync,
    composeDisjointEventChanges: args.composeDisjointEventChanges,
    captureChangedCells: args.captureChangedCells,
  })

  const { tryApplyFreshDenseRectangularNumericLiteralBatch } = createOperationFreshRectangularLiteralBatchFastPath({
    state: args.state,
    emitBatch,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    writeNumericLiteralToCellStore,
    materializeDeferredStructuralFormulaSources: args.materializeDeferredStructuralFormulaSources,
    beginMutationCollection: args.beginMutationCollection,
    ensureRecalcScratchCapacity: args.ensureRecalcScratchCapacity,
    resetMaterializedCellScratch: args.resetMaterializedCellScratch,
    getBatchMutationDepth: args.getBatchMutationDepth,
    setBatchMutationDepth: args.setBatchMutationDepth,
    markInputChanged: args.markInputChanged,
    markExplicitChanged: args.markExplicitChanged,
    getChangedInputBuffer: args.getChangedInputBuffer,
    deferKernelSync: args.deferKernelSync,
    captureChangedCells: args.captureChangedCells,
  })

  const tryApplyLookupOnlyNumericColumnLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32) {
      return false
    }
    const firstMutation = firstRef.mutation
    if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (!sheet || sheet.structureVersion !== 1 || hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col)) {
      return false
    }
    const hasExactLookupDependents = hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col)
    const hasSortedLookupDependents = hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col)
    if (!hasExactLookupDependents && !hasSortedLookupDependents) {
      return false
    }
    const exactSkipPlan = hasExactLookupDependents
      ? prepareExactLookupBatchSkipPlan({
          col: firstMutation.col,
          formulas: args.state.formulas,
          getEntityDependents: args.getEntityDependents,
          getSingleEntityDependent: args.getSingleEntityDependent,
          readNumericValue: directScalarCellNumericValue,
          sheetId: firstRef.sheetId,
        })
      : ({ kind: 'all' } satisfies ExactLookupBatchSkipPlan)

    const inputCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    let ascending = true
    let descending = true
    let previousRow = -1
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (
        ref.sheetId !== firstRef.sheetId ||
        mutation.kind !== 'setCellValue' ||
        mutation.col !== firstMutation.col ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0)
      ) {
        return false
      }
      if (refIndex > 0) {
        ascending &&= mutation.row > previousRow
        descending &&= mutation.row < previousRow
      }
      previousRow = mutation.row
      const existingIndex =
        ref.cellIndex !== undefined &&
        cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
        cellStore.rows[ref.cellIndex] === mutation.row &&
        cellStore.cols[ref.cellIndex] === mutation.col
          ? ref.cellIndex
          : sheet.grid.getPhysical(mutation.row, mutation.col)
      if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex)) {
        return false
      }
      const plannedOldNumber = exactLookupBatchOldNumeric(exactSkipPlan, mutation.row)
      const oldNumber =
        plannedOldNumber !== undefined && !hasSortedLookupDependents ? plannedOldNumber : directScalarCellNumericValue(existingIndex)
      if (oldNumber === undefined) {
        return false
      }
      if (
        !(
          exactLookupBatchWriteHandled(exactSkipPlan, mutation.row, oldNumber, mutation.value) ||
          (exactSkipPlan.kind === 'fallback' &&
            planExactLookupNumericColumnWrite(firstRef.sheetId, firstMutation.col, mutation.row, oldNumber, mutation.value).handled)
        )
      ) {
        return false
      }
      if (
        hasSortedLookupDependents &&
        !planApproximateLookupNumericColumnWrite(firstRef.sheetId, sheet.name, firstMutation.col, mutation.row, oldNumber, mutation.value)
          .handled
      ) {
        return false
      }
      inputCellIndices[refIndex] = existingIndex
      inputNumericValues[refIndex] = mutation.value
    }
    if (!ascending && !descending) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    const reservedNewCells = potentialNewCells ?? 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (hasExactLookupDependents) {
      args.invalidateExactLookupColumn({ sheetName: sheet.name, col: firstMutation.col })
    }
    if (hasSortedLookupDependents) {
      args.invalidateSortedLookupColumn({ sheetName: sheet.name, col: firstMutation.col })
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    const changed = requiresChangedSet ? (ascending ? inputCellIndices : reverseUint32Array(inputCellIndices)) : EMPTY_CHANGED_CELLS
    addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
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
      emitBatch,
    })
    return true
  }

  const tryApplyCoalescedDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ): boolean => {
    if (
      source === 'restore' ||
      (source !== 'local' && batch !== null) ||
      refs.length < 32 ||
      args.state.formulas.size === 0 ||
      args.state.workbook.hasPivots()
    ) {
      return false
    }
    if (args.hasVolatileFormulas?.()) {
      return false
    }
    if (tryApplyLookupOnlyNumericColumnLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyFreshDenseRectangularNumericLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseRectangularDirectAggregateLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseSingleColumnDirectScalarLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseRowPairSimpleDirectScalarLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }
    if (tryApplyDenseRowPairDirectScalarLiteralBatch(refs, batch, potentialNewCells)) {
      return true
    }

    const pendingNumbers = new PendingNumericCellValues()
    const inputCellIndices: number[] = []
    const inputNumericValues = new Float64Array(refs.length)
    const formulaCellIndices: number[] = []
    const formulaSeen = new Set<number>()
    let canUseNumericInputWrites = true
    const physicalSliceTracker = new DirectScalarPhysicalSliceTracker()
    const trackedColumnDependencyFlagsBySheet = new Map<number, Map<number, boolean>>()
    const hasTrackedColumnDependencies = (sheetId: number, col: number): boolean => {
      let flagsByColumn = trackedColumnDependencyFlagsBySheet.get(sheetId)
      if (flagsByColumn === undefined) {
        flagsByColumn = new Map()
        trackedColumnDependencyFlagsBySheet.set(sheetId, flagsByColumn)
      }
      const cached = flagsByColumn.get(col)
      if (cached !== undefined) {
        return cached
      }
      const hasDependencies =
        hasTrackedExactLookupDependents(sheetId, col) ||
        hasTrackedSortedLookupDependents(sheetId, col) ||
        hasTrackedDirectRangeDependents(sheetId, col)
      flagsByColumn.set(col, hasDependencies)
      return hasDependencies
    }
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (mutation.kind !== 'setCellValue') {
        return false
      }
      if (typeof mutation.value !== 'number' || Object.is(mutation.value, -0)) {
        canUseNumericInputWrites = false
      }
      const nextNumber = directScalarLiteralNumericValue(mutation.value)
      if (nextNumber === undefined) {
        return false
      }
      inputNumericValues[refIndex] = nextNumber
      const sheet = args.state.workbook.getSheetById(ref.sheetId)
      if (!sheet || sheet.structureVersion !== 1) {
        return false
      }
      const candidate = ref.cellIndex
      const existingIndex =
        candidate !== undefined &&
        args.state.workbook.cellStore.sheetIds[candidate] === ref.sheetId &&
        args.state.workbook.cellStore.rows[candidate] === mutation.row &&
        args.state.workbook.cellStore.cols[candidate] === mutation.col
          ? candidate
          : sheet.grid.getPhysical(mutation.row, mutation.col)
      if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex) || pendingNumbers.has(existingIndex)) {
        return false
      }
      physicalSliceTracker.noteCell(ref.sheetId, mutation.row, mutation.col, 'input')
      if (hasTrackedColumnDependencies(ref.sheetId, mutation.col)) {
        return false
      }
      const dependents = args.getEntityDependents(makeCellEntity(existingIndex))
      for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
        const formulaCellIndex = dependents[dependentIndex]!
        const formula = args.state.formulas.get(formulaCellIndex)
        if (
          !formula ||
          formula.directScalar === undefined ||
          !canUseDirectFormulaPostRecalc(formulaCellIndex) ||
          ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
        ) {
          return false
        }
        if (!formulaSeen.has(formulaCellIndex)) {
          formulaSeen.add(formulaCellIndex)
          formulaCellIndices.push(formulaCellIndex)
          const cellStore = args.state.workbook.cellStore
          const formulaSheetId = cellStore.sheetIds[formulaCellIndex]
          const formulaSheet = formulaSheetId === undefined ? undefined : args.state.workbook.getSheetById(formulaSheetId)
          if (formulaSheetId === undefined || (formulaSheet && formulaSheet.structureVersion !== 1)) {
            physicalSliceTracker.markUntrusted()
          } else {
            physicalSliceTracker.noteCell(
              formulaSheetId,
              cellStore.rows[formulaCellIndex] ?? 0,
              cellStore.cols[formulaCellIndex] ?? 0,
              'formula',
            )
          }
        }
      }
      pendingNumbers.set(existingIndex, nextNumber)
      inputCellIndices.push(existingIndex)
    }
    if (inputCellIndices.length === 0 || formulaCellIndices.length === 0) {
      return false
    }
    const formulaNumericResults = new Float64Array(formulaCellIndices.length)
    let canUseNumericFormulaResults = true
    for (let index = 0; index < formulaCellIndices.length; index += 1) {
      const formula = args.state.formulas.get(formulaCellIndices[index]!)
      const result = formula?.directScalar
        ? tryEvaluateDirectScalarNumericWithPendingNumbers(formula.directScalar, pendingNumbers)
        : undefined
      if (result === undefined) {
        canUseNumericFormulaResults = false
        break
      }
      formulaNumericResults[index] = result
    }
    let formulaResults: DirectScalarCurrentOperand[] | undefined
    if (!canUseNumericFormulaResults) {
      const evaluatedFormulaResults: DirectScalarCurrentOperand[] = []
      for (let index = 0; index < formulaCellIndices.length; index += 1) {
        const formula = args.state.formulas.get(formulaCellIndices[index]!)
        const result = formula?.directScalar ? tryEvaluateDirectScalarWithPendingNumbers(formula.directScalar, pendingNumbers) : undefined
        if (result === undefined) {
          return false
        }
        evaluatedFormulaResults[index] = result
      }
      formulaResults = evaluatedFormulaResults
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    const reservedNewCells = potentialNewCells ?? 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + formulaCellIndices.length + refs.length + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let index = 0; index < refs.length; index += 1) {
          const cellIndex = inputCellIndices[index]!
          if (canUseNumericInputWrites) {
            writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
          } else {
            const mutation = refs[index]!.mutation
            if (mutation.kind !== 'setCellValue') {
              throw new Error('Expected coalesced direct scalar batch to contain only literal writes')
            }
            writeLiteralToCellStore(args.state.workbook.cellStore, cellIndex, mutation.value, args.state.strings)
          }
          args.state.workbook.notifyCellValueWritten(cellIndex)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
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
    const formulaChanged = requiresChangedSet ? new Uint32Array(formulaCellIndices.length) : EMPTY_CHANGED_CELLS
    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      for (let index = 0; index < formulaCellIndices.length; index += 1) {
        const cellIndex = formulaCellIndices[index]!
        if (canUseNumericFormulaResults) {
          applyDirectFormulaNumericResult(cellIndex, formulaNumericResults[index]!)
        } else if (!applyDirectFormulaCurrentResult(cellIndex, formulaResults![index]!)) {
          throw new Error('Failed to apply direct scalar batch result')
        }
        if (requiresChangedSet) {
          formulaChanged[index] = cellIndex
        }
      }
    })
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', formulaCellIndices.length)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')

    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaChanged, explicitChangedCount) : EMPTY_CHANGED_CELLS
    const trustedPhysicalSheetId = hasTrackedEventListeners
      ? physicalSliceTracker.getTrustedSheetIdForTrackedChanges(explicitChangedCount, changed.length)
      : undefined
    if (requiresChangedSet && trustedPhysicalSheetId !== undefined) {
      tagTrustedPhysicalTrackedChanges(changed, trustedPhysicalSheetId, explicitChangedCount)
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
      emitBatch,
    })
    return true
  }

  return {
    tryApplyCoalescedDirectScalarLiteralBatch,
    tryApplyDenseRowPairDirectScalarLiteralBatch,
    tryApplyLookupOnlyNumericColumnLiteralBatch,
    tryApplyDenseRectangularDirectAggregateLiteralBatch,
    tryApplyFreshDenseRectangularNumericLiteralBatch,
  }
}
