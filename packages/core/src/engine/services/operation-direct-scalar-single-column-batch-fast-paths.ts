import type { EngineOpBatch } from '@bilig/workbook-domain'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markBatchApplied } from '../../replica-state.js'
import type { OperationDirectScalarBatchFastPathArgs } from './operation-direct-scalar-batch-fast-paths.js'
import {
  directScalarCellNumber,
  evaluateDirectScalarWithReplacementNumbers,
  singleInputAffineDirectScalar,
} from './direct-scalar-helpers.js'
import { tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

export function createOperationDirectScalarSingleColumnBatchFastPaths(args: OperationDirectScalarBatchFastPathArgs): {
  readonly tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyDenseSingleColumnDirectScalarLiteralBatch: (
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
    writeNumericLiteralToCellStore,
    applyTerminalDirectFormulaNumericResult,
  } = args

  const tryApplyDenseSingleColumnDirectScalarLiteralBatch = (
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
    const secondMutation = refs[1]?.mutation
    if (secondMutation?.kind !== 'setCellValue') {
      return false
    }
    const rowOrder = secondMutation.row > firstMutation.row ? 1 : secondMutation.row < firstMutation.row ? -1 : 0
    if (rowOrder === 0) {
      return false
    }
    const firstSheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !firstSheet ||
      firstSheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col)
    ) {
      return false
    }
    const inputCellIndices = new Uint32Array(refs.length)
    const formulaCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const formulaNumericResults = new Float64Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    const readDirectScalarNumber = (cellIndex: number): number | undefined => directScalarCellNumber(cellStore, cellIndex)
    let previousRow = firstMutation.row
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (refIndex > 0) {
        if ((rowOrder > 0 && mutation.row <= previousRow) || (rowOrder < 0 && mutation.row >= previousRow)) {
          return false
        }
        previousRow = mutation.row
      }
      if (
        ref.sheetId !== firstRef.sheetId ||
        mutation.kind !== 'setCellValue' ||
        mutation.col !== firstMutation.col ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0)
      ) {
        return false
      }
      const existingIndex =
        ref.cellIndex !== undefined &&
        args.state.workbook.cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
        args.state.workbook.cellStore.rows[ref.cellIndex] === mutation.row &&
        args.state.workbook.cellStore.cols[ref.cellIndex] === mutation.col
          ? ref.cellIndex
          : firstSheet.grid.getPhysical(mutation.row, mutation.col)
      if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex)) {
        return false
      }
      const singleDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
      if (singleDependent < 0 || !canUseDirectFormulaPostRecalc(singleDependent) || !canSkipFormulaColumnVersion(singleDependent)) {
        return false
      }
      const formula = args.state.formulas.get(singleDependent)
      const result =
        formula?.directScalar === undefined
          ? undefined
          : evaluateDirectScalarWithReplacementNumbers(formula.directScalar, existingIndex, mutation.value, readDirectScalarNumber)
      if (result === undefined) {
        return false
      }
      const outputIndex = rowOrder < 0 ? refs.length - 1 - refIndex : refIndex
      inputCellIndices[outputIndex] = existingIndex
      formulaCellIndices[outputIndex] = singleDependent
      inputNumericValues[outputIndex] = mutation.value
      formulaNumericResults[outputIndex] = result
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    const reservedNewCells = potentialNewCells ?? 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length * 2 + 1)
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
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    const formulaChanged = requiresChangedSet ? new Uint32Array(refs.length) : EMPTY_CHANGED_CELLS
    for (let index = 0; index < refs.length; index += 1) {
      const formulaCellIndex = formulaCellIndices[index]!
      applyTerminalDirectFormulaNumericResult(formulaCellIndex, formulaNumericResults[index]!)
      if (requiresChangedSet) {
        formulaChanged[index] = formulaCellIndex
      }
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', refs.length)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaChanged, explicitChangedCount) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstRef.sheetId, explicitChangedCount)
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

  const tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32 || (potentialNewCells ?? 0) !== 0) {
      return false
    }
    const firstMutation = firstRef.mutation
    if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
      return false
    }
    const secondMutation = refs[1]?.mutation
    if (secondMutation?.kind !== 'setCellValue') {
      return false
    }
    const rowOrder = secondMutation.row > firstMutation.row ? 1 : secondMutation.row < firstMutation.row ? -1 : 0
    if (rowOrder === 0) {
      return false
    }
    const firstSheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !firstSheet ||
      firstSheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col)
    ) {
      return false
    }

    const inputCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const formulaCellIndices = new Uint32Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    let previousRow = firstMutation.row
    let previousFormulaRow = -1
    let previousFormulaCol = -1
    let affineScale: number | undefined
    let affineOffset: number | undefined
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (refIndex > 0) {
        if ((rowOrder > 0 && mutation.row <= previousRow) || (rowOrder < 0 && mutation.row >= previousRow)) {
          return false
        }
        previousRow = mutation.row
      }
      if (
        ref.sheetId !== firstRef.sheetId ||
        mutation.kind !== 'setCellValue' ||
        mutation.col !== firstMutation.col ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0) ||
        ref.cellIndex === undefined ||
        cellStore.sheetIds[ref.cellIndex] !== ref.sheetId ||
        cellStore.rows[ref.cellIndex] !== mutation.row ||
        cellStore.cols[ref.cellIndex] !== mutation.col
      ) {
        return false
      }
      const existingIndex = ref.cellIndex
      if (!canFastPathLiteralOverwrite(existingIndex)) {
        return false
      }
      const singleDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
      if (singleDependent < 0 || !canUseDirectFormulaPostRecalc(singleDependent) || !canSkipFormulaColumnVersion(singleDependent)) {
        return false
      }
      const formula = args.state.formulas.get(singleDependent)
      if (
        !formula ||
        formula.directScalar === undefined ||
        cellStore.sheetIds[singleDependent] !== firstRef.sheetId ||
        cellStore.rows[singleDependent] !== mutation.row
      ) {
        return false
      }
      const formulaRow = cellStore.rows[singleDependent] ?? 0
      const formulaCol = cellStore.cols[singleDependent] ?? 0
      if (
        refIndex > 0 &&
        ((rowOrder > 0 && (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol))) ||
          (rowOrder < 0 && (formulaRow > previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol >= previousFormulaCol))))
      ) {
        return false
      }
      const affine = singleInputAffineDirectScalar(formula.directScalar, existingIndex)
      if (affine === null) {
        return false
      }
      if (affineScale === undefined) {
        affineScale = affine.scale
        affineOffset = affine.offset
      } else if (!Object.is(affineScale, affine.scale) || !Object.is(affineOffset, affine.offset)) {
        return false
      }
      const outputIndex = rowOrder < 0 ? refs.length - 1 - refIndex : refIndex
      inputCellIndices[outputIndex] = existingIndex
      inputNumericValues[outputIndex] = mutation.value
      formulaCellIndices[outputIndex] = singleDependent
      previousFormulaRow = formulaRow
      previousFormulaCol = formulaCol
    }
    if (affineScale === undefined || affineOffset === undefined) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length * 2 + 1)
    args.resetMaterializedCellScratch(0)

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
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    for (let index = 0; index < refs.length; index += 1) {
      applyTerminalDirectFormulaNumericResult(formulaCellIndices[index]!, inputNumericValues[index]! * affineScale + affineOffset)
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', refs.length)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaCellIndices, explicitChangedCount) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstRef.sheetId, explicitChangedCount)
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
    tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch,
    tryApplyDenseSingleColumnDirectScalarLiteralBatch,
  }
}
