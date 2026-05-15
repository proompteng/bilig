import { ValueTag, type CellValue, type EngineEvent, type LiteralInput } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import type {
  EngineCellMutationRef,
  EngineExistingLiteralCellMutationRef,
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationResult,
} from '../../cell-mutations-at.js'
import { writeLiteralToCellStore } from '../../engine-value-utils.js'
import { isRangeEntity, makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeDirectScalarDescriptor, U32 } from '../runtime-state.js'
import type { SheetRecord } from '../../workbook-store.js'
import { DirectFormulaIndexCollection, type DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import {
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  hasCompleteDirectFormulaDeltas,
} from './direct-formula-recalc-helpers.js'
import { directScalarLiteralNumericValue } from './direct-scalar-helpers.js'
import {
  canTrustPhysicalTrackedChangeSplit,
  makeExistingNumericMutationResult,
  tagTrustedPhysicalTrackedChanges,
} from './operation-change-helpers.js'
import { tryApplySinglePostRecalcDirectFormula, type DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'

const DIRECT_RANGE_POST_RECALC_LIMIT = 16_384
const EMPTY_CHANGED_CELLS = new Uint32Array(0)

type MutationSource = 'local' | 'restore' | 'undo' | 'redo'

type SingleExistingLiteralState = Pick<
  EngineRuntimeState,
  'workbook' | 'strings' | 'events' | 'formulas' | 'counters' | 'trackReplicaVersions' | 'getLastMetrics' | 'setLastMetrics'
>

interface LookupTailPatchTarget {
  tailPatch?: {
    row: number
    oldNumeric: number
    newNumeric: number
    columnVersion: number
  }
}

interface LookupWritePlan {
  readonly handled: boolean
  readonly tailPatchTarget?: LookupTailPatchTarget
}

interface OperationSingleExistingLiteralFastPathArgs {
  readonly state: SingleExistingLiteralState
  readonly hasVolatileFormulas: (() => boolean) | undefined
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly noteAggregateLiteralWrite: (request: {
    readonly sheetName: string
    readonly row: number
    readonly col: number
    readonly oldValue: CellValue
    readonly newValue: CellValue
  }) => void
  readonly evaluateDirectFormula: (cellIndex: number) => readonly number[] | undefined
  readonly invalidateExactLookupColumn: (request: { readonly sheetName: string; readonly col: number }) => void
  readonly invalidateSortedLookupColumn: (request: { readonly sheetName: string; readonly col: number }) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly canSkipApproximateLookupNewNumericColumnWrite: (sheetId: number, col: number, row: number) => boolean
  readonly writeNumericLiteralToExistingCell: (cellIndex: number, value: number) => void
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => EngineEvent['metrics']
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly directScalarCellNumericValue: (cellIndex: number) => number | undefined
  readonly tryApplySingleDirectAggregateLiteralMutationFastPath: (request: {
    readonly existingIndex: number
    readonly sheetId?: number
    readonly sheetName: string
    readonly row: number
    readonly col: number
    readonly value: LiteralInput
    readonly delta: number
    readonly emitTracked: boolean
    readonly singleRangeEntityDependent?: number
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation: (request: {
    readonly existingIndex: number
    readonly rangeEntityDependent: number
    readonly sheet: SheetRecord
    readonly sheetId: number
    readonly col: number
    readonly value: number
    readonly delta: number
    readonly hasExactLookupDependents: boolean
    readonly hasSortedLookupDependents: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedDirectScalarClosureExistingNumericMutation: (request: {
    readonly existingIndex: number
    readonly sheet: SheetRecord
    readonly sheetId: number
    readonly col: number
    readonly value: number
    readonly oldNumber: number
    readonly hasTrackedEventListeners: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedFormulaLeafExistingNumericMutation: (request: {
    readonly existingIndex: number
    readonly formulaCellIndex: number
    readonly sheet: SheetRecord
    readonly col: number
    readonly value: number
    readonly hasTrackedEventListeners: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyFormulaLeafExistingLiteralMutation: (request: {
    readonly existingIndex: number
    readonly formulaCellIndex: number
    readonly value: LiteralInput
    readonly hasTrackedEventListeners: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly planExactLookupNumericColumnWrite: (
    sheetId: number,
    col: number,
    row: number,
    oldValue: number,
    newValue: number,
  ) => LookupWritePlan
  readonly planApproximateLookupNumericColumnWrite: (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldValue: number,
    newValue: number,
  ) => LookupWritePlan
  readonly patchUniformLookupTailWrites: (request: {
    readonly sheetId: number
    readonly col: number
    readonly row: number
    readonly oldNumeric: number
    readonly newNumeric: number
    readonly exact: boolean
    readonly sorted: boolean
  }) => { readonly exact: boolean; readonly sorted: boolean }
  readonly tryApplySingleKernelSyncOnlyLiteralMutationFastPath: (request: {
    readonly existingIndex: number
    readonly value: LiteralInput
    readonly emitTracked: boolean
  }) => boolean
  readonly tryApplySingleDirectFormulaLiteralMutationWithoutEvents: (request: {
    readonly existingIndex: number
    readonly formulaCellIndex: number
    readonly value: LiteralInput
    readonly oldNumber: number
    readonly newNumber: number
    readonly exactLookupValue: number | undefined
    readonly approximateLookupValue: number | undefined
  }) => boolean
  readonly tryApplySingleDirectScalarLiteralMutationWithoutEvents: (request: {
    readonly existingIndex: number
    readonly value: LiteralInput
    readonly oldNumber: number
    readonly newNumber: number
  }) => boolean
  readonly tryApplySingleDirectLookupOperandMutationFastPath: (request: {
    readonly existingIndex: number
    readonly formulaCellIndex: number
    readonly value: LiteralInput
    readonly exactLookupValue: number | undefined
    readonly approximateLookupValue: number | undefined
    readonly emitTracked: boolean
    readonly lookupSheetHint?: SheetRecord | undefined
    readonly trustedInputSheet?: SheetRecord | undefined
    readonly trustedInputCol?: number | undefined
  }) => EngineExistingNumericCellMutationResult | null
  readonly markPostRecalcDirectScalarNumericDependents: (
    cellIndex: number,
    oldNumber: number,
    newNumber: number,
    collection: DirectFormulaIndexCollection,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
  ) => boolean
  readonly tryMarkDirectScalarLinearDeltaClosure: (
    cellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    collection: DirectFormulaIndexCollection,
  ) => boolean
  readonly collectSingleAffectedDirectRangeDependent: (request: {
    readonly sheetName: string
    readonly sheetId?: number
    readonly row: number
    readonly col: number
  }) => number
  readonly collectAffectedDirectRangeDependents: (request: {
    readonly sheetName: string
    readonly row: number
    readonly col: number
  }) => readonly number[]
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, value: DirectScalarCurrentOperand) => boolean
  readonly applyDirectFormulaNumericDelta: (cellIndex: number, delta: number) => boolean
  readonly applyDirectScalarCurrentValue: (cellIndex: number, directScalar: RuntimeDirectScalarDescriptor) => boolean
  readonly tryApplyDirectScalarDeltas: (collection: DirectFormulaIndexCollection, collectChanged?: boolean) => U32 | undefined
  readonly tryApplyDirectFormulaDeltas: (collection: DirectFormulaIndexCollection, collectChanged?: boolean) => U32 | undefined
  readonly countPostRecalcDirectFormulaMetric: (cellIndex: number, counts: DirectFormulaMetricCounts) => void
  readonly hasDynamicFormulaDependents: (cellIndex: number) => boolean
}

export function createOperationSingleExistingLiteralFastPath(args: OperationSingleExistingLiteralFastPathArgs): {
  readonly tryApplySingleExistingDirectLiteralMutation: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: MutationSource,
  ) => boolean
  readonly applyExistingNumericCellMutationAtNow: (
    request: EngineExistingNumericCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly applyExistingLiteralCellMutationAtNow: (
    request: EngineExistingLiteralCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
} {
  const {
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    canSkipApproximateLookupNewNumericColumnWrite,
    writeNumericLiteralToExistingCell,
    deferSingleCellKernelSync,
    makeSingleLiteralSkipMetrics,
    canFastPathLiteralOverwrite,
    directScalarCellNumericValue,
    tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation,
    tryApplyTrustedDirectScalarClosureExistingNumericMutation,
    tryApplyTrustedFormulaLeafExistingNumericMutation,
    tryApplyFormulaLeafExistingLiteralMutation,
    tryApplySingleDirectAggregateLiteralMutationFastPath,
    planExactLookupNumericColumnWrite,
    planApproximateLookupNumericColumnWrite,
    patchUniformLookupTailWrites,
    tryApplySingleKernelSyncOnlyLiteralMutationFastPath,
    tryApplySingleDirectFormulaLiteralMutationWithoutEvents,
    tryApplySingleDirectScalarLiteralMutationWithoutEvents,
    tryApplySingleDirectLookupOperandMutationFastPath,
    markPostRecalcDirectScalarNumericDependents,
    tryMarkDirectScalarLinearDeltaClosure,
    collectSingleAffectedDirectRangeDependent,
    collectAffectedDirectRangeDependents,
    applyDirectFormulaCurrentResult,
    applyDirectFormulaNumericDelta,
    applyDirectScalarCurrentValue,
    tryApplyDirectScalarDeltas,
    tryApplyDirectFormulaDeltas,
    countPostRecalcDirectFormulaMetric,
    hasDynamicFormulaDependents,
  } = args

  const tryApplySingleExistingDirectLiteralMutation = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
  ): boolean => {
    if (
      source !== 'local' ||
      batch !== null ||
      refs.length !== 1 ||
      args.state.workbook.hasPivots() ||
      args.state.events.hasListeners() ||
      args.state.events.hasCellListeners()
    ) {
      return false
    }
    if (args.hasVolatileFormulas?.()) {
      return false
    }
    const ref = refs[0]!
    const mutation = ref.mutation
    if (mutation.kind !== 'setCellValue' || mutation.value === null) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(ref.sheetId)
    if (!sheet || sheet.structureVersion !== 1) {
      return false
    }
    const existingIndex =
      ref.cellIndex !== undefined &&
      args.state.workbook.cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
      args.state.workbook.cellStore.rows[ref.cellIndex] === mutation.row &&
      args.state.workbook.cellStore.cols[ref.cellIndex] === mutation.col
        ? ref.cellIndex
        : sheet.grid.getPhysical(mutation.row, mutation.col)
    const sheetName = sheet.name
    const hasExactLookupDependents = hasTrackedExactLookupDependents(ref.sheetId, mutation.col)
    const hasSortedLookupDependents = hasTrackedSortedLookupDependents(ref.sheetId, mutation.col)
    const hasAggregateDependents = hasTrackedDirectRangeDependents(ref.sheetId, mutation.col)
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    if (existingIndex === -1) {
      if (
        args.state.trackReplicaVersions ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0) ||
        hasExactLookupDependents ||
        hasAggregateDependents ||
        (hasSortedLookupDependents && !canSkipApproximateLookupNewNumericColumnWrite(ref.sheetId, mutation.col, mutation.row))
      ) {
        return false
      }
      const cellIndex = args.state.workbook.ensureCellAt(ref.sheetId, mutation.row, mutation.col).cellIndex
      writeNumericLiteralToExistingCell(cellIndex, mutation.value)
      deferSingleCellKernelSync(cellIndex)
      const lastMetrics = makeSingleLiteralSkipMetrics()
      args.state.setLastMetrics(lastMetrics)
      addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
      if (hasTrackedEventListeners) {
        args.state.events.emitTracked({
          kind: 'batch',
          invalidation: 'cells',
          changedCellIndices: Uint32Array.of(cellIndex),
          invalidatedRanges: [],
          invalidatedRows: [],
          invalidatedColumns: [],
          metrics: lastMetrics,
          explicitChangedCount: 1,
        })
      }
      return true
    }
    if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex)) {
      return false
    }
    if (hasDynamicFormulaDependents(existingIndex)) {
      return false
    }
    const singleExistingCellDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
    const oldNumber = directScalarCellNumericValue(existingIndex)
    const newNumber = directScalarLiteralNumericValue(mutation.value)
    if (oldNumber === undefined || newNumber === undefined) {
      const formulaLeafResult =
        !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents
          ? tryApplyFormulaLeafExistingLiteralMutation({
              existingIndex,
              formulaCellIndex: singleExistingCellDependent,
              value: mutation.value,
              hasTrackedEventListeners,
            })
          : null
      return formulaLeafResult !== null
    }

    if (
      hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      (singleExistingCellDependent === -1 || isRangeEntity(singleExistingCellDependent)) &&
      tryApplySingleDirectAggregateLiteralMutationFastPath({
        existingIndex,
        sheetId: ref.sheetId,
        sheetName,
        row: mutation.row,
        col: mutation.col,
        value: mutation.value,
        delta: newNumber - oldNumber,
        emitTracked: hasTrackedEventListeners,
        ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
      })
    ) {
      return true
    }
    const existingTag = (args.state.workbook.cellStore.tags[existingIndex] as ValueTag | undefined) ?? ValueTag.Empty
    const mutationIsNumber = typeof mutation.value === 'number'
    const directLookupExactMutationNumber = mutationIsNumber ? newNumber : undefined
    const directLookupApproximateMutationNumber = newNumber
    const oldExactLookupNumber = hasExactLookupDependents && existingTag === ValueTag.Number ? oldNumber : undefined
    const newExactLookupNumber = hasExactLookupDependents && mutationIsNumber ? newNumber : undefined
    const oldApproximateLookupNumber = hasSortedLookupDependents ? oldNumber : undefined
    const newApproximateLookupNumber = hasSortedLookupDependents ? newNumber : undefined
    const exactLookupWritePlan =
      hasExactLookupDependents && oldExactLookupNumber !== undefined && newExactLookupNumber !== undefined
        ? planExactLookupNumericColumnWrite(ref.sheetId, mutation.col, mutation.row, oldExactLookupNumber, newExactLookupNumber)
        : { handled: false }
    const sortedLookupWritePlan =
      hasSortedLookupDependents && oldApproximateLookupNumber !== undefined && newApproximateLookupNumber !== undefined
        ? planApproximateLookupNumericColumnWrite(
            ref.sheetId,
            sheetName,
            mutation.col,
            mutation.row,
            oldApproximateLookupNumber,
            newApproximateLookupNumber,
          )
        : { handled: false }
    const exactLookupDependentsHandled = hasExactLookupDependents && exactLookupWritePlan.handled
    const sortedLookupDependentsHandled = hasSortedLookupDependents && sortedLookupWritePlan.handled
    if ((hasExactLookupDependents && !exactLookupDependentsHandled) || (hasSortedLookupDependents && !sortedLookupDependentsHandled)) {
      return false
    }

    const lookupDependentsHandled =
      (hasExactLookupDependents && exactLookupDependentsHandled) || (hasSortedLookupDependents && sortedLookupDependentsHandled)
    const canUseNumericLookupWriteFastPath = lookupDependentsHandled && existingTag === ValueTag.Number && mutationIsNumber
    if (!hasAggregateDependents && (hasExactLookupDependents || hasSortedLookupDependents) && singleExistingCellDependent === -1) {
      if (canUseNumericLookupWriteFastPath) {
        writeNumericLiteralToExistingCell(existingIndex, newNumber)
        const currentColumnVersion = sheet.columnVersions[mutation.col] ?? 0
        if (exactLookupWritePlan.tailPatchTarget !== undefined) {
          exactLookupWritePlan.tailPatchTarget.tailPatch = {
            row: mutation.row,
            oldNumeric: oldNumber,
            newNumeric: newNumber,
            columnVersion: currentColumnVersion,
          }
        }
        if (sortedLookupWritePlan.tailPatchTarget !== undefined) {
          sortedLookupWritePlan.tailPatchTarget.tailPatch = {
            row: mutation.row,
            oldNumeric: oldNumber,
            newNumeric: newNumber,
            columnVersion: currentColumnVersion,
          }
        }
        const needsExactPatch =
          hasExactLookupDependents && exactLookupDependentsHandled && exactLookupWritePlan.tailPatchTarget === undefined
        const needsSortedPatch =
          hasSortedLookupDependents && sortedLookupDependentsHandled && sortedLookupWritePlan.tailPatchTarget === undefined
        const patchedLookupOwners =
          needsExactPatch || needsSortedPatch
            ? patchUniformLookupTailWrites({
                sheetId: ref.sheetId,
                col: mutation.col,
                row: mutation.row,
                oldNumeric: oldNumber,
                newNumeric: newNumber,
                exact: needsExactPatch,
                sorted: needsSortedPatch,
              })
            : { exact: true, sorted: true }
        if (needsExactPatch && !patchedLookupOwners.exact) {
          args.invalidateExactLookupColumn({ sheetName, col: mutation.col })
        }
        if (needsSortedPatch && !patchedLookupOwners.sorted) {
          args.invalidateSortedLookupColumn({ sheetName, col: mutation.col })
        }
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
        deferSingleCellKernelSync(existingIndex)
        const lastMetrics = makeSingleLiteralSkipMetrics()
        args.state.setLastMetrics(lastMetrics)
        if (hasTrackedEventListeners) {
          args.state.events.emitTracked({
            kind: 'batch',
            invalidation: 'cells',
            changedCellIndices: Uint32Array.of(existingIndex),
            invalidatedRanges: [],
            invalidatedRows: [],
            invalidatedColumns: [],
            metrics: lastMetrics,
            explicitChangedCount: 1,
          })
        }
        return true
      }
      if (
        tryApplySingleKernelSyncOnlyLiteralMutationFastPath({
          existingIndex,
          value: mutation.value,
          emitTracked: hasTrackedEventListeners,
        })
      ) {
        return true
      }
    }

    if (!hasTrackedEventListeners && !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents) {
      if (
        tryApplySingleDirectFormulaLiteralMutationWithoutEvents({
          existingIndex,
          formulaCellIndex: singleExistingCellDependent,
          value: mutation.value,
          oldNumber,
          newNumber,
          exactLookupValue: directLookupExactMutationNumber,
          approximateLookupValue: directLookupApproximateMutationNumber,
        }) ||
        tryApplySingleDirectScalarLiteralMutationWithoutEvents({
          existingIndex,
          value: mutation.value,
          oldNumber,
          newNumber,
        })
      ) {
        return true
      }
    }
    if (!hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents) {
      if (
        tryApplySingleDirectLookupOperandMutationFastPath({
          existingIndex,
          formulaCellIndex: singleExistingCellDependent,
          value: mutation.value,
          exactLookupValue: directLookupExactMutationNumber,
          approximateLookupValue: directLookupApproximateMutationNumber,
          emitTracked: hasTrackedEventListeners,
          lookupSheetHint: sheet,
        })
      ) {
        return true
      }
    }
    const oldValue: CellValue = { tag: ValueTag.Number, value: oldNumber }
    const newValue: CellValue = { tag: ValueTag.Number, value: newNumber }
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
    let directDependentsHandled = markPostRecalcDirectScalarNumericDependents(
      existingIndex,
      oldNumber,
      newNumber,
      postRecalcDirectFormulaIndices,
      directLookupExactMutationNumber,
      directLookupApproximateMutationNumber,
    )
    if (
      !directDependentsHandled &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      !hasAggregateDependents &&
      tryMarkDirectScalarLinearDeltaClosure(existingIndex, oldValue, newValue, postRecalcDirectFormulaIndices)
    ) {
      directDependentsHandled = true
    }
    if (
      hasAggregateDependents &&
      directDependentsHandled &&
      postRecalcDirectFormulaIndices.size === 0 &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      tryApplySingleDirectAggregateLiteralMutationFastPath({
        existingIndex,
        sheetId: ref.sheetId,
        sheetName,
        row: mutation.row,
        col: mutation.col,
        value: mutation.value,
        delta: newNumber - oldNumber,
        emitTracked: hasTrackedEventListeners,
        ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
      })
    ) {
      return true
    }
    let shouldNoteAggregateLiteralWrite = false
    if (hasAggregateDependents) {
      if (!directDependentsHandled) {
        directDependentsHandled = true
      }
      const singleAffected = collectSingleAffectedDirectRangeDependent({
        sheetName,
        sheetId: ref.sheetId,
        row: mutation.row,
        col: mutation.col,
      })
      if (singleAffected >= 0) {
        const formula = args.state.formulas.get(singleAffected)
        if (
          !formula ||
          formula.directAggregate?.aggregateKind !== 'sum' ||
          formula.dependencyIndices.length !== 0 ||
          args.getSingleEntityDependent(makeCellEntity(singleAffected)) !== -1
        ) {
          return false
        }
        postRecalcDirectFormulaIndices.addDelta(singleAffected, newNumber - oldNumber)
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(existingIndex)
        shouldNoteAggregateLiteralWrite = true
      } else if (singleAffected === -2) {
        const affected = collectAffectedDirectRangeDependents({
          sheetName,
          row: mutation.row,
          col: mutation.col,
        })
        if (affected.length === 0 || affected.length > DIRECT_RANGE_POST_RECALC_LIMIT) {
          return false
        }
        for (let index = 0; index < affected.length; index += 1) {
          const formulaCellIndex = affected[index]!
          const formula = args.state.formulas.get(formulaCellIndex)
          if (
            !formula ||
            formula.directAggregate?.aggregateKind !== 'sum' ||
            formula.dependencyIndices.length !== 0 ||
            args.getSingleEntityDependent(makeCellEntity(formulaCellIndex)) !== -1
          ) {
            return false
          }
        }
        postRecalcDirectFormulaIndices.appendConstantDelta(affected, newNumber - oldNumber)
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(existingIndex)
        shouldNoteAggregateLiteralWrite = true
      }
    }
    if (
      !directDependentsHandled ||
      (postRecalcDirectFormulaIndices.size > 1 && !hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices))
    ) {
      return false
    }

    const explicitChangedCount = hasTrackedEventListeners ? 1 : 0
    const postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts = {
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    }
    let recalculated: U32 = EMPTY_CHANGED_CELLS
    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      writeLiteralToCellStore(args.state.workbook.cellStore, existingIndex, mutation.value, args.state.strings)
      args.state.workbook.notifyCellValueWritten(existingIndex)
      if (shouldNoteAggregateLiteralWrite) {
        args.noteAggregateLiteralWrite({
          sheetName,
          row: mutation.row,
          col: mutation.col,
          oldValue,
          newValue,
        })
      }
      if (postRecalcDirectFormulaIndices.size > 0) {
        if (hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices)) {
          countDirectFormulaDeltaSkip(args.state.formulas, postRecalcDirectFormulaIndices, args.state.counters)
        } else if (canEvaluatePostRecalcDirectFormulasWithoutKernel(args.state.formulas, postRecalcDirectFormulaIndices)) {
          addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
        }
        const directChanged =
          tryApplyDirectScalarDeltas(postRecalcDirectFormulaIndices, hasTrackedEventListeners) ??
          tryApplyDirectFormulaDeltas(postRecalcDirectFormulaIndices, hasTrackedEventListeners) ??
          tryApplySinglePostRecalcDirectFormula(
            {
              state: args.state,
              collection: postRecalcDirectFormulaIndices,
              recalculated: EMPTY_CHANGED_CELLS,
              didRunRecalc: false,
              metrics: postRecalcDirectFormulaMetrics,
              applyDirectFormulaCurrentResult,
              applyDirectFormulaNumericDelta,
              applyDirectScalarCurrentValue,
              tryApplyDirectScalarDeltas,
              tryApplyDirectFormulaDeltas,
              countPostRecalcDirectFormulaMetric,
              evaluateDirectFormula: args.evaluateDirectFormula,
            },
            hasTrackedEventListeners,
          )
        if (directChanged === undefined) {
          throw new Error('Failed to apply single direct literal mutation fast path')
        }
        recalculated = directChanged
      } else if (hasExactLookupDependents || hasSortedLookupDependents) {
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
      }
    })

    deferSingleCellKernelSync(existingIndex)
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: postRecalcDirectFormulaMetrics.wasmFormulaCount,
      jsFormulaCount: postRecalcDirectFormulaMetrics.jsFormulaCount,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount: 1,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasTrackedEventListeners) {
      const changed = composeSingleDisjointExplicitEventChanges(existingIndex, recalculated)
      if (changed.length > 4 && canTrustPhysicalTrackedChangeSplit(changed, ref.sheetId, explicitChangedCount, args.state.workbook)) {
        tagTrustedPhysicalTrackedChanges(changed, ref.sheetId, explicitChangedCount)
      }
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      })
    }
    return true
  }

  const applyExistingNumericCellMutationAtNow = (
    request: EngineExistingNumericCellMutationRef,
  ): EngineExistingNumericCellMutationResult | null => {
    if (
      args.state.workbook.hasPivots() ||
      args.state.events.hasListeners() ||
      args.state.events.hasCellListeners() ||
      args.hasVolatileFormulas?.()
    ) {
      return null
    }
    const sheet = args.state.workbook.getSheetById(request.sheetId)
    const cellStore = args.state.workbook.cellStore
    const existingIndex = request.cellIndex
    const trustedExistingNumericLiteral = request.trustedExistingNumericLiteral === true
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      (!trustedExistingNumericLiteral &&
        (cellStore.sheetIds[existingIndex] !== request.sheetId ||
          cellStore.rows[existingIndex] !== request.row ||
          cellStore.cols[existingIndex] !== request.col ||
          !canFastPathLiteralOverwrite(existingIndex)))
    ) {
      return null
    }
    if (hasDynamicFormulaDependents(existingIndex)) {
      return null
    }
    const oldNumber = trustedExistingNumericLiteral
      ? request.oldNumericValue === undefined || Object.is(request.oldNumericValue, -0)
        ? 0
        : request.oldNumericValue
      : directScalarCellNumericValue(existingIndex)
    if (oldNumber === undefined || Object.is(request.value, -0)) {
      return null
    }
    const sheetName = sheet.name
    const singleExistingCellDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
    let hasExactLookupDependents: boolean | undefined
    let hasSortedLookupDependents: boolean | undefined
    if (trustedExistingNumericLiteral && request.emitTracked === false && isRangeEntity(singleExistingCellDependent)) {
      hasExactLookupDependents = hasTrackedExactLookupDependents(request.sheetId, request.col)
      hasSortedLookupDependents = hasTrackedSortedLookupDependents(request.sheetId, request.col)
      const trustedAggregateResult = tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation({
        existingIndex,
        rangeEntityDependent: singleExistingCellDependent,
        sheet,
        sheetId: request.sheetId,
        col: request.col,
        value: request.value,
        delta: request.value - oldNumber,
        hasExactLookupDependents,
        hasSortedLookupDependents,
      })
      if (trustedAggregateResult) {
        return trustedAggregateResult
      }
    }
    hasExactLookupDependents ??= hasTrackedExactLookupDependents(request.sheetId, request.col)
    hasSortedLookupDependents ??= hasTrackedSortedLookupDependents(request.sheetId, request.col)
    const hasTrackedEventListeners = request.emitTracked !== false && args.state.events.hasTrackedListeners()
    const hasAggregateDependents =
      isRangeEntity(singleExistingCellDependent) || hasTrackedDirectRangeDependents(request.sheetId, request.col)
    if (
      trustedExistingNumericLiteral &&
      !hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      singleExistingCellDependent >= 0
    ) {
      const scalarClosureResult = tryApplyTrustedDirectScalarClosureExistingNumericMutation({
        existingIndex,
        sheet,
        sheetId: request.sheetId,
        col: request.col,
        value: request.value,
        oldNumber,
        hasTrackedEventListeners,
      })
      if (scalarClosureResult) {
        return scalarClosureResult
      }
      const formulaLeafResult = tryApplyTrustedFormulaLeafExistingNumericMutation({
        existingIndex,
        formulaCellIndex: singleExistingCellDependent,
        sheet,
        col: request.col,
        value: request.value,
        hasTrackedEventListeners,
      })
      if (formulaLeafResult) {
        return formulaLeafResult
      }
    }
    if (!hasAggregateDependents && (hasExactLookupDependents || hasSortedLookupDependents) && singleExistingCellDependent === -1) {
      const exactLookupWritePlan = hasExactLookupDependents
        ? planExactLookupNumericColumnWrite(request.sheetId, request.col, request.row, oldNumber, request.value)
        : { handled: true }
      const sortedLookupWritePlan = hasSortedLookupDependents
        ? planApproximateLookupNumericColumnWrite(request.sheetId, sheetName, request.col, request.row, oldNumber, request.value)
        : { handled: true }
      if (
        exactLookupWritePlan.handled &&
        sortedLookupWritePlan.handled &&
        (exactLookupWritePlan.tailPatchTarget === undefined || exactLookupWritePlan.tailPatchTarget.tailPatch === undefined) &&
        (sortedLookupWritePlan.tailPatchTarget === undefined || sortedLookupWritePlan.tailPatchTarget.tailPatch === undefined)
      ) {
        writeNumericLiteralToExistingCell(existingIndex, request.value)
        const currentColumnVersion = sheet.columnVersions[request.col] ?? 0
        if (exactLookupWritePlan.tailPatchTarget !== undefined) {
          exactLookupWritePlan.tailPatchTarget.tailPatch = {
            row: request.row,
            oldNumeric: oldNumber,
            newNumeric: request.value,
            columnVersion: currentColumnVersion,
          }
        }
        if (sortedLookupWritePlan.tailPatchTarget !== undefined) {
          sortedLookupWritePlan.tailPatchTarget.tailPatch = {
            row: request.row,
            oldNumeric: oldNumber,
            newNumeric: request.value,
            columnVersion: currentColumnVersion,
          }
        }
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
        deferSingleCellKernelSync(existingIndex)
        const lastMetrics = makeSingleLiteralSkipMetrics()
        args.state.setLastMetrics(lastMetrics)
        const changedCellIndices = Uint32Array.of(existingIndex)
        if (hasTrackedEventListeners) {
          args.state.events.emitTracked({
            kind: 'batch',
            invalidation: 'cells',
            changedCellIndices,
            invalidatedRanges: [],
            invalidatedRows: [],
            invalidatedColumns: [],
            metrics: lastMetrics,
            explicitChangedCount: 1,
          })
        }
        return makeExistingNumericMutationResult(changedCellIndices, 1)
      }
    }
    const aggregateFastPathResult =
      hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      (singleExistingCellDependent === -1 || isRangeEntity(singleExistingCellDependent))
        ? tryApplySingleDirectAggregateLiteralMutationFastPath({
            existingIndex,
            sheetId: request.sheetId,
            sheetName,
            row: request.row,
            col: request.col,
            value: request.value,
            delta: request.value - oldNumber,
            emitTracked: hasTrackedEventListeners,
            ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
          })
        : null
    if (aggregateFastPathResult) {
      return aggregateFastPathResult
    }
    const directLookupFastPathResult =
      !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents
        ? tryApplySingleDirectLookupOperandMutationFastPath({
            existingIndex,
            formulaCellIndex: singleExistingCellDependent,
            value: request.value,
            exactLookupValue: request.value,
            approximateLookupValue: request.value,
            emitTracked: hasTrackedEventListeners,
            lookupSheetHint: sheet,
            ...(trustedExistingNumericLiteral ? { trustedInputSheet: sheet, trustedInputCol: request.col } : {}),
          })
        : null
    if (directLookupFastPathResult) {
      return directLookupFastPathResult
    }
    return null
  }

  const applyExistingLiteralCellMutationAtNow = (
    request: EngineExistingLiteralCellMutationRef,
  ): EngineExistingNumericCellMutationResult | null => {
    if (typeof request.value === 'number') {
      return applyExistingNumericCellMutationAtNow({
        sheetId: request.sheetId,
        row: request.row,
        col: request.col,
        cellIndex: request.cellIndex,
        value: request.value,
        ...(request.emitTracked === undefined ? {} : { emitTracked: request.emitTracked }),
      })
    }
    if (
      args.state.workbook.hasPivots() ||
      args.state.events.hasListeners() ||
      args.state.events.hasCellListeners() ||
      args.hasVolatileFormulas?.()
    ) {
      return null
    }
    const sheet = args.state.workbook.getSheetById(request.sheetId)
    const cellStore = args.state.workbook.cellStore
    const existingIndex = request.cellIndex
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      cellStore.sheetIds[existingIndex] !== request.sheetId ||
      cellStore.rows[existingIndex] !== request.row ||
      cellStore.cols[existingIndex] !== request.col ||
      !canFastPathLiteralOverwrite(existingIndex) ||
      hasDynamicFormulaDependents(existingIndex)
    ) {
      return null
    }
    if (
      hasTrackedDirectRangeDependents(request.sheetId, request.col) ||
      hasTrackedExactLookupDependents(request.sheetId, request.col) ||
      hasTrackedSortedLookupDependents(request.sheetId, request.col)
    ) {
      return null
    }
    return tryApplyFormulaLeafExistingLiteralMutation({
      existingIndex,
      formulaCellIndex: args.getSingleEntityDependent(makeCellEntity(existingIndex)),
      value: request.value,
      hasTrackedEventListeners: request.emitTracked !== false && args.state.events.hasTrackedListeners(),
    })
  }

  return { tryApplySingleExistingDirectLiteralMutation, applyExistingNumericCellMutationAtNow, applyExistingLiteralCellMutationAtNow }
}
