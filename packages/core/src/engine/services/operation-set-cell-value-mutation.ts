import { formatAddress } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { literalToValue, writeLiteralToCellStore } from '../../engine-value-utils.js'
import type { OpOrder } from '../../replica-state.js'
import { exactLookupLiteralNumericValue, withOptionalLookupStringIds } from './direct-lookup-helpers.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import { directScalarLiteralNumericValue } from './direct-scalar-helpers.js'
import type { OperationTrackedColumnDependencyFlags } from './operation-column-dependency-tracker.js'
import type { OperationLookupAccess } from './operation-lookup-access.js'
import type { ExactLookupImpactCaches } from './operation-lookup-dirty-markers.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

type SetCellValueMutation = Extract<EngineCellMutationRef['mutation'], { kind: 'setCellValue' }>

interface SetCellValueMutationCounts {
  readonly changedInputCount: number
  readonly formulaChangedCount: number
  readonly explicitChangedCount: number
  readonly topologyChanged: boolean
}

interface ApplySetCellValueMutationArgs extends SetCellValueMutationCounts {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly sheetId: number
  readonly sheetName: string
  readonly mutation: SetCellValueMutation
  readonly existingIndex: number | undefined
  readonly isRestore: boolean
  readonly trackExplicitChanges: boolean
  readonly order: OpOrder | undefined
  readonly dependencyFlags: OperationTrackedColumnDependencyFlags
  readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
  readonly exactLookupImpactCaches: ExactLookupImpactCaches
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly isNullLiteralWriteNoOp: (cellIndex: number) => boolean
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly readCellValueForLookup: OperationLookupAccess['readCellValueForLookup']
  readonly readApproximateNumericValueForLookup: OperationLookupAccess['readApproximateNumericValueForLookup']
  readonly readExactNumericValueForLookup: OperationLookupAccess['readExactNumericValueForLookup']
  readonly canSkipExactLookupNumericColumnWrite: (sheetId: number, col: number, row: number, oldValue: number, newValue: number) => boolean
  readonly canSkipApproximateLookupNumericColumnWrite: (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldValue: number,
    newValue: number,
  ) => boolean
  readonly rebindValueSensitiveFormulaDependents: (cellIndex: number, counts: SetCellValueMutationCounts) => SetCellValueMutationCounts
  readonly markPostRecalcDirectFormulaDependents: (
    cellIndex: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
    oldValue?: CellValue,
    newValue?: CellValue,
  ) => boolean
  readonly markDirectScalarDeltaClosure: (
    rootCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ) => void
  readonly markPostRecalcDirectScalarNumericDependents: (
    cellIndex: number,
    oldNumber: number,
    newNumber: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
    exactLookupValue?: number,
    approximateLookupValue?: number,
  ) => boolean
  readonly markPostRecalcDirectLookupCurrentDependentsFromNumeric: (
    cellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ) => boolean
  readonly directScalarCellNumericValue: (cellIndex: number) => number | undefined
  readonly noteExactLookupLiteralWriteWhenDirty: (
    request: {
      readonly sheetName: string
      readonly row: number
      readonly col: number
      readonly oldValue: CellValue
      readonly newValue: CellValue
      readonly oldStringId?: number
      readonly newStringId?: number
      readonly inputCellIndex?: number
    },
    formulaChangedCount: number,
    caches: ExactLookupImpactCaches,
  ) => number
  readonly noteSortedLookupLiteralWriteWhenDirty: (
    request: {
      readonly sheetName: string
      readonly row: number
      readonly col: number
      readonly oldValue: CellValue
      readonly newValue: CellValue
      readonly oldStringId?: number
      readonly newStringId?: number
    },
    formulaChangedCount: number,
  ) => number
  readonly markAffectedDirectRangeDependents: (
    request: {
      readonly sheetName: string
      readonly row: number
      readonly col: number
      readonly oldValue: CellValue
      readonly newValue: CellValue
      readonly oldStringId?: number
      readonly newStringId?: number
      readonly inputCellIndex?: number
    },
    formulaChangedCount: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ) => number
  readonly queueHandledLookupInvalidation: (sheetId: number, sheetName: string, col: number, exact: boolean, sorted: boolean) => void
  readonly noteHandledLookupInputCellIndex: (cellIndex: number) => void
  readonly clearTrackedColumnDependencyFlagCache: () => void
  readonly pruneCellIfOrphaned: (cellIndex: number) => void
}

export function applySetCellValueMutation(request: ApplySetCellValueMutationArgs): SetCellValueMutationCounts {
  const args = request.serviceArgs
  const { existingIndex, isRestore, mutation, sheetId, sheetName } = request
  const { hasAggregateDependents, hasExactLookupDependents, hasSortedLookupDependents } = request.dependencyFlags
  let changedInputCount = request.changedInputCount
  let formulaChangedCount = request.formulaChangedCount
  let explicitChangedCount = request.explicitChangedCount
  let topologyChanged = request.topologyChanged

  const markReplicaVersion = (): void => {
    if (!isRestore && args.state.trackReplicaVersions) {
      request.setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), request.order!)
    }
  }

  const rebindValueSensitiveDependents = (cellIndex: number): void => {
    const rebound = request.rebindValueSensitiveFormulaDependents(cellIndex, {
      changedInputCount,
      formulaChangedCount,
      explicitChangedCount,
      topologyChanged,
    })
    changedInputCount = rebound.changedInputCount
    formulaChangedCount = rebound.formulaChangedCount
    explicitChangedCount = rebound.explicitChangedCount
    topologyChanged = rebound.topologyChanged
  }

  if (mutation.value === null && !isRestore && (existingIndex === undefined || request.isNullLiteralWriteNoOp(existingIndex))) {
    return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged }
  }

  const canFastOverwriteExisting = existingIndex !== undefined && request.canFastPathLiteralOverwrite(existingIndex)
  const needsDirectLookupNumericValue = canFastOverwriteExisting
  const oldExactLookupNumber =
    canFastOverwriteExisting && hasExactLookupDependents ? request.readExactNumericValueForLookup(existingIndex) : undefined
  const newExactLookupNumber =
    hasExactLookupDependents || needsDirectLookupNumericValue ? exactLookupLiteralNumericValue(mutation.value) : undefined
  const oldApproximateLookupNumber =
    canFastOverwriteExisting && hasSortedLookupDependents ? request.readApproximateNumericValueForLookup(existingIndex) : undefined
  const newApproximateLookupNumber =
    hasSortedLookupDependents || needsDirectLookupNumericValue ? directScalarLiteralNumericValue(mutation.value) : undefined
  const exactLookupDependentsHandled =
    !isRestore &&
    hasExactLookupDependents &&
    !hasAggregateDependents &&
    oldExactLookupNumber !== undefined &&
    newExactLookupNumber !== undefined &&
    request.canSkipExactLookupNumericColumnWrite(sheetId, mutation.col, mutation.row, oldExactLookupNumber, newExactLookupNumber)
  const sortedLookupDependentsHandled =
    !isRestore &&
    hasSortedLookupDependents &&
    oldApproximateLookupNumber !== undefined &&
    newApproximateLookupNumber !== undefined &&
    request.canSkipApproximateLookupNumericColumnWrite(
      sheetId,
      sheetName,
      mutation.col,
      mutation.row,
      oldApproximateLookupNumber,
      newApproximateLookupNumber,
    )
  const needsLookupValueRead =
    hasAggregateDependents ||
    (hasExactLookupDependents && !exactLookupDependentsHandled) ||
    (hasSortedLookupDependents && !sortedLookupDependentsHandled)
  const needsLookupOwnerInvalidation =
    (hasExactLookupDependents && exactLookupDependentsHandled) || (hasSortedLookupDependents && sortedLookupDependentsHandled)
  let directDependentsHandled = false
  if (!isRestore && canFastOverwriteExisting) {
    const oldNumber = request.directScalarCellNumericValue(existingIndex)
    const newNumber = directScalarLiteralNumericValue(mutation.value)
    if (oldNumber !== undefined && newNumber !== undefined) {
      directDependentsHandled = request.markPostRecalcDirectScalarNumericDependents(
        existingIndex,
        oldNumber,
        newNumber,
        request.postRecalcDirectFormulaIndices,
        newExactLookupNumber,
        newApproximateLookupNumber,
      )
    }
  }
  const canUseDirectLookupCurrent =
    !isRestore &&
    canFastOverwriteExisting &&
    (newExactLookupNumber !== undefined || newApproximateLookupNumber !== undefined) &&
    !needsLookupValueRead &&
    !directDependentsHandled
  if (canUseDirectLookupCurrent) {
    directDependentsHandled = request.markPostRecalcDirectLookupCurrentDependentsFromNumeric(
      existingIndex,
      newExactLookupNumber,
      newApproximateLookupNumber,
      request.postRecalcDirectFormulaIndices,
    )
  }
  let prior = needsLookupValueRead || !directDependentsHandled ? request.readCellValueForLookup(existingIndex) : undefined

  const markDirectFormulaDependents = (cellIndex: number, newValue: CellValue | undefined): void => {
    if (isRestore || directDependentsHandled || newValue === undefined) {
      return
    }
    prior ??= request.readCellValueForLookup(existingIndex)
    const genericDirectDependentsHandled = request.markPostRecalcDirectFormulaDependents(
      cellIndex,
      request.postRecalcDirectFormulaIndices,
      prior.value,
      newValue,
    )
    if (!genericDirectDependentsHandled) {
      request.markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, request.postRecalcDirectFormulaIndices)
    }
  }

  const markLookupDependents = (cellIndex: number, newValue: CellValue | undefined): void => {
    if (!needsLookupValueRead) {
      return
    }
    const newStringId = typeof mutation.value === 'string' ? args.state.workbook.cellStore.stringIds[cellIndex] : undefined
    const priorLookup = prior ?? request.readCellValueForLookup(existingIndex)
    const newLookupValue = newValue ?? literalToValue(mutation.value, args.state.strings)
    if (hasExactLookupDependents || hasAggregateDependents) {
      const exactLookupRequest = withOptionalLookupStringIds({
        sheetName,
        row: mutation.row,
        col: mutation.col,
        oldValue: priorLookup.value,
        newValue: newLookupValue,
        oldStringId: priorLookup.stringId,
        newStringId,
        inputCellIndex: cellIndex,
      })
      if (hasExactLookupDependents && !exactLookupDependentsHandled) {
        formulaChangedCount = request.noteExactLookupLiteralWriteWhenDirty(
          exactLookupRequest,
          formulaChangedCount,
          request.exactLookupImpactCaches,
        )
      }
      if (hasAggregateDependents) {
        args.noteAggregateLiteralWrite({
          sheetName: exactLookupRequest.sheetName,
          row: exactLookupRequest.row,
          col: exactLookupRequest.col,
          oldValue: exactLookupRequest.oldValue,
          newValue: exactLookupRequest.newValue,
        })
        formulaChangedCount = request.markAffectedDirectRangeDependents(
          exactLookupRequest,
          formulaChangedCount,
          request.postRecalcDirectFormulaIndices,
        )
      }
    }
    if (hasSortedLookupDependents) {
      const sortedLookupRequest = withOptionalLookupStringIds({
        sheetName,
        row: mutation.row,
        col: mutation.col,
        oldValue: priorLookup.value,
        newValue: newLookupValue,
        oldStringId: priorLookup.stringId,
        newStringId,
      })
      if (!sortedLookupDependentsHandled) {
        formulaChangedCount = request.noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
      }
    }
  }

  if (canFastOverwriteExisting) {
    writeLiteralToCellStore(args.state.workbook.cellStore, existingIndex, mutation.value, args.state.strings)
    args.state.workbook.notifyCellValueWritten(existingIndex)
    if (!isRestore) {
      rebindValueSensitiveDependents(existingIndex)
    }
    if (needsLookupOwnerInvalidation) {
      request.queueHandledLookupInvalidation(
        sheetId,
        sheetName,
        mutation.col,
        hasExactLookupDependents && exactLookupDependentsHandled,
        hasSortedLookupDependents && sortedLookupDependentsHandled,
      )
      if (!needsLookupValueRead) {
        request.noteHandledLookupInputCellIndex(existingIndex)
      }
    }
    const newValue = needsLookupValueRead || !directDependentsHandled ? literalToValue(mutation.value, args.state.strings) : undefined
    markDirectFormulaDependents(existingIndex, newValue)
    markLookupDependents(existingIndex, newValue)
    changedInputCount = args.markInputChanged(existingIndex, changedInputCount)
    if (request.trackExplicitChanges) {
      explicitChangedCount = args.markExplicitChanged(existingIndex, explicitChangedCount)
    }
    markReplicaVersion()
    return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged }
  }

  if (existingIndex !== undefined) {
    changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
  }
  const cellIndex = args.state.workbook.ensureCellAt(sheetId, mutation.row, mutation.col).cellIndex
  if (!isRestore && existingIndex !== undefined) {
    changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
    const removedFormula = args.removeFormula(cellIndex)
    topologyChanged = removedFormula || topologyChanged
    if (removedFormula) {
      args.invalidateAggregateColumn({ sheetName, col: mutation.col })
      request.clearTrackedColumnDependencyFlagCache()
    }
  }
  writeLiteralToCellStore(args.state.workbook.cellStore, cellIndex, mutation.value, args.state.strings)
  args.state.workbook.notifyCellValueWritten(cellIndex)
  if (!isRestore) {
    rebindValueSensitiveDependents(cellIndex)
  }
  const newValue = needsLookupValueRead || !directDependentsHandled ? literalToValue(mutation.value, args.state.strings) : undefined
  markDirectFormulaDependents(cellIndex, newValue)
  markLookupDependents(cellIndex, newValue)
  args.state.workbook.cellStore.flags[cellIndex] =
    (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
    ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
  if (!isRestore && mutation.value === null) {
    request.pruneCellIfOrphaned(cellIndex)
  }
  changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
  if (request.trackExplicitChanges) {
    explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
  }
  markReplicaVersion()
  return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged }
}
