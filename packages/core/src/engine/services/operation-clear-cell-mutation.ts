import { formatAddress } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { emptyValue } from '../../engine-value-utils.js'
import type { OpOrder } from '../../replica-state.js'
import type { CreateEngineOperationServiceArgs, MutationSource } from './operation-service-types.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import { withOptionalLookupStringIds } from './direct-lookup-helpers.js'
import type { OperationTrackedColumnDependencyFlags } from './operation-column-dependency-tracker.js'
import type { ExactLookupImpactCaches } from './operation-lookup-dirty-markers.js'
import type { OperationLookupAccess } from './operation-lookup-access.js'

type ClearCellMutation = Extract<EngineCellMutationRef['mutation'], { kind: 'clearCell' }>
type OperationCellMutationSource = Exclude<MutationSource, 'remote'>

interface ClearCellMutationCounts {
  readonly changedInputCount: number
  readonly formulaChangedCount: number
  readonly explicitChangedCount: number
  readonly topologyChanged: boolean
}

interface ApplyClearCellMutationArgs extends ClearCellMutationCounts {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly sheetId: number
  readonly mutation: ClearCellMutation
  readonly existingIndex: number | undefined
  readonly source: OperationCellMutationSource
  readonly isRestore: boolean
  readonly trackExplicitChanges: boolean
  readonly order: OpOrder | undefined
  readonly dependencyFlags: OperationTrackedColumnDependencyFlags
  readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
  readonly exactLookupImpactCaches: ExactLookupImpactCaches
  readonly resolveSheetName: (sheetId: number) => string
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly isClearCellNoOp: (cellIndex: number) => boolean
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly readCellValueForLookup: OperationLookupAccess['readCellValueForLookup']
  readonly rebindValueSensitiveFormulaDependents: (cellIndex: number, counts: ClearCellMutationCounts) => ClearCellMutationCounts
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
  readonly clearTrackedColumnDependencyFlagCache: () => void
  readonly pruneCellIfOrphaned: (cellIndex: number) => void
  readonly normalizeHistoryDependencyPlaceholder: (cellIndex: number, source: MutationSource) => void
}

export function applyClearCellMutation(request: ApplyClearCellMutationArgs): ClearCellMutationCounts {
  const args = request.serviceArgs
  const { existingIndex, isRestore, mutation, sheetId } = request
  const { hasAggregateDependents, hasExactLookupDependents, hasSortedLookupDependents, needsLookupValueRead } = request.dependencyFlags
  const prior = request.readCellValueForLookup(existingIndex)
  let changedInputCount = request.changedInputCount
  let formulaChangedCount = request.formulaChangedCount
  let explicitChangedCount = request.explicitChangedCount
  let topologyChanged = request.topologyChanged

  const markReplicaVersion = (): void => {
    if (!isRestore && args.state.trackReplicaVersions) {
      request.setCellEntityVersion(request.resolveSheetName(sheetId), formatAddress(mutation.row, mutation.col), request.order!)
    }
  }

  const markDirectDependents = (cellIndex: number): void => {
    if (isRestore) {
      return
    }
    const nextValue = emptyValue()
    const directDependentsHandled = request.markPostRecalcDirectFormulaDependents(
      cellIndex,
      request.postRecalcDirectFormulaIndices,
      prior.value,
      nextValue,
    )
    if (!directDependentsHandled) {
      request.markDirectScalarDeltaClosure(cellIndex, prior.value, nextValue, request.postRecalcDirectFormulaIndices)
    }
  }

  const markLookupAndAggregateDependents = (cellIndex: number): void => {
    if (!needsLookupValueRead) {
      return
    }
    const sheetName = request.resolveSheetName(sheetId)
    const nextValue = emptyValue()
    if (hasExactLookupDependents || hasAggregateDependents) {
      const exactLookupRequest = withOptionalLookupStringIds({
        sheetName,
        row: mutation.row,
        col: mutation.col,
        oldValue: prior.value,
        newValue: nextValue,
        oldStringId: prior.stringId,
        newStringId: undefined,
        inputCellIndex: cellIndex,
      })
      if (hasExactLookupDependents) {
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
        oldValue: prior.value,
        newValue: nextValue,
        oldStringId: prior.stringId,
        newStringId: undefined,
      })
      formulaChangedCount = request.noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
    }
  }

  const markChanged = (cellIndex: number): void => {
    changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
    if (request.trackExplicitChanges) {
      explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
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

  if (existingIndex !== undefined && request.isClearCellNoOp(existingIndex)) {
    return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged }
  }

  if (existingIndex !== undefined && request.canFastPathLiteralOverwrite(existingIndex)) {
    args.state.workbook.cellStore.setValue(existingIndex, emptyValue())
    args.state.workbook.notifyCellValueWritten(existingIndex)
    if (!isRestore) {
      rebindValueSensitiveDependents(existingIndex)
    }
    markDirectDependents(existingIndex)
    markLookupAndAggregateDependents(existingIndex)
    markChanged(existingIndex)
    markReplicaVersion()
    return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged }
  }

  if (existingIndex === undefined) {
    markReplicaVersion()
    return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged }
  }

  changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
  changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(existingIndex), changedInputCount)
  const removedFormula = args.removeFormula(existingIndex)
  topologyChanged = removedFormula || topologyChanged
  if (removedFormula) {
    args.invalidateAggregateColumn({ sheetName: request.resolveSheetName(sheetId), col: mutation.col })
    request.clearTrackedColumnDependencyFlagCache()
  }
  args.state.workbook.cellStore.setValue(existingIndex, emptyValue())
  args.state.workbook.notifyCellValueWritten(existingIndex)
  if (!isRestore) {
    rebindValueSensitiveDependents(existingIndex)
  }
  markDirectDependents(existingIndex)
  markLookupAndAggregateDependents(existingIndex)
  args.state.workbook.cellStore.flags[existingIndex] =
    (args.state.workbook.cellStore.flags[existingIndex] ?? 0) &
    ~(CellFlags.AuthoredBlank | CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
  request.normalizeHistoryDependencyPlaceholder(existingIndex, request.source)
  if (!isRestore) {
    request.pruneCellIfOrphaned(existingIndex)
  }
  markChanged(existingIndex)
  markReplicaVersion()
  return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged }
}
