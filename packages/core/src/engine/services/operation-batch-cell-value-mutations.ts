import { parseCellAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import type { CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { emptyValue, literalToValue, writeLiteralToCellStore } from '../../engine-value-utils.js'
import type { OpOrder } from '../../replica-state.js'
import type { PreparedCellAddress } from '../runtime-state.js'
import { withOptionalLookupStringIds } from './direct-lookup-helpers.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import { cellTouchesOperationPivotSource } from './operation-pivot-source-helpers.js'
import type { ExactLookupImpactCaches } from './operation-lookup-dirty-markers.js'
import type { CreateEngineOperationServiceArgs, MutationSource } from './operation-service-types.js'

type BatchSetCellValueOp = Extract<EngineOp, { kind: 'setCellValue' }>
type BatchClearCellOp = Extract<EngineOp, { kind: 'clearCell' }>

interface OperationBatchPreparedCells {
  readonly getExistingCellIndex: (sheetName: string, address: string, preparedCellAddress: PreparedCellAddress | null) => number | undefined
  readonly ensureCellTracked: (sheetName: string, address: string, preparedCellAddress: PreparedCellAddress | null) => number
}

interface BatchCellValueMutationCounts {
  readonly changedInputCount: number
  readonly formulaChangedCount: number
  readonly explicitChangedCount: number
  readonly topologyChanged: boolean
  readonly refreshAllPivots: boolean
}

interface BatchCellValueMutationBaseArgs extends BatchCellValueMutationCounts {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly order: OpOrder
  readonly source: MutationSource
  readonly isRestore: boolean
  readonly preparedCellAddress: PreparedCellAddress | null
  readonly preparedCells: OperationBatchPreparedCells
  readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
  readonly exactLookupImpactCaches: ExactLookupImpactCaches
  readonly setEntityVersionForOp: (op: EngineOp, order: OpOrder) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly readCellValueForLookup: (cellIndex: number | undefined) => { readonly value: CellValue; readonly stringId: number | undefined }
  readonly rebindValueSensitiveFormulaDependents: (cellIndex: number, counts: BatchCellValueMutationCounts) => BatchCellValueMutationCounts
  readonly refreshDependentRangesAndRebindFormulaDependents: (cellIndex: number, formulaChangedCount: number) => number
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
  readonly clearLookupImpactCaches: () => void
  readonly pruneCellIfOrphaned: (cellIndex: number) => void
}

interface ApplyBatchSetCellValueOpArgs extends BatchCellValueMutationBaseArgs {
  readonly op: BatchSetCellValueOp
  readonly isNullLiteralWriteNoOp: (cellIndex: number) => boolean
  readonly lookupHandledInputCellIndices: number[]
}

interface ApplyBatchClearCellOpArgs extends BatchCellValueMutationBaseArgs {
  readonly op: BatchClearCellOp
  readonly isClearCellNoOp: (cellIndex: number) => boolean
  readonly normalizeHistoryDependencyPlaceholder: (cellIndex: number, source: MutationSource) => void
}

export function applyBatchSetCellValueOp(request: ApplyBatchSetCellValueOpArgs): BatchCellValueMutationCounts {
  const args = request.serviceArgs
  const { op, preparedCellAddress } = request
  const existingIndex = request.preparedCells.getExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
  const parsedAddress = preparedCellAddress ?? parseCellAddress(op.address, op.sheetName)
  const sheet = args.state.workbook.getSheet(op.sheetName)
  const sheetId = sheet?.id
  const hasExactLookupDependents = sheetId !== undefined ? request.hasTrackedExactLookupDependents(sheetId, parsedAddress.col) : false
  const hasSortedLookupDependents = sheetId !== undefined ? request.hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) : false
  const hasAggregateDependents = sheetId !== undefined ? request.hasTrackedDirectRangeDependents(sheetId, parsedAddress.col) : false
  let changedInputCount = request.changedInputCount
  let formulaChangedCount = request.formulaChangedCount
  let explicitChangedCount = request.explicitChangedCount
  let topologyChanged = request.topologyChanged
  let refreshAllPivots = request.refreshAllPivots

  if (
    !request.isRestore &&
    cellTouchesOperationPivotSource({
      workbook: args.state.workbook,
      sheetName: op.sheetName,
      row: parsedAddress.row,
      col: parsedAddress.col,
    })
  ) {
    refreshAllPivots = true
  }
  const needsLookupValueRead = hasExactLookupDependents || hasSortedLookupDependents || hasAggregateDependents
  const prior = request.readCellValueForLookup(existingIndex)
  if (!request.isRestore) {
    if (op.value === null && (existingIndex === undefined || request.isNullLiteralWriteNoOp(existingIndex))) {
      return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, refreshAllPivots }
    }
    if (existingIndex !== undefined) {
      changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
    }
  }
  const cellIndex = request.preparedCells.ensureCellTracked(op.sheetName, op.address, preparedCellAddress)
  if (!request.isRestore) {
    changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
    const removedFormula = args.removeFormula(cellIndex)
    if (removedFormula) {
      args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
      request.clearLookupImpactCaches()
      formulaChangedCount = request.refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
    }
    topologyChanged = removedFormula || topologyChanged
  }
  writeLiteralToCellStore(args.state.workbook.cellStore, cellIndex, op.value, args.state.strings)
  if (op.value === null) {
    args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.AuthoredBlank
  }
  args.state.workbook.notifyCellValueWritten(cellIndex)
  if (!request.isRestore) {
    ;({ changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, refreshAllPivots } =
      request.rebindValueSensitiveFormulaDependents(cellIndex, {
        changedInputCount,
        formulaChangedCount,
        explicitChangedCount,
        topologyChanged,
        refreshAllPivots,
      }))
  }
  if (needsLookupValueRead) {
    const formulaChangedCountBeforeLookupNotes = formulaChangedCount
    const newValue = literalToValue(op.value, args.state.strings)
    const newStringId = typeof op.value === 'string' ? args.state.workbook.cellStore.stringIds[cellIndex] : undefined
    if (!request.isRestore) {
      const directDependentsHandled = request.markPostRecalcDirectFormulaDependents(
        cellIndex,
        request.postRecalcDirectFormulaIndices,
        prior.value,
        newValue,
      )
      if (!directDependentsHandled) {
        request.markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, request.postRecalcDirectFormulaIndices)
      }
    }
    if (hasExactLookupDependents || hasAggregateDependents) {
      const exactLookupRequest = withOptionalLookupStringIds({
        sheetName: op.sheetName,
        row: parsedAddress.row,
        col: parsedAddress.col,
        oldValue: prior.value,
        newValue,
        oldStringId: prior.stringId,
        newStringId,
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
        sheetName: op.sheetName,
        row: parsedAddress.row,
        col: parsedAddress.col,
        oldValue: prior.value,
        newValue,
        oldStringId: prior.stringId,
        newStringId,
      })
      formulaChangedCount = request.noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
    }
    if (
      !hasAggregateDependents &&
      (hasExactLookupDependents || hasSortedLookupDependents) &&
      formulaChangedCount === formulaChangedCountBeforeLookupNotes
    ) {
      request.lookupHandledInputCellIndices.push(cellIndex)
    }
  } else if (!request.isRestore) {
    const newValue = literalToValue(op.value, args.state.strings)
    const directDependentsHandled = request.markPostRecalcDirectFormulaDependents(
      cellIndex,
      request.postRecalcDirectFormulaIndices,
      prior.value,
      newValue,
    )
    if (!directDependentsHandled) {
      request.markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, request.postRecalcDirectFormulaIndices)
    }
  }
  args.state.workbook.cellStore.flags[cellIndex] =
    (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
    ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
  if (!request.isRestore && op.value === null) {
    request.pruneCellIfOrphaned(cellIndex)
  }
  changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
  if (!request.isRestore) {
    explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
    request.setEntityVersionForOp(op, request.order)
  }
  return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, refreshAllPivots }
}

export function applyBatchClearCellOp(request: ApplyBatchClearCellOpArgs): BatchCellValueMutationCounts {
  const args = request.serviceArgs
  const { op, preparedCellAddress } = request
  const cellIndex = request.preparedCells.getExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
  const parsedAddress = preparedCellAddress ?? parseCellAddress(op.address, op.sheetName)
  const sheet = args.state.workbook.getSheet(op.sheetName)
  const sheetId = sheet?.id
  const hasExactLookupDependents = sheetId !== undefined ? request.hasTrackedExactLookupDependents(sheetId, parsedAddress.col) : false
  const hasSortedLookupDependents = sheetId !== undefined ? request.hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) : false
  const hasAggregateDependents = sheetId !== undefined ? request.hasTrackedDirectRangeDependents(sheetId, parsedAddress.col) : false
  const needsLookupValueRead = hasExactLookupDependents || hasSortedLookupDependents || hasAggregateDependents
  let changedInputCount = request.changedInputCount
  let formulaChangedCount = request.formulaChangedCount
  let explicitChangedCount = request.explicitChangedCount
  let topologyChanged = request.topologyChanged
  let refreshAllPivots = request.refreshAllPivots

  if (
    !request.isRestore &&
    cellTouchesOperationPivotSource({
      workbook: args.state.workbook,
      sheetName: op.sheetName,
      row: parsedAddress.row,
      col: parsedAddress.col,
    })
  ) {
    refreshAllPivots = true
  }
  const prior = request.readCellValueForLookup(cellIndex)
  if (cellIndex === undefined) {
    request.setEntityVersionForOp(op, request.order)
    return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, refreshAllPivots }
  }
  if (request.isClearCellNoOp(cellIndex)) {
    return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, refreshAllPivots }
  }
  changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(cellIndex), changedInputCount)
  changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
  const removedFormula = args.removeFormula(cellIndex)
  if (removedFormula) {
    args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
    request.clearLookupImpactCaches()
    formulaChangedCount = request.refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
  }
  topologyChanged = removedFormula || topologyChanged
  args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
  args.state.workbook.notifyCellValueWritten(cellIndex)
  if (!request.isRestore) {
    ;({ changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, refreshAllPivots } =
      request.rebindValueSensitiveFormulaDependents(cellIndex, {
        changedInputCount,
        formulaChangedCount,
        explicitChangedCount,
        topologyChanged,
        refreshAllPivots,
      }))
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
  if (needsLookupValueRead) {
    if (hasExactLookupDependents || hasAggregateDependents) {
      const exactLookupRequest = withOptionalLookupStringIds({
        sheetName: op.sheetName,
        row: parsedAddress.row,
        col: parsedAddress.col,
        oldValue: prior.value,
        newValue: emptyValue(),
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
        sheetName: op.sheetName,
        row: parsedAddress.row,
        col: parsedAddress.col,
        oldValue: prior.value,
        newValue: emptyValue(),
        oldStringId: prior.stringId,
        newStringId: undefined,
      })
      formulaChangedCount = request.noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
    }
  }
  args.state.workbook.cellStore.flags[cellIndex] =
    (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
    ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput | CellFlags.AuthoredBlank)
  request.normalizeHistoryDependencyPlaceholder(cellIndex, request.source)
  request.pruneCellIfOrphaned(cellIndex)
  changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
  explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
  request.setEntityVersionForOp(op, request.order)
  return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, refreshAllPivots }
}
