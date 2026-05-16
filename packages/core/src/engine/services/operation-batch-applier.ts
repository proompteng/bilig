import { parseCellAddress } from '@bilig/formula'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import { FormulaMode, type CellRangeRef, type CellValue } from '@bilig/protocol'
import { batchOpOrder, compareOpOrder, markBatchApplied, type OpOrder } from '../../replica-state.js'
import { CellFlags } from '../../cell-store.js'
import { emptyValue, literalToValue, writeLiteralToCellStore } from '../../engine-value-utils.js'
import { calculationSettingsEqual, normalizeWorkbookCalculationSettings, tableDependencyKey } from '../../engine-metadata-utils.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { PreparedCellAddress } from '../runtime-state.js'
import { withOptionalLookupStringIds } from './direct-lookup-helpers.js'
import { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import { assertNever } from './operation-change-helpers.js'
import { isScalarOnlyDefinedNameValue } from './defined-name-value-helpers.js'
import { shouldApplyOp as shouldApplyReplicaOp, type OperationReplicaVersionWriter } from './operation-replica-helpers.js'
import { assertProtectionAllowsOp as assertProtectionAllowsProtectedOp } from './operation-protection-helpers.js'
import { cellTouchesOperationPivotSource } from './operation-pivot-source-helpers.js'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import { applyOperationStructuralMetadataOp } from './operation-structural-metadata-ops.js'
import { createOperationPreparedCellTracker } from './operation-cell-address-resolver.js'
import type { OperationLookupAccess } from './operation-lookup-access.js'
import type { ExactLookupImpactCaches, OperationLookupDirtyMarkerService } from './operation-lookup-dirty-markers.js'
import type { OperationColumnDependencyTrackerService } from './operation-column-dependency-tracker.js'
import type { OperationDirectRangeDependentService } from './operation-direct-range-dependents.js'
import { finalizeOperationRecalcAndEvents } from './operation-recalc-finalizer.js'
import type { CreateEngineOperationServiceArgs, MutationSource } from './operation-service-types.js'

type OperationBatchDirectFormulaCallbacks = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['directFormulaCallbacks']
type OperationBatchDirtyTraversalSkip = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['canSkipDirtyTraversalForChangedInputs']
type OperationBatchChangedInputsNeedRegionQueryIndices = Parameters<
  typeof finalizeOperationRecalcAndEvents
>[0]['changedInputsNeedRegionQueryIndices']
type OperationBatchCycleInputMarker = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['markCycleMemberInputsChanged']
type OperationBatchDerivedOp<K extends EngineOp['kind']> = Extract<EngineOp, { kind: K }>

interface CreateOperationBatchApplierArgs {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly replicaVersionWriter: OperationReplicaVersionWriter
  readonly isNullLiteralWriteNoOp: (cellIndex: number) => boolean
  readonly isClearCellNoOp: (cellIndex: number) => boolean
  readonly readCellValueForLookup: OperationLookupAccess['readCellValueForLookup']
  readonly readExactNumericValueForLookup: OperationLookupAccess['readExactNumericValueForLookup']
  readonly hasTrackedExactLookupDependents: OperationColumnDependencyTrackerService['hasTrackedExactLookupDependents']
  readonly hasTrackedSortedLookupDependents: OperationColumnDependencyTrackerService['hasTrackedSortedLookupDependents']
  readonly hasTrackedDirectRangeDependents: OperationColumnDependencyTrackerService['hasTrackedDirectRangeDependents']
  readonly noteExactLookupLiteralWriteWhenDirty: OperationLookupDirtyMarkerService['noteExactLookupLiteralWriteWhenDirty']
  readonly noteSortedLookupLiteralWriteWhenDirty: OperationLookupDirtyMarkerService['noteSortedLookupLiteralWriteWhenDirty']
  readonly markAffectedDirectRangeDependents: OperationDirectRangeDependentService['markAffectedDirectRangeDependents']
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
  readonly collectAffectedDirectRangeDependents: OperationDirectRangeDependentService['collectAffectedDirectRangeDependents']
  readonly tryApplyFormulaReplacementAsDirectScalarDeltaRoot: (request: {
    readonly cellIndex: number
    readonly oldNumber: number | undefined
    readonly changedTopology: boolean
    readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
    readonly postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts
  }) => boolean
  readonly rebindDynamicFormulaDependents: (cellIndex: number, formulaChangedCount: number) => number
  readonly refreshDependentRangesAndRebindFormulaDependents: (cellIndex: number, formulaChangedCount: number) => number
  readonly pruneCellIfOrphaned: (cellIndex: number) => void
  readonly normalizeHistoryDependencyPlaceholder: (cellIndex: number, source: MutationSource) => void
  readonly markCycleMemberInputsChanged: OperationBatchCycleInputMarker
  readonly hasCycleMembersNow: () => boolean
  readonly canSkipDirtyTraversalForChangedInputs: OperationBatchDirtyTraversalSkip
  readonly changedInputsNeedRegionQueryIndices: OperationBatchChangedInputsNeedRegionQueryIndices
  readonly directFormulaCallbacks: OperationBatchDirectFormulaCallbacks
  readonly applySpillRangeOp: (op: OperationBatchDerivedOp<'upsertSpillRange' | 'deleteSpillRange'>, order: OpOrder) => number[]
  readonly applyPivotUpsertOp: (op: OperationBatchDerivedOp<'upsertPivotTable'>, order: OpOrder) => number[]
  readonly applyPivotDeleteOp: (op: OperationBatchDerivedOp<'deletePivotTable'>, order: OpOrder) => number[]
}

export function createOperationBatchApplier(input: CreateOperationBatchApplierArgs) {
  const {
    serviceArgs: args,
    emitBatch,
    isNullLiteralWriteNoOp,
    isClearCellNoOp,
    readCellValueForLookup,
    readExactNumericValueForLookup,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    noteExactLookupLiteralWriteWhenDirty,
    noteSortedLookupLiteralWriteWhenDirty,
    markAffectedDirectRangeDependents,
    markPostRecalcDirectFormulaDependents,
    markDirectScalarDeltaClosure,
    collectAffectedDirectRangeDependents,
    tryApplyFormulaReplacementAsDirectScalarDeltaRoot,
    rebindDynamicFormulaDependents,
    refreshDependentRangesAndRebindFormulaDependents,
    pruneCellIfOrphaned,
    normalizeHistoryDependencyPlaceholder,
    markCycleMemberInputsChanged,
    hasCycleMembersNow,
    canSkipDirtyTraversalForChangedInputs,
    changedInputsNeedRegionQueryIndices,
    directFormulaCallbacks: {
      applyDirectFormulaCurrentResult,
      applyDirectFormulaNumericDelta,
      applyDirectScalarCurrentValue,
      tryApplyDirectScalarDeltas,
      tryApplyDirectFormulaDeltas,
      countPostRecalcDirectFormulaMetric,
    },
    applySpillRangeOp,
    applyPivotUpsertOp,
    applyPivotDeleteOp,
  } = input
  const { setEntityVersionForOp, setSheetDeleteVersion, stores: replicaStores } = input.replicaVersionWriter

  return function applyBatchNow(
    batch: EngineOpBatch,
    source: MutationSource,
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
  ): void {
    if (preparedCellAddressesByOpIndex && preparedCellAddressesByOpIndex.length !== batch.ops.length) {
      throw new Error('Prepared cell addresses must align with batch operations')
    }
    const isRestore = source === 'restore'
    args.beginMutationCollection()
    let changedInputCount = 0
    let formulaChangedCount = 0
    let explicitChangedCount = 0
    let topologyChanged = false
    let sheetDeleted = false
    let structuralInvalidation = false
    let compileMs = 0
    const invalidatedRanges: CellRangeRef[] = []
    const invalidatedRows: { sheetName: string; startIndex: number; endIndex: number }[] = []
    const invalidatedColumns: { sheetName: string; startIndex: number; endIndex: number }[] = []
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
    const precomputedKernelSyncCellIndices: number[] = []
    let refreshAllPivots = false
    let appliedOps = 0
    const postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts = {
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    }
    const lookupHandledInputCellIndices: number[] = []
    const canSkipOrderChecks = source !== 'remote'
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    const exactLookupImpactCaches: ExactLookupImpactCaches = new Map()
    const clearLookupImpactCaches = (): void => {
      exactLookupImpactCaches.clear()
    }
    const rebindValueSensitiveFormulaDependents = (cellIndex: number): void => {
      const reboundCount = formulaChangedCount
      formulaChangedCount = rebindDynamicFormulaDependents(cellIndex, formulaChangedCount)
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
    }

    const reservedNewCells = potentialNewCells ?? args.estimatePotentialNewCells(batch.ops)
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const preparedCells = createOperationPreparedCellTracker({
      workbook: args.state.workbook,
      ensureCellTracked: args.ensureCellTracked,
    })

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      if (!isRestore && source !== 'undo' && source !== 'redo') {
        batch.ops.forEach((op) => {
          assertProtectionAllowsProtectedOp(args.state.workbook, op)
        })
      }
      batch.ops.forEach((op, opIndex) => {
        const order = batchOpOrder(batch, opIndex)
        const preparedCellAddress = preparedCellAddressesByOpIndex?.[opIndex] ?? null
        if (!canSkipOrderChecks && !shouldApplyReplicaOp(op, order, replicaStores)) {
          return
        }
        args.materializeDeferredStructuralFormulaSources()

        switch (op.kind) {
          case 'upsertWorkbook':
            args.state.workbook.workbookName = op.name
            setEntityVersionForOp(op, order)
            break
          case 'setWorkbookMetadata':
            args.state.workbook.setWorkbookProperty(op.key, op.value)
            setEntityVersionForOp(op, order)
            break
          case 'setCalculationSettings':
            const previousCalculationSettings = args.state.workbook.getCalculationSettings()
            const nextCalculationSettings = normalizeWorkbookCalculationSettings(op.settings, previousCalculationSettings)
            if (calculationSettingsEqual(previousCalculationSettings, nextCalculationSettings)) {
              break
            }
            args.state.workbook.setCalculationSettings(nextCalculationSettings)
            if (previousCalculationSettings.dateSystem !== nextCalculationSettings.dateSystem) {
              const reboundCount = formulaChangedCount
              formulaChangedCount = args.rebindFormulaCells([...args.state.formulas.keys()], formulaChangedCount)
              topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            } else {
              args.state.formulas.forEach((_formula, cellIndex) => {
                formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
              })
            }
            setEntityVersionForOp(op, order)
            break
          case 'setVolatileContext':
            args.state.workbook.setVolatileContext(op.context)
            setEntityVersionForOp(op, order)
            break
          case 'upsertSheet': {
            preparedCells.invalidateSheetName(op.name)
            args.state.workbook.createSheet(op.name, op.order, op.id)
            setEntityVersionForOp(op, order)
            const tombstone = replicaStores.sheetDeleteVersions.get(op.name)
            if (!tombstone || compareOpOrder(order, tombstone) > 0) {
              args.state.sheetDeleteVersions.delete(op.name)
            }
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindFormulasForSheet(op.name, formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            refreshAllPivots = true
            break
          }
          case 'renameSheet': {
            preparedCells.invalidateSheetName(op.oldName)
            preparedCells.invalidateSheetName(op.newName)
            const renamedSheet = args.state.workbook.renameSheet(op.oldName, op.newName)
            if (args.state.trackReplicaVersions) {
              replicaStores.entityVersions.set(`sheet:${op.oldName}`, order)
              replicaStores.entityVersions.set(`sheet:${op.newName}`, order)
            }
            setSheetDeleteVersion(op.oldName, order)
            const renamedTombstone = replicaStores.sheetDeleteVersions.get(op.newName)
            if (!renamedTombstone || compareOpOrder(order, renamedTombstone) > 0) {
              args.state.sheetDeleteVersions.delete(op.newName)
            }
            if (!renamedSheet) {
              break
            }
            const selection = args.getSelectionState()
            if (selection.sheetName === op.oldName) {
              args.setSelection(op.newName, selection.address ?? 'A1')
            }
            if (args.state.workbook.metadata.definedNames.size > 0) {
              args.rewriteDefinedNamesForSheetRename(op.oldName, op.newName)
            }
            formulaChangedCount = args.rewriteCellFormulasForSheetRename(op.oldName, op.newName, formulaChangedCount)
            refreshAllPivots = true
            break
          }
          case 'deleteSheet': {
            preparedCells.invalidateSheetName(op.name)
            const removal = args.removeSheetRuntime(op.name, explicitChangedCount)
            changedInputCount += removal.changedInputCount
            formulaChangedCount += removal.formulaChangedCount
            explicitChangedCount = removal.explicitChangedCount
            setEntityVersionForOp(op, order)
            setSheetDeleteVersion(op.name, order)
            topologyChanged = true
            sheetDeleted = true
            structuralInvalidation = true
            refreshAllPivots = true
            break
          }
          case 'insertRows':
          case 'deleteRows':
          case 'moveRows':
          case 'insertColumns':
          case 'deleteColumns':
          case 'moveColumns': {
            const structural = args.applyStructuralAxisOp(op)
            structural.transaction.removedCellIndices.forEach((cellIndex) => {
              precomputedKernelSyncCellIndices.push(cellIndex)
            })
            structural.precomputedChangedInputCellIndices.forEach((cellIndex) => {
              precomputedKernelSyncCellIndices.push(cellIndex)
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            })
            structural.formulaCellIndices.forEach((cellIndex) => {
              formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
            })
            structural.transaction.invalidationSpans.forEach((invalidation) => {
              if (invalidation.axis === 'row') {
                invalidatedRows.push({
                  sheetName: op.sheetName,
                  startIndex: invalidation.start,
                  endIndex: invalidation.end - 1,
                })
                return
              }
              invalidatedColumns.push({
                sheetName: op.sheetName,
                startIndex: invalidation.start,
                endIndex: invalidation.end - 1,
              })
            })
            topologyChanged = structural.graphRefreshRequired || topologyChanged
            refreshAllPivots = true
            setEntityVersionForOp(op, order)
            break
          }
          case 'updateRowMetadata':
          case 'updateColumnMetadata':
          case 'setFreezePane':
          case 'clearFreezePane':
          case 'mergeCells':
          case 'unmergeCells':
          case 'setSheetProtection':
          case 'clearSheetProtection':
          case 'setFilter':
          case 'clearFilter':
          case 'setSort':
          case 'clearSort':
          case 'setDataValidation':
          case 'clearDataValidation':
          case 'upsertConditionalFormat':
          case 'deleteConditionalFormat':
          case 'upsertRangeProtection':
          case 'deleteRangeProtection':
          case 'upsertCommentThread':
          case 'deleteCommentThread':
          case 'upsertNote':
          case 'deleteNote': {
            const metadataChange = applyOperationStructuralMetadataOp({
              workbook: args.state.workbook,
              op,
              order,
              source,
              setEntityVersionForOp,
            })
            structuralInvalidation = structuralInvalidation || metadataChange.structuralInvalidation
            invalidatedRanges.push(...metadataChange.invalidatedRanges)
            invalidatedRows.push(...metadataChange.invalidatedRows)
            invalidatedColumns.push(...metadataChange.invalidatedColumns)
            break
          }
          case 'upsertTable': {
            args.state.workbook.setTable(op.table)
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindTableDependents([tableDependencyKey(op.table.name)], formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            setEntityVersionForOp(op, order)
            break
          }
          case 'deleteTable': {
            args.state.workbook.deleteTable(op.name)
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindTableDependents([tableDependencyKey(op.name)], formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            setEntityVersionForOp(op, order)
            break
          }
          case 'upsertSpillRange':
          case 'deleteSpillRange': {
            const reboundCount = formulaChangedCount
            formulaChangedCount = args.rebindFormulaCells(applySpillRangeOp(op, order), formulaChangedCount)
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            break
          }
          case 'setCellValue': {
            const existingIndex = preparedCells.getExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
            const parsedAddress = preparedCellAddress ?? parseCellAddress(op.address, op.sheetName)
            const sheet = args.state.workbook.getSheet(op.sheetName)
            const sheetId = sheet?.id
            const hasExactLookupDependents = sheetId !== undefined ? hasTrackedExactLookupDependents(sheetId, parsedAddress.col) : false
            const hasSortedLookupDependents = sheetId !== undefined ? hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) : false
            const hasAggregateDependents = sheetId !== undefined ? hasTrackedDirectRangeDependents(sheetId, parsedAddress.col) : false
            if (
              !isRestore &&
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
            const prior = readCellValueForLookup(existingIndex)
            if (!isRestore) {
              if (op.value === null && (existingIndex === undefined || isNullLiteralWriteNoOp(existingIndex))) {
                break
              }
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              }
            }
            const cellIndex = preparedCells.ensureCellTracked(op.sheetName, op.address, preparedCellAddress)
            if (!isRestore) {
              changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
              const removedFormula = args.removeFormula(cellIndex)
              if (removedFormula) {
                args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
                clearLookupImpactCaches()
              }
              if (removedFormula) {
                formulaChangedCount = refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
              }
              topologyChanged = removedFormula || topologyChanged
            }
            writeLiteralToCellStore(args.state.workbook.cellStore, cellIndex, op.value, args.state.strings)
            if (op.value === null) {
              args.state.workbook.cellStore.flags[cellIndex] =
                (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.AuthoredBlank
            }
            args.state.workbook.notifyCellValueWritten(cellIndex)
            if (!isRestore) {
              rebindValueSensitiveFormulaDependents(cellIndex)
            }
            if (needsLookupValueRead) {
              const formulaChangedCountBeforeLookupNotes = formulaChangedCount
              const newValue = literalToValue(op.value, args.state.strings)
              const newStringId = typeof op.value === 'string' ? args.state.workbook.cellStore.stringIds[cellIndex] : undefined
              if (!isRestore) {
                const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                  cellIndex,
                  postRecalcDirectFormulaIndices,
                  prior.value,
                  newValue,
                )
                if (!directDependentsHandled) {
                  markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, postRecalcDirectFormulaIndices)
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
                  formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                    exactLookupRequest,
                    formulaChangedCount,
                    exactLookupImpactCaches,
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
                  formulaChangedCount = markAffectedDirectRangeDependents(
                    exactLookupRequest,
                    formulaChangedCount,
                    postRecalcDirectFormulaIndices,
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
                formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
              }
              if (
                !hasAggregateDependents &&
                (hasExactLookupDependents || hasSortedLookupDependents) &&
                formulaChangedCount === formulaChangedCountBeforeLookupNotes
              ) {
                lookupHandledInputCellIndices.push(cellIndex)
              }
            } else if (!isRestore) {
              const newValue = literalToValue(op.value, args.state.strings)
              const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                cellIndex,
                postRecalcDirectFormulaIndices,
                prior.value,
                newValue,
              )
              if (!directDependentsHandled) {
                markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, postRecalcDirectFormulaIndices)
              }
            }
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
            if (!isRestore && op.value === null) {
              pruneCellIfOrphaned(cellIndex)
            }
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            if (!isRestore) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              setEntityVersionForOp(op, order)
            }
            break
          }
          case 'setCellFormula': {
            const parsedAddress = parseCellAddress(op.address, op.sheetName)
            const sheetId = args.state.workbook.getSheet(op.sheetName)?.id
            if (
              !isRestore &&
              cellTouchesOperationPivotSource({
                workbook: args.state.workbook,
                sheetName: op.sheetName,
                row: parsedAddress.row,
                col: parsedAddress.col,
              })
            ) {
              refreshAllPivots = true
            }
            args.invalidateExactLookupColumn({ sheetName: op.sheetName, col: parsedAddress.col })
            args.invalidateSortedLookupColumn({ sheetName: op.sheetName, col: parsedAddress.col })
            if (!isRestore) {
              const existingIndex = preparedCells.getExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              }
            }
            const cellIndex = preparedCells.ensureCellTracked(op.sheetName, op.address, preparedCellAddress)
            const priorHadFormula = args.state.formulas.get(cellIndex) !== undefined
            const oldFormulaNumber = !isRestore && priorHadFormula ? readExactNumericValueForLookup(cellIndex) : undefined
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.AuthoredBlank
            if (!isRestore) {
              changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
            }
            const compileStarted = isRestore ? 0 : performance.now()
            const hasFormulaColumnAggregateDependents = sheetId !== undefined && hasTrackedDirectRangeDependents(sheetId, parsedAddress.col)
            try {
              const priorDirectScalarFormula = args.state.formulas.get(cellIndex)?.directScalar !== undefined
              const canRewriteFormulaPreservingBinding =
                !isRestore &&
                priorDirectScalarFormula &&
                sheetId !== undefined &&
                !hasTrackedExactLookupDependents(sheetId, parsedAddress.col) &&
                !hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) &&
                !hasFormulaColumnAggregateDependents &&
                args.rewriteFormulaSourcePreservingBinding !== undefined
              const changedTopology = canRewriteFormulaPreservingBinding
                ? args.rewriteFormulaSourcePreservingBinding(cellIndex, op.sheetName, op.formula)
                  ? false
                  : args.bindFormula(cellIndex, op.sheetName, op.formula)
                : args.bindFormula(cellIndex, op.sheetName, op.formula)
              if (hasFormulaColumnAggregateDependents) {
                args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
              }
              clearLookupImpactCaches()
              if (!isRestore) {
                compileMs += performance.now() - compileStarted
              }
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              const handledFormulaReplacementAsDirectDelta =
                priorHadFormula &&
                sheetId !== undefined &&
                !hasTrackedExactLookupDependents(sheetId, parsedAddress.col) &&
                !hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) &&
                !hasFormulaColumnAggregateDependents &&
                tryApplyFormulaReplacementAsDirectScalarDeltaRoot({
                  cellIndex,
                  oldNumber: oldFormulaNumber,
                  changedTopology,
                  postRecalcDirectFormulaIndices,
                  postRecalcDirectFormulaMetrics,
                })
              if (!handledFormulaReplacementAsDirectDelta) {
                formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
              }
              topologyChanged = topologyChanged || changedTopology
              if (!priorHadFormula) {
                formulaChangedCount = refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
                topologyChanged = true
              }
              const aggregateDependents = hasFormulaColumnAggregateDependents
                ? collectAffectedDirectRangeDependents({
                    sheetName: op.sheetName,
                    row: parsedAddress.row,
                    col: parsedAddress.col,
                  }).filter((candidate) => candidate !== cellIndex)
                : []
              if (aggregateDependents.length > 0) {
                formulaChangedCount = args.rebindFormulaCells(aggregateDependents, formulaChangedCount)
                for (let index = 0; index < aggregateDependents.length; index += 1) {
                  postRecalcDirectFormulaIndices.add(aggregateDependents[index]!)
                  formulaChangedCount = args.markFormulaChanged(aggregateDependents[index]!, formulaChangedCount)
                  changedInputCount = args.markInputChanged(aggregateDependents[index]!, changedInputCount)
                }
                topologyChanged = true
              }
            } catch {
              if (!isRestore) {
                compileMs += performance.now() - compileStarted
              }
              topologyChanged = args.removeFormula(cellIndex) || topologyChanged
              args.setInvalidFormulaValue(cellIndex)
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            }
            if (!isRestore) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              setEntityVersionForOp(op, order)
            }
            break
          }
          case 'setCellFormat': {
            const cellIndex = args.ensureCellTracked(op.sheetName, op.address)
            args.state.workbook.setCellFormat(cellIndex, op.format)
            pruneCellIfOrphaned(cellIndex)
            if (!isRestore) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              setEntityVersionForOp(op, order)
            }
            break
          }
          case 'upsertCellStyle':
          case 'upsertCellNumberFormat':
          case 'setStyleRange':
          case 'setFormatRange': {
            const metadataChange = applyOperationStructuralMetadataOp({
              workbook: args.state.workbook,
              op,
              order,
              source,
              setEntityVersionForOp,
            })
            structuralInvalidation = structuralInvalidation || metadataChange.structuralInvalidation
            invalidatedRanges.push(...metadataChange.invalidatedRanges)
            invalidatedRows.push(...metadataChange.invalidatedRows)
            invalidatedColumns.push(...metadataChange.invalidatedColumns)
            break
          }
          case 'clearCell': {
            const cellIndex = preparedCells.getExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
            const parsedAddress = preparedCellAddress ?? parseCellAddress(op.address, op.sheetName)
            const sheet = args.state.workbook.getSheet(op.sheetName)
            const sheetId = sheet?.id
            const hasExactLookupDependents = sheetId !== undefined ? hasTrackedExactLookupDependents(sheetId, parsedAddress.col) : false
            const hasSortedLookupDependents = sheetId !== undefined ? hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) : false
            const hasAggregateDependents = sheetId !== undefined ? hasTrackedDirectRangeDependents(sheetId, parsedAddress.col) : false
            const needsLookupValueRead = hasExactLookupDependents || hasSortedLookupDependents || hasAggregateDependents
            if (
              !isRestore &&
              cellTouchesOperationPivotSource({
                workbook: args.state.workbook,
                sheetName: op.sheetName,
                row: parsedAddress.row,
                col: parsedAddress.col,
              })
            ) {
              refreshAllPivots = true
            }
            const prior = readCellValueForLookup(cellIndex)
            if (cellIndex === undefined) {
              setEntityVersionForOp(op, order)
              break
            }
            if (isClearCellNoOp(cellIndex)) {
              break
            }
            changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(cellIndex), changedInputCount)
            changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
            const removedFormula = args.removeFormula(cellIndex)
            if (removedFormula) {
              args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
              clearLookupImpactCaches()
            }
            if (removedFormula) {
              formulaChangedCount = refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
            }
            topologyChanged = removedFormula || topologyChanged
            args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
            args.state.workbook.notifyCellValueWritten(cellIndex)
            if (!isRestore) {
              rebindValueSensitiveFormulaDependents(cellIndex)
            }
            if (!isRestore) {
              const nextValue = emptyValue()
              const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                cellIndex,
                postRecalcDirectFormulaIndices,
                prior.value,
                nextValue,
              )
              if (!directDependentsHandled) {
                markDirectScalarDeltaClosure(cellIndex, prior.value, nextValue, postRecalcDirectFormulaIndices)
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
                  formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                    exactLookupRequest,
                    formulaChangedCount,
                    exactLookupImpactCaches,
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
                  formulaChangedCount = markAffectedDirectRangeDependents(
                    exactLookupRequest,
                    formulaChangedCount,
                    postRecalcDirectFormulaIndices,
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
                formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
              }
            }
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(
                CellFlags.HasFormula |
                CellFlags.JsOnly |
                CellFlags.InCycle |
                CellFlags.SpillChild |
                CellFlags.PivotOutput |
                CellFlags.AuthoredBlank
              )
            normalizeHistoryDependencyPlaceholder(cellIndex, source)
            pruneCellIfOrphaned(cellIndex)
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            setEntityVersionForOp(op, order)
            break
          }
          case 'upsertDefinedName': {
            const normalizedName = normalizeDefinedName(op.name)
            args.state.workbook.setDefinedName(op.name, op.value)
            const dependentFormulaCells = args.collectFormulaCellsForDefinedNames([normalizedName])
            const canRecalculateWithoutRebind =
              isScalarOnlyDefinedNameValue(op.value) &&
              dependentFormulaCells.every((cellIndex) => args.state.formulas.get(cellIndex)?.compiled.mode === FormulaMode.JsOnly)
            if (canRecalculateWithoutRebind) {
              for (const cellIndex of dependentFormulaCells) {
                formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
              }
            } else {
              formulaChangedCount = args.rebindDefinedNameDependents([normalizedName], formulaChangedCount)
            }
            setEntityVersionForOp(op, order)
            break
          }
          case 'deleteDefinedName': {
            const normalizedName = normalizeDefinedName(op.name)
            args.state.workbook.deleteDefinedName(op.name)
            formulaChangedCount = args.rebindDefinedNameDependents([normalizedName], formulaChangedCount)
            setEntityVersionForOp(op, order)
            break
          }
          case 'upsertPivotTable': {
            const changedPivotUpsertOutputs = applyPivotUpsertOp(op, order)
            changedInputCount = args.markPivotRootsChanged(changedPivotUpsertOutputs, changedInputCount)
            changedPivotUpsertOutputs.forEach((cellIndex) => {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            })
            refreshAllPivots = true
            break
          }
          case 'deletePivotTable': {
            const changedPivotOutputs = applyPivotDeleteOp(op, order)
            changedInputCount = args.markPivotRootsChanged(changedPivotOutputs, changedInputCount)
            changedPivotOutputs.forEach((cellIndex) => {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            })
            refreshAllPivots = true
            break
          }
          case 'upsertChart':
          case 'deleteChart':
          case 'upsertImage':
          case 'deleteImage':
          case 'upsertShape':
          case 'deleteShape': {
            const metadataChange = applyOperationStructuralMetadataOp({
              workbook: args.state.workbook,
              op,
              order,
              source,
              setEntityVersionForOp,
            })
            structuralInvalidation = structuralInvalidation || metadataChange.structuralInvalidation
            invalidatedRanges.push(...metadataChange.invalidatedRanges)
            invalidatedRows.push(...metadataChange.invalidatedRows)
            invalidatedColumns.push(...metadataChange.invalidatedColumns)
            break
          }
          default:
            assertNever(op)
        }
        appliedOps += 1
      })

      const reboundCount = formulaChangedCount
      formulaChangedCount = args.syncDynamicRanges(formulaChangedCount)
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    markBatchApplied(args.state.replicaState, batch)
    if (appliedOps === 0) {
      if (source === 'local') {
        emitBatch(batch)
      }
      return
    }

    finalizeOperationRecalcAndEvents({
      serviceArgs: args,
      isRestore,
      topologyChanged,
      sheetDeleted,
      structuralInvalidation,
      refreshAllPivots,
      changedInputCount,
      formulaChangedCount,
      explicitChangedCount,
      compileMs,
      precomputedKernelSyncCellIndices,
      postRecalcDirectFormulaIndices,
      postRecalcDirectFormulaMetrics,
      lookupHandledInputCellIndices,
      invalidatedRanges,
      invalidatedRows,
      invalidatedColumns,
      hadCycleMembersBeforeNow,
      markCycleMemberInputsChanged,
      canSkipDirtyTraversalForChangedInputs,
      changedInputsNeedRegionQueryIndices,
      directFormulaCallbacks: {
        applyDirectFormulaCurrentResult,
        applyDirectFormulaNumericDelta,
        applyDirectScalarCurrentValue,
        tryApplyDirectScalarDeltas,
        tryApplyDirectFormulaDeltas,
        countPostRecalcDirectFormulaMetric,
      },
    })
    if (source === 'local') {
      void args.state.getSyncClientConnection()?.send(batch)
      emitBatch(batch)
    } else if (source === 'remote' && args.state.redoStack.length > 0) {
      args.state.redoStack.length = 0
    }
  }
}
