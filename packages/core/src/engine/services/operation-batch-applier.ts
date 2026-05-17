import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import { FormulaMode, type CellRangeRef, type CellValue } from '@bilig/protocol'
import { batchOpOrder, compareOpOrder, markBatchApplied, type OpOrder } from '../../replica-state.js'
import { calculationSettingsEqual, normalizeWorkbookCalculationSettings, tableDependencyKey } from '../../engine-metadata-utils.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { PreparedCellAddress } from '../runtime-state.js'
import { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import { assertNever } from './operation-change-helpers.js'
import { isScalarOnlyDefinedNameValue } from './defined-name-value-helpers.js'
import { shouldApplyOp as shouldApplyReplicaOp, type OperationReplicaVersionWriter } from './operation-replica-helpers.js'
import { assertProtectionAllowsOp as assertProtectionAllowsProtectedOp } from './operation-protection-helpers.js'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import { applyOperationStructuralMetadataOp } from './operation-structural-metadata-ops.js'
import { createOperationPreparedCellTracker } from './operation-cell-address-resolver.js'
import type { OperationLookupAccess } from './operation-lookup-access.js'
import type { OperationLookupPlanner } from './operation-lookup-planner.js'
import type { ExactLookupImpactCaches, OperationLookupDirtyMarkerService } from './operation-lookup-dirty-markers.js'
import type { OperationColumnDependencyTrackerService } from './operation-column-dependency-tracker.js'
import type { OperationDirectRangeDependentService } from './operation-direct-range-dependents.js'
import { finalizeOperationRecalcAndEvents } from './operation-recalc-finalizer.js'
import type { CreateEngineOperationServiceArgs, MutationSource } from './operation-service-types.js'
import { applyBatchSetCellFormulaOp } from './operation-batch-cell-formula-mutations.js'
import { applyBatchClearCellOp, applyBatchSetCellValueOp } from './operation-batch-cell-value-mutations.js'

type OperationBatchDirectFormulaCallbacks = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['directFormulaCallbacks']
type OperationBatchDirtyTraversalSkip = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['canSkipDirtyTraversalForChangedInputs']
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
  readonly planExactLookupNumericColumnWrite: OperationLookupPlanner['planExactLookupNumericColumnWrite']
  readonly planApproximateLookupNumericColumnWrite: OperationLookupPlanner['planApproximateLookupNumericColumnWrite']
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
    planExactLookupNumericColumnWrite,
    planApproximateLookupNumericColumnWrite,
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
            const setValueResult = applyBatchSetCellValueOp({
              serviceArgs: args,
              op,
              order,
              source,
              isRestore,
              preparedCellAddress,
              preparedCells,
              changedInputCount,
              formulaChangedCount,
              explicitChangedCount,
              topologyChanged,
              refreshAllPivots,
              postRecalcDirectFormulaIndices,
              exactLookupImpactCaches,
              setEntityVersionForOp,
              hasTrackedExactLookupDependents,
              hasTrackedSortedLookupDependents,
              hasTrackedDirectRangeDependents,
              planExactLookupNumericColumnWrite,
              planApproximateLookupNumericColumnWrite,
              allowLookupTailPatch: batch.ops.length === 1,
              readCellValueForLookup,
              isNullLiteralWriteNoOp,
              rebindValueSensitiveFormulaDependents: (cellIndex, counts) => {
                const reboundCount = counts.formulaChangedCount
                const nextFormulaChangedCount = rebindDynamicFormulaDependents(cellIndex, counts.formulaChangedCount)
                return {
                  ...counts,
                  formulaChangedCount: nextFormulaChangedCount,
                  topologyChanged: counts.topologyChanged || nextFormulaChangedCount !== reboundCount,
                }
              },
              refreshDependentRangesAndRebindFormulaDependents,
              markPostRecalcDirectFormulaDependents,
              markDirectScalarDeltaClosure,
              noteExactLookupLiteralWriteWhenDirty,
              noteSortedLookupLiteralWriteWhenDirty,
              markAffectedDirectRangeDependents,
              lookupHandledInputCellIndices,
              clearLookupImpactCaches,
              pruneCellIfOrphaned,
            })
            changedInputCount = setValueResult.changedInputCount
            formulaChangedCount = setValueResult.formulaChangedCount
            explicitChangedCount = setValueResult.explicitChangedCount
            topologyChanged = setValueResult.topologyChanged
            refreshAllPivots = setValueResult.refreshAllPivots
            break
          }
          case 'setCellFormula': {
            const formulaResult = applyBatchSetCellFormulaOp({
              serviceArgs: args,
              op,
              order,
              isRestore,
              preparedCellAddress,
              preparedCells,
              changedInputCount,
              formulaChangedCount,
              explicitChangedCount,
              topologyChanged,
              refreshAllPivots,
              compileMs,
              postRecalcDirectFormulaIndices,
              postRecalcDirectFormulaMetrics,
              setEntityVersionForOp,
              hasTrackedExactLookupDependents,
              hasTrackedSortedLookupDependents,
              hasTrackedDirectRangeDependents,
              readExactNumericValueForLookup,
              tryApplyFormulaReplacementAsDirectScalarDeltaRoot,
              refreshDependentRangesAndRebindFormulaDependents,
              collectAffectedDirectRangeDependents,
              clearLookupImpactCaches,
            })
            changedInputCount = formulaResult.changedInputCount
            formulaChangedCount = formulaResult.formulaChangedCount
            explicitChangedCount = formulaResult.explicitChangedCount
            topologyChanged = formulaResult.topologyChanged
            refreshAllPivots = formulaResult.refreshAllPivots
            compileMs = formulaResult.compileMs
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
            const clearResult = applyBatchClearCellOp({
              serviceArgs: args,
              op,
              order,
              source,
              isRestore,
              preparedCellAddress,
              preparedCells,
              changedInputCount,
              formulaChangedCount,
              explicitChangedCount,
              topologyChanged,
              refreshAllPivots,
              postRecalcDirectFormulaIndices,
              exactLookupImpactCaches,
              setEntityVersionForOp,
              hasTrackedExactLookupDependents,
              hasTrackedSortedLookupDependents,
              hasTrackedDirectRangeDependents,
              readCellValueForLookup,
              isClearCellNoOp,
              rebindValueSensitiveFormulaDependents: (cellIndex, counts) => {
                const reboundCount = counts.formulaChangedCount
                const nextFormulaChangedCount = rebindDynamicFormulaDependents(cellIndex, counts.formulaChangedCount)
                return {
                  ...counts,
                  formulaChangedCount: nextFormulaChangedCount,
                  topologyChanged: counts.topologyChanged || nextFormulaChangedCount !== reboundCount,
                }
              },
              refreshDependentRangesAndRebindFormulaDependents,
              markPostRecalcDirectFormulaDependents,
              markDirectScalarDeltaClosure,
              noteExactLookupLiteralWriteWhenDirty,
              noteSortedLookupLiteralWriteWhenDirty,
              markAffectedDirectRangeDependents,
              clearLookupImpactCaches,
              pruneCellIfOrphaned,
              normalizeHistoryDependencyPlaceholder,
            })
            changedInputCount = clearResult.changedInputCount
            formulaChangedCount = clearResult.formulaChangedCount
            explicitChangedCount = clearResult.explicitChangedCount
            topologyChanged = clearResult.topologyChanged
            refreshAllPivots = clearResult.refreshAllPivots
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
