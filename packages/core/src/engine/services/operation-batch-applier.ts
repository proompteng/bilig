import { parseCellAddress } from '@bilig/formula'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { FormulaMode, type CellRangeRef, type EngineEvent } from '@bilig/protocol'
import { batchOpOrder, compareOpOrder, markBatchApplied } from '../../replica-state.js'
import { CellFlags } from '../../cell-store.js'
import { calculationSettingsEqual, normalizeWorkbookCalculationSettings, tableDependencyKey } from '../../engine-metadata-utils.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { PreparedCellAddress } from '../runtime-state.js'
import { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import { assertNever } from './operation-change-helpers.js'
import { isScalarOnlyDefinedNameValue } from './defined-name-value-helpers.js'
import { shouldApplyOp as shouldApplyReplicaOp } from './operation-replica-helpers.js'
import { assertProtectionAllowsOp as assertProtectionAllowsProtectedOp } from './operation-protection-helpers.js'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import { applyOperationStructuralMetadataOp } from './operation-structural-metadata-ops.js'
import { createOperationPreparedCellTracker } from './operation-cell-address-resolver.js'
import type { ExactLookupImpactCaches } from './operation-lookup-dirty-markers.js'
import { shouldMaterializeOperationChangedCells } from './operation-event-emission.js'
import { finalizeOperationMutationEvents } from './operation-mutation-event-finalizer.js'
import { finalizeOperationRecalcAndEvents } from './operation-recalc-finalizer.js'
import { createOperationBatchMetrics } from './operation-batch-metrics.js'
import { applyBatchSetCellFormulaOp } from './operation-batch-cell-formula-mutations.js'
import type { MutationSource, StructuralAxisOp } from './operation-service-types.js'
import { applyBatchClearCellOp, applyBatchSetCellValueOp } from './operation-batch-cell-value-mutations.js'
import { canFinalizeStructuralNoValueMutationWithoutRecalc, isStructuralAxisOp } from './operation-structural-no-value-finalization.js'
import type { CreateOperationBatchApplierArgs } from './operation-batch-applier-types.js'

const EMPTY_INVALIDATED_RANGES: CellRangeRef[] = []
const EMPTY_INVALIDATED_ROWS: NonNullable<EngineEvent['invalidatedRows']> = []
const EMPTY_INVALIDATED_COLUMNS: NonNullable<EngineEvent['invalidatedColumns']> = []
const EMPTY_CELL_INDICES: number[] = []

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

  const tryApplySingleStructuralAxisOpBatchNow = (
    batch: EngineOpBatch,
    source: MutationSource,
    options: { readonly emitTracked?: boolean } | undefined,
  ): boolean => {
    if (source === 'remote' || batch.ops.length !== 1) {
      return false
    }
    const op = batch.ops[0]!
    if (!isStructuralAxisOp(op)) {
      return false
    }
    return tryApplySingleStructuralAxisOpNow(op, source, options, batch)
  }

  const tryApplySingleStructuralAxisOpNow = (
    op: StructuralAxisOp,
    source: MutationSource,
    options: { readonly emitTracked?: boolean } | undefined,
    batch?: EngineOpBatch,
  ): boolean => {
    if (source === 'remote') {
      return false
    }
    const isRestore = source === 'restore'
    if (!isRestore && source !== 'undo' && source !== 'redo') {
      assertProtectionAllowsProtectedOp(args.state.workbook, op)
    }

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)

    let formulaChangedCount = 0
    let explicitChangedCount = 0
    let topologyChanged = false
    let mutationCollectionStarted = false
    const beginStructuralMutationCollection = (): void => {
      if (mutationCollectionStarted) {
        return
      }
      args.beginMutationCollection({ ensureScratch: false })
      args.resetMaterializedCellScratch(0)
      mutationCollectionStarted = true
    }
    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = options?.emitTracked === false ? false : args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const shouldCollectInvalidationPayloads = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const invalidatedRows: EngineEvent['invalidatedRows'] = shouldCollectInvalidationPayloads ? [] : EMPTY_INVALIDATED_ROWS
    const invalidatedColumns: EngineEvent['invalidatedColumns'] = shouldCollectInvalidationPayloads ? [] : EMPTY_INVALIDATED_COLUMNS
    let invalidatedRowCount = 0
    let invalidatedColumnCount = 0
    let precomputedKernelSyncCellIndices: number[] | undefined
    let precomputedKernelSyncCellCount = 0
    let hadCycleMembersBefore: boolean | undefined
    let finalizedNoValueWithoutEvents = false
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())

    try {
      const order = batch === undefined ? undefined : batchOpOrder(batch, 0)
      const structural = args.applyStructuralAxisOp(op)
      const activeFormulaCount = args.state.formulas.size
      const canFinalizeNoValueWithoutEvents =
        !isRestore &&
        options?.emitTracked === false &&
        !hasGeneralEventListeners &&
        !hasTrackedEventListeners &&
        !hasWatchedCellListeners &&
        !structural.graphRefreshRequired &&
        structural.transaction.invalidationSpans.length > 0 &&
        structural.transaction.removedCellIndices.length === 0 &&
        structural.precomputedChangedInputCellIndices.length === 0 &&
        structural.formulaCellIndices.length === 0 &&
        !args.state.workbook.hasPivots() &&
        (activeFormulaCount === 0 || args.hasVolatileFormulas?.() === false)
      if (canFinalizeNoValueWithoutEvents) {
        if (order !== undefined) {
          setEntityVersionForOp(op, order)
        }
        finalizedNoValueWithoutEvents = true
      } else {
        if (
          structural.transaction.removedCellIndices.length > 0 ||
          structural.precomputedChangedInputCellIndices.length > 0 ||
          structural.formulaCellIndices.length > 0
        ) {
          beginStructuralMutationCollection()
          args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + 1)
        }
        structural.transaction.removedCellIndices.forEach((cellIndex) => {
          precomputedKernelSyncCellIndices ??= []
          precomputedKernelSyncCellIndices.push(cellIndex)
          precomputedKernelSyncCellCount += 1
        })
        structural.precomputedChangedInputCellIndices.forEach((cellIndex) => {
          precomputedKernelSyncCellIndices ??= []
          precomputedKernelSyncCellIndices.push(cellIndex)
          precomputedKernelSyncCellCount += 1
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        })
        structural.formulaCellIndices.forEach((cellIndex) => {
          formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
        })
        structural.transaction.invalidationSpans.forEach((invalidation) => {
          if (invalidation.axis === 'row') {
            invalidatedRowCount += 1
            if (shouldCollectInvalidationPayloads) {
              invalidatedRows.push({
                sheetName: op.sheetName,
                startIndex: invalidation.start,
                endIndex: invalidation.end - 1,
              })
            }
            return
          }
          invalidatedColumnCount += 1
          if (shouldCollectInvalidationPayloads) {
            invalidatedColumns.push({
              sheetName: op.sheetName,
              startIndex: invalidation.start,
              endIndex: invalidation.end - 1,
            })
          }
        })
        topologyChanged = structural.graphRefreshRequired
        if (order !== undefined) {
          setEntityVersionForOp(op, order)
        }

        const reboundCount = formulaChangedCount
        if (mutationCollectionStarted) {
          formulaChangedCount = args.syncDynamicRanges(formulaChangedCount)
        }
        topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
      }
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    if (batch !== undefined) {
      markBatchApplied(args.state.replicaState, batch)
    }
    if (finalizedNoValueWithoutEvents) {
      args.state.setLastMetrics(
        createOperationBatchMetrics({
          previousMetrics: args.state.getLastMetrics(),
          didRunRecalc: false,
          directFormulaMetrics: {
            wasmFormulaCount: 0,
            jsFormulaCount: 0,
          },
          changedInputCount: 0,
          formulaChangedCount: 0,
          compileMs: 0,
        }),
      )
      if (source === 'local' && batch !== undefined) {
        void args.state.getSyncClientConnection()?.send(batch)
        emitBatch(batch)
      }
      return true
    }
    const activeFormulaCount = args.state.formulas.size
    const canFinalizeWithoutRecalc = canFinalizeStructuralNoValueMutationWithoutRecalc({
      isRestore,
      topologyChanged,
      formulaChangedCount,
      explicitChangedCount,
      precomputedKernelSyncCellCount,
      invalidatedRangeCount: 0,
      invalidatedRowCount,
      invalidatedColumnCount,
      activeFormulaCount,
      hasVolatileFormulas: activeFormulaCount > 0 ? args.hasVolatileFormulas?.() : false,
      hasActivePivots: args.state.workbook.hasPivots(),
    })
    if (canFinalizeWithoutRecalc) {
      if (!hasGeneralEventListeners && !hasTrackedEventListeners && !hasWatchedCellListeners) {
        args.state.setLastMetrics(
          createOperationBatchMetrics({
            previousMetrics: args.state.getLastMetrics(),
            didRunRecalc: false,
            directFormulaMetrics: {
              wasmFormulaCount: 0,
              jsFormulaCount: 0,
            },
            changedInputCount: 0,
            formulaChangedCount: 0,
            compileMs: 0,
          }),
        )
        if (source === 'local' && batch !== undefined) {
          void args.state.getSyncClientConnection()?.send(batch)
          emitBatch(batch)
        }
        return true
      }
      finalizeOperationMutationEvents({
        serviceArgs: args,
        suppressChangedSet: true,
        canComposeDisjointEventChanges: false,
        recalculated: new Uint32Array(),
        explicitChangedCount: 0,
        changedInputCount: 0,
        formulaChangedCount: 0,
        compileMs: 0,
        didRunRecalc: false,
        directFormulaMetrics: {
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
        },
        invalidation: 'cells',
        invalidatedRanges: EMPTY_INVALIDATED_RANGES,
        invalidatedRows,
        invalidatedColumns,
        hasGeneralEventListeners,
        hasTrackedEventListeners,
        hasWatchedCellListeners,
        shouldMaterializeChangedCells: (changedLength) =>
          shouldMaterializeOperationChangedCells({
            changedLength,
            hasGeneralEventListeners,
            invalidation: 'cells',
            invalidations: {
              invalidatedRanges: [],
              invalidatedRows,
              invalidatedColumns,
            },
          }),
      })
      if (source === 'local' && batch !== undefined) {
        void args.state.getSyncClientConnection()?.send(batch)
        emitBatch(batch)
      }
      return true
    }
    if (!mutationCollectionStarted) {
      beginStructuralMutationCollection()
      args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + 1)
    }
    finalizeOperationRecalcAndEvents({
      serviceArgs: args,
      isRestore,
      topologyChanged,
      sheetDeleted: false,
      structuralInvalidation: false,
      refreshAllPivots: true,
      changedInputCount: 0,
      formulaChangedCount,
      explicitChangedCount,
      compileMs: 0,
      precomputedKernelSyncCellIndices: precomputedKernelSyncCellIndices ?? EMPTY_CELL_INDICES,
      postRecalcDirectFormulaIndices: new DirectFormulaIndexCollection(),
      postRecalcDirectFormulaMetrics: {
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
      },
      lookupHandledInputCellIndices: [],
      invalidatedRanges: [],
      invalidatedRows,
      invalidatedColumns,
      invalidatedRowCount,
      invalidatedColumnCount,
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
    if (source === 'local' && batch !== undefined) {
      void args.state.getSyncClientConnection()?.send(batch)
      emitBatch(batch)
    }
    return true
  }

  const applyBatchNow = (
    batch: EngineOpBatch,
    source: MutationSource,
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
    options?: { readonly emitTracked?: boolean },
  ): void => {
    if (preparedCellAddressesByOpIndex && preparedCellAddressesByOpIndex.length !== batch.ops.length) {
      throw new Error('Prepared cell addresses must align with batch operations')
    }
    if (preparedCellAddressesByOpIndex === undefined && tryApplySingleStructuralAxisOpBatchNow(batch, source, options)) {
      return
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
      invalidatedRowCount: invalidatedRows.length,
      invalidatedColumnCount: invalidatedColumns.length,
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

  return {
    applyBatchNow,
    applyLocalSingleStructuralAxisOpWithoutBatchNow(op: StructuralAxisOp, options?: { readonly emitTracked?: boolean }): boolean {
      return tryApplySingleStructuralAxisOpNow(op, 'local', options)
    },
  }
}
