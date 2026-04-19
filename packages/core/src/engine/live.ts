import { Effect, Exit, Cause } from 'effect'
import { ValueTag, type CellSnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import { CellFlags } from '../cell-store.js'
import { createColumnIndexStore } from '../indexes/column-index-store.js'
import type { EngineRuntimeState } from './runtime-state.js'
import { createEngineCellStateService, type EngineCellStateService } from './services/cell-state-service.js'
import { createEngineEventService, type EngineEventService } from './services/event-service.js'
import { createEngineChangeSetEmitterService } from './services/change-set-emitter-service.js'
import { createEngineFormulaEvaluationService, type EngineFormulaEvaluationService } from './services/formula-evaluation-service.js'
import { createCriterionRangeCacheService } from './services/criterion-range-cache-service.js'
import { createExactColumnIndexService } from './services/exact-column-index-service.js'
import { createRangeAggregateCacheService } from './services/range-aggregate-cache-service.js'
import { createSortedColumnSearchService } from './services/sorted-column-search-service.js'
import { createEngineFormulaBindingService, type EngineFormulaBindingService } from './services/formula-binding-service.js'
import {
  createEngineFormulaInitializationService,
  type EngineFormulaInitializationService,
} from './services/formula-initialization-service.js'
import { createEngineFormulaTemplateNormalizationService } from './services/formula-template-normalization-service.js'
import { createEngineCompiledPlanService } from './services/compiled-plan-service.js'
import { createFormulaInstanceTable } from '../formula/formula-instance-table.js'
import { createEngineFormulaGraphService, type EngineFormulaGraphService } from './services/formula-graph-service.js'
import { createEngineHistoryService, type EngineHistoryService } from './services/history-service.js'
import { createEngineMaintenanceService, type EngineMaintenanceService } from './services/maintenance-service.js'
import { createEngineMutationService, type EngineMutationService } from './services/mutation-service.js'
import { createEngineMutationSupportService, type EngineMutationSupportService } from './services/mutation-support-service.js'
import { createEngineOperationService, type EngineOperationService } from './services/operation-service.js'
import { createEnginePivotService, type EnginePivotService } from './services/pivot-service.js'
import {
  createEngineDirtyFrontierSchedulerService,
  type EngineDirtyFrontierSchedulerService,
} from './services/dirty-frontier-scheduler-service.js'
import { createEngineRuntimeColumnStoreService } from './services/runtime-column-store-service.js'
import { createEngineReplicaSyncService, type EngineReplicaSyncService } from './services/replica-sync-service.js'
import { createEngineReadService, type EngineReadService } from './services/read-service.js'
import { createEngineRecalcService, type EngineRecalcService } from './services/recalc-service.js'
import { createEngineRuntimeScratchService } from './services/runtime-scratch-service.js'
import { createEngineSelectionService, type EngineSelectionService } from './services/selection-service.js'
import { createEngineSnapshotService, type EngineSnapshotService } from './services/snapshot-service.js'
import { createEngineStructureService, type EngineStructureService } from './services/structure-service.js'
import { createEngineTraversalService, type EngineTraversalService } from './services/traversal-service.js'

export interface EngineServiceRuntime {
  readonly cellState: EngineCellStateService
  readonly maintenance: EngineMaintenanceService
  readonly traversal: EngineTraversalService
  readonly events: EngineEventService
  readonly evaluation: EngineFormulaEvaluationService
  readonly selection: EngineSelectionService
  readonly binding: EngineFormulaBindingService
  readonly formulaInitialization: EngineFormulaInitializationService
  readonly graph: EngineFormulaGraphService
  readonly history: EngineHistoryService
  readonly mutation: EngineMutationService
  readonly support: EngineMutationSupportService
  readonly operations: EngineOperationService
  readonly pivot: EnginePivotService
  readonly read: EngineReadService
  readonly recalc: EngineRecalcService
  readonly structure: EngineStructureService
  readonly snapshot: EngineSnapshotService
  readonly sync: EngineReplicaSyncService
}

type EngineMutationSupportRuntimeConfig = Omit<
  Parameters<typeof createEngineMutationSupportService>[0],
  | 'removeFormula'
  | 'rebindFormulasForSheet'
  | 'applyDerivedOp'
  | 'scheduleWasmProgramSync'
  | 'collectFormulaDependents'
  | 'ensureRecalcScratchCapacity'
  | 'getChangedInputEpoch'
  | 'setChangedInputEpoch'
  | 'getChangedInputSeen'
  | 'setChangedInputSeen'
  | 'getChangedInputBuffer'
  | 'setChangedInputBuffer'
  | 'getChangedFormulaEpoch'
  | 'setChangedFormulaEpoch'
  | 'getChangedFormulaSeen'
  | 'setChangedFormulaSeen'
  | 'getChangedFormulaBuffer'
  | 'setChangedFormulaBuffer'
  | 'getChangedUnionEpoch'
  | 'setChangedUnionEpoch'
  | 'getChangedUnionSeen'
  | 'setChangedUnionSeen'
  | 'getChangedUnion'
  | 'setChangedUnion'
  | 'getMutationRoots'
  | 'setMutationRoots'
  | 'getMaterializedCellCount'
  | 'setMaterializedCellCount'
  | 'getMaterializedCells'
  | 'setMaterializedCells'
  | 'getExplicitChangedEpoch'
  | 'setExplicitChangedEpoch'
  | 'getExplicitChangedSeen'
  | 'setExplicitChangedSeen'
  | 'getExplicitChangedBuffer'
  | 'setExplicitChangedBuffer'
  | 'getImpactedFormulaEpoch'
  | 'setImpactedFormulaEpoch'
  | 'getImpactedFormulaSeen'
  | 'setImpactedFormulaSeen'
  | 'getImpactedFormulaBuffer'
  | 'setImpactedFormulaBuffer'
>

type EngineFormulaBindingRuntimeConfig = Omit<
  Parameters<typeof createEngineFormulaBindingService>[0],
  | 'compiledPlans'
  | 'formulaInstances'
  | 'resolveTemplateForCell'
  | 'exactLookup'
  | 'sortedLookup'
  | 'ensureCellTracked'
  | 'ensureCellTrackedByCoords'
  | 'markFormulaChanged'
  | 'forEachSheetCell'
  | 'lookup'
  | 'resolveStructuredReference'
  | 'resolveSpillReference'
  | 'scheduleWasmProgramSync'
>

type EngineFormulaGraphRuntimeConfig = Omit<
  Parameters<typeof createEngineFormulaGraphService>[0],
  'notifyCellValueWritten' | 'forEachFormulaDependencyCell' | 'collectFormulaDependents'
>

type EngineRecalcRuntimeConfig = Omit<
  Parameters<typeof createEngineRecalcService>[0],
  | 'beginMutationCollection'
  | 'markInputChanged'
  | 'markFormulaChanged'
  | 'markExplicitChanged'
  | 'composeMutationRoots'
  | 'composeEventChanges'
  | 'captureChangedCells'
  | 'unionChangedSets'
  | 'composeChangedRootsAndOrdered'
  | 'emptyChangedSet'
  | 'ensureRecalcScratchCapacity'
  | 'getPendingKernelSync'
  | 'getDeferredKernelSyncCount'
  | 'setDeferredKernelSyncCount'
  | 'getDeferredKernelSyncEpoch'
  | 'setDeferredKernelSyncEpoch'
  | 'getDeferredKernelSyncSeen'
  | 'getWasmBatch'
  | 'getChangedInputBuffer'
  | 'dirtyScheduler'
  | 'materializeSpill'
  | 'clearOwnedSpill'
  | 'evaluateDirectLookupFormula'
  | 'evaluateUnsupportedFormula'
  | 'materializePivot'
>

type EngineMaintenanceRuntimeConfig = Omit<
  Parameters<typeof createEngineMaintenanceService>[0],
  | 'captureSheetCellState'
  | 'captureRowRangeCellState'
  | 'captureColumnRangeCellState'
  | 'setMaterializedCellCount'
  | 'resetFormulaRuntimeCaches'
  | 'scheduleWasmProgramSync'
>

type EnginePivotRuntimeConfig = Omit<
  Parameters<typeof createEnginePivotService>[0],
  | 'ensureCellTrackedByCoords'
  | 'forEachSheetCell'
  | 'flushDeferredKernelSync'
  | 'scheduleWasmProgramSync'
  | 'flushWasmProgramSync'
  | 'applyDerivedOp'
>

type EngineOperationRuntimeConfig = Omit<
  Parameters<typeof createEngineOperationService>[0],
  | 'getSelectionState'
  | 'setSelection'
  | 'rewriteDefinedNamesForSheetRename'
  | 'rewriteCellFormulasForSheetRename'
  | 'estimatePotentialNewCells'
  | 'rebindDefinedNameDependents'
  | 'rebindTableDependents'
  | 'rebindFormulaCells'
  | 'rebindFormulasForSheet'
  | 'removeSheetRuntime'
  | 'applyStructuralAxisOp'
  | 'clearOwnedSpill'
  | 'clearPivotForCell'
  | 'clearOwnedPivot'
  | 'removeFormula'
  | 'bindFormula'
  | 'setInvalidFormulaValue'
  | 'beginMutationCollection'
  | 'markInputChanged'
  | 'markFormulaChanged'
  | 'markVolatileFormulasChanged'
  | 'markSpillRootsChanged'
  | 'markPivotRootsChanged'
  | 'markExplicitChanged'
  | 'composeMutationRoots'
  | 'composeEventChanges'
  | 'captureChangedCells'
  | 'getChangedInputBuffer'
  | 'ensureRecalcScratchCapacity'
  | 'ensureCellTracked'
  | 'resetMaterializedCellScratch'
  | 'syncDynamicRanges'
  | 'rebuildTopoRanks'
  | 'detectCycles'
  | 'recalculate'
  | 'evaluateDirectFormula'
  | 'reconcilePivotOutputs'
  | 'flushWasmProgramSync'
>

type EngineTraversalRuntimeConfig = Parameters<typeof createEngineTraversalService>[0]

function requireService<Service>(service: Service | undefined, name: string): Service {
  if (service === undefined) {
    throw new Error(`Engine service ${name} is not initialized`)
  }
  return service
}

export function createEngineServiceRuntime(args: {
  readonly state: EngineRuntimeState
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly exportSnapshot: () => WorkbookSnapshot
  readonly importSnapshot: (snapshot: WorkbookSnapshot) => void
  readonly maintenance: EngineMaintenanceRuntimeConfig
  readonly mutationSupport: EngineMutationSupportRuntimeConfig
  readonly formulaBinding: EngineFormulaBindingRuntimeConfig
  readonly formulaGraph: EngineFormulaGraphRuntimeConfig
  readonly recalc: EngineRecalcRuntimeConfig
  readonly pivot: EnginePivotRuntimeConfig
  readonly operation: EngineOperationRuntimeConfig
  readonly traversal: EngineTraversalRuntimeConfig
  readonly cellToCsvValue: (cell: CellSnapshot) => string
  readonly serializeCsv: (rows: string[][]) => string
  readonly pivotState: {
    readonly pivotOutputOwners: Map<number, string>
  }
  readonly applyRemoteSnapshot: (snapshot: WorkbookSnapshot) => void
}): EngineServiceRuntime {
  const scratch = createEngineRuntimeScratchService()
  const traversal = createEngineTraversalService(args.traversal)
  const changeSetEmitter = createEngineChangeSetEmitterService({ state: args.state })
  const runtimeColumnStore = createEngineRuntimeColumnStoreService({ state: args.state })
  const columnIndexStore = createColumnIndexStore({ state: args.state, runtimeColumnStore })
  const compiledPlans = createEngineCompiledPlanService()
  const formulaTemplates = createEngineFormulaTemplateNormalizationService({ counters: args.state.counters })
  const formulaInstances = createFormulaInstanceTable()
  const criterionCache = createCriterionRangeCacheService({ runtimeColumnStore })
  const aggregateCache = createRangeAggregateCacheService({
    state: args.state,
    runtimeColumnStore,
  })
  const exactLookup = createExactColumnIndexService({ state: args.state, runtimeColumnStore, columnIndexStore })
  const sortedLookup = createSortedColumnSearchService({ state: args.state, runtimeColumnStore, columnIndexStore })
  const graph = createEngineFormulaGraphService({
    ...args.formulaGraph,
    notifyCellValueWritten: (cellIndex) => args.state.workbook.notifyCellValueWritten(cellIndex),
    forEachFormulaDependencyCell: (cellIndex, fn) => traversal.forEachFormulaDependencyCellNow(cellIndex, fn),
    collectFormulaDependents: (entityId) => traversal.collectFormulaDependentsNow(entityId),
  })
  let binding: EngineFormulaBindingService | undefined
  let operations: EngineOperationService | undefined
  let pivot: EnginePivotService | undefined
  let recalc: EngineRecalcService | undefined
  let formulaInitialization: EngineFormulaInitializationService | undefined
  let dirtyScheduler: EngineDirtyFrontierSchedulerService | undefined
  const selection = createEngineSelectionService(args.state)
  const support = createEngineMutationSupportService({
    ...args.mutationSupport,
    removeFormula: (cellIndex) => runEngineEffect(requireService(binding, 'binding').clearFormula(cellIndex)),
    rebindFormulasForSheet: (sheetName, formulaChangedCount, candidates) =>
      runEngineEffect(requireService(binding, 'binding').rebindFormulasForSheet(sheetName, formulaChangedCount, candidates)),
    applyDerivedOp: (op) => runEngineEffect(requireService(operations, 'operations').applyDerivedOp(op)),
    collectFormulaDependents: (entityId) => traversal.collectFormulaDependentsNow(entityId),
    ensureRecalcScratchCapacity: (size) => runEngineEffect(scratch.ensureRecalcCapacity(size)),
    getChangedInputEpoch: () => scratch.getChangedInputEpochNow(),
    setChangedInputEpoch: (next) => {
      scratch.setChangedInputEpochNow(next)
    },
    getChangedInputSeen: () => scratch.getChangedInputSeenNow(),
    setChangedInputSeen: (next) => {
      scratch.setChangedInputSeenNow(next)
    },
    getChangedInputBuffer: () => scratch.getChangedInputBufferNow(),
    setChangedInputBuffer: (next) => {
      scratch.setChangedInputBufferNow(next)
    },
    getChangedFormulaEpoch: () => scratch.getChangedFormulaEpochNow(),
    setChangedFormulaEpoch: (next) => {
      scratch.setChangedFormulaEpochNow(next)
    },
    getChangedFormulaSeen: () => scratch.getChangedFormulaSeenNow(),
    setChangedFormulaSeen: (next) => {
      scratch.setChangedFormulaSeenNow(next)
    },
    getChangedFormulaBuffer: () => scratch.getChangedFormulaBufferNow(),
    setChangedFormulaBuffer: (next) => {
      scratch.setChangedFormulaBufferNow(next)
    },
    getChangedUnionEpoch: () => scratch.getChangedUnionEpochNow(),
    setChangedUnionEpoch: (next) => {
      scratch.setChangedUnionEpochNow(next)
    },
    getChangedUnionSeen: () => scratch.getChangedUnionSeenNow(),
    setChangedUnionSeen: (next) => {
      scratch.setChangedUnionSeenNow(next)
    },
    getChangedUnion: () => scratch.getChangedUnionNow(),
    setChangedUnion: (next) => {
      scratch.setChangedUnionNow(next)
    },
    getMutationRoots: () => scratch.getMutationRootsNow(),
    setMutationRoots: (next) => {
      scratch.setMutationRootsNow(next)
    },
    getMaterializedCellCount: () => scratch.getMaterializedCellCountNow(),
    setMaterializedCellCount: (next) => {
      scratch.setMaterializedCellCountNow(next)
    },
    getMaterializedCells: () => scratch.getMaterializedCellsNow(),
    setMaterializedCells: (next) => {
      scratch.setMaterializedCellsNow(next)
    },
    getExplicitChangedEpoch: () => scratch.getExplicitChangedEpochNow(),
    setExplicitChangedEpoch: (next) => {
      scratch.setExplicitChangedEpochNow(next)
    },
    getExplicitChangedSeen: () => scratch.getExplicitChangedSeenNow(),
    setExplicitChangedSeen: (next) => {
      scratch.setExplicitChangedSeenNow(next)
    },
    getExplicitChangedBuffer: () => scratch.getExplicitChangedBufferNow(),
    setExplicitChangedBuffer: (next) => {
      scratch.setExplicitChangedBufferNow(next)
    },
    getImpactedFormulaEpoch: () => scratch.getImpactedFormulaEpochNow(),
    setImpactedFormulaEpoch: (next) => {
      scratch.setImpactedFormulaEpochNow(next)
    },
    getImpactedFormulaSeen: () => scratch.getImpactedFormulaSeenNow(),
    setImpactedFormulaSeen: (next) => {
      scratch.setImpactedFormulaSeenNow(next)
    },
    getImpactedFormulaBuffer: () => scratch.getImpactedFormulaBufferNow(),
    setImpactedFormulaBuffer: (next) => {
      scratch.setImpactedFormulaBufferNow(next)
    },
    scheduleWasmProgramSync: () => runEngineEffect(graph.scheduleWasmProgramSync()),
  })
  const evaluation = createEngineFormulaEvaluationService({
    state: args.state,
    runtimeColumnStore,
    criterionCache,
    aggregateCache,
    exactLookup,
    sortedLookup,
    materializeSpill: (cellIndex, arrayValue) => support.materializeSpillNow(cellIndex, arrayValue),
    clearOwnedSpill: (cellIndex) => support.clearOwnedSpillNow(cellIndex),
    resolvePivotData: (sheetName, address, dataField, filters) =>
      runEngineEffect(requireService(pivot, 'pivot').resolvePivotData(sheetName, address, dataField, filters)),
  })
  binding = createEngineFormulaBindingService({
    ...args.formulaBinding,
    compiledPlans,
    formulaInstances,
    resolveTemplateForCell: (source, row, col) => formulaTemplates.resolveForCell(source, row, col),
    exactLookup,
    sortedLookup,
    ensureCellTracked: (sheetName, address) => support.ensureCellTrackedNow(sheetName, address),
    ensureCellTrackedByCoords: (sheetId, row, col) => support.ensureCellTrackedByCoordsNow(sheetId, row, col),
    markFormulaChanged: (cellIndex, count) => support.markFormulaChangedNow(cellIndex, count),
    resolveStructuredReference: (tableName, columnName) => runEngineEffect(evaluation.resolveStructuredReference(tableName, columnName)),
    resolveSpillReference: (currentSheetName, sheetName, address) =>
      runEngineEffect(evaluation.resolveSpillReference(currentSheetName, sheetName, address)),
    forEachSheetCell: (sheetId, fn) => traversal.forEachSheetCellNow(sheetId, fn),
    scheduleWasmProgramSync: () => graph.scheduleWasmProgramSyncNow(),
  })
  const read = createEngineReadService({
    state: args.state,
    runtimeColumnStore,
    forEachFormulaDependencyCell: (cellIndex, fn) => traversal.forEachFormulaDependencyCellNow(cellIndex, fn),
    getEntityDependents: (entityId) => traversal.getEntityDependentsNow(entityId),
    cellToCsvValue: args.cellToCsvValue,
    serializeCsv: args.serializeCsv,
  })
  const cellState = createEngineCellStateService({
    state: args.state,
    getCell: (sheetName, address) => runEngineEffect(read.getCell(sheetName, address)),
    getCellByIndex: (cellIndex) => runEngineEffect(read.getCellByIndex(cellIndex)),
  })
  const structure = createEngineStructureService({
    state: {
      workbook: args.state.workbook,
      formulas: args.state.formulas,
      ranges: args.state.ranges,
      pivotOutputOwners: args.pivotState.pivotOutputOwners,
    },
    captureStoredCellOps: (cellIndex, sheetName, address, sourceSheetName, sourceAddress) =>
      cellState.captureStoredCellOpsNow(cellIndex, sheetName, address, sourceSheetName, sourceAddress),
    removeFormula: (cellIndex) => binding.clearFormulaNow(cellIndex),
    clearOwnedPivot: (pivotRecord) => requireService(pivot, 'pivot').clearOwnedPivotNow(pivotRecord),
    refreshRangeDependencies: (rangeIndices) => binding.refreshRangeDependenciesNow(rangeIndices),
    retargetRangeDependencies: (transaction, rangeIndices) => binding.retargetRangeDependenciesNow(transaction, rangeIndices),
    rebindFormulaCells: (inputs) => {
      const pending = inputs.filter(({ cellIndex }) => args.state.formulas.get(cellIndex))
      pending.forEach(({ cellIndex, ownerSheetName, source, compiled, preservesBinding }) => {
        try {
          if (compiled) {
            if (preservesBinding === true && binding.rewriteFormulaCompiledPreservingBindingNow(cellIndex, source, compiled)) {
              return
            }
            binding.bindPreparedFormulaNow(cellIndex, ownerSheetName, source, compiled)
          } else {
            binding.bindFormulaNow(cellIndex, ownerSheetName, source)
          }
        } catch {
          binding.invalidateFormulaNow(cellIndex)
        }
      })
    },
  })
  const maintenance = createEngineMaintenanceService({
    ...args.maintenance,
    captureSheetCellState: (sheetName) => runEngineEffect(structure.captureSheetCellState(sheetName)),
    captureRowRangeCellState: (sheetName, start, count) => runEngineEffect(structure.captureRowRangeCellState(sheetName, start, count)),
    captureColumnRangeCellState: (sheetName, start, count) =>
      runEngineEffect(structure.captureColumnRangeCellState(sheetName, start, count)),
    setMaterializedCellCount: (next) => {
      scratch.setMaterializedCellCountNow(next)
    },
    resetFormulaRuntimeCaches: () => {
      compiledPlans.clear()
      formulaTemplates.reset()
      formulaInstances.clear()
    },
    resetWasmState: () => {
      args.state.wasm.resetStoreState()
    },
    scheduleWasmProgramSync: () => runEngineEffect(graph.scheduleWasmProgramSync()),
  })
  dirtyScheduler = createEngineDirtyFrontierSchedulerService({
    state: args.state,
    getEntityDependents: (entityId) => traversal.getEntityDependentsNow(entityId),
  })
  recalc = createEngineRecalcService({
    ...args.recalc,
    beginMutationCollection: () => support.beginMutationCollectionNow(),
    markInputChanged: (cellIndex, count) => support.markInputChangedNow(cellIndex, count),
    markFormulaChanged: (cellIndex, count) => support.markFormulaChangedNow(cellIndex, count),
    markExplicitChanged: (cellIndex, count) => support.markExplicitChangedNow(cellIndex, count),
    composeMutationRoots: (changedInputCount, formulaChangedCount) =>
      support.composeMutationRootsNow(changedInputCount, formulaChangedCount),
    composeEventChanges: (recalculated, explicitChangedCount) => support.composeEventChangesNow(recalculated, explicitChangedCount),
    captureChangedCells: (changedCellIndices) => changeSetEmitter.captureChangedCells(changedCellIndices),
    unionChangedSets: (...sets) => support.unionChangedSetsNow(...sets),
    composeChangedRootsAndOrdered: (changedRoots, ordered, orderedCount) =>
      support.composeChangedRootsAndOrderedNow(changedRoots, ordered, orderedCount),
    emptyChangedSet: () => support.unionChangedSetsNow(),
    ensureRecalcScratchCapacity: (size) => scratch.ensureRecalcCapacityNow(size),
    getPendingKernelSync: () => scratch.getPendingKernelSyncNow(),
    getDeferredKernelSyncCount: () => scratch.getDeferredKernelSyncCountNow(),
    setDeferredKernelSyncCount: (next) => scratch.setDeferredKernelSyncCountNow(next),
    getDeferredKernelSyncEpoch: () => scratch.getDeferredKernelSyncEpochNow(),
    setDeferredKernelSyncEpoch: (next) => scratch.setDeferredKernelSyncEpochNow(next),
    getDeferredKernelSyncSeen: () => scratch.getDeferredKernelSyncSeenNow(),
    getWasmBatch: () => scratch.getWasmBatchNow(),
    getChangedInputBuffer: () => support.getChangedInputBufferNow(),
    dirtyScheduler,
    materializeSpill: (cellIndex, arrayValue) => support.materializeSpillNow(cellIndex, arrayValue),
    clearOwnedSpill: (cellIndex) => support.clearOwnedSpillNow(cellIndex),
    evaluateDirectLookupFormula: (cellIndex) => evaluation.evaluateDirectLookupFormulaNow(cellIndex),
    evaluateUnsupportedFormula: (cellIndex) => runEngineEffect(evaluation.evaluateUnsupportedFormula(cellIndex)),
    materializePivot: (pivotRecord) => requireService(pivot, 'pivot').materializePivotNow(pivotRecord),
  })
  formulaInitialization = createEngineFormulaInitializationService({
    state: args.state,
    beginMutationCollection: () => support.beginMutationCollectionNow(),
    ensureRecalcScratchCapacity: (size) => scratch.ensureRecalcCapacityNow(size),
    ensureCellTrackedByCoords: (sheetId, row, col) => support.ensureCellTrackedByCoordsNow(sheetId, row, col),
    resetMaterializedCellScratch: (expectedSize) => support.resetMaterializedCellScratchNow(expectedSize),
    bindFormula: (cellIndex, ownerSheetName, source) => binding.bindInitialFormulaNow(cellIndex, ownerSheetName, source),
    bindPreparedFormula: (cellIndex, ownerSheetName, source, compiled, templateId) =>
      binding.bindPreparedFormulaNow(cellIndex, ownerSheetName, source, compiled, templateId),
    compileTemplateFormula: (source, row, col) => formulaTemplates.resolveForCell(source, row, col),
    clearTemplateFormulaCache: () => formulaTemplates.clear(),
    removeFormula: (cellIndex) => binding.clearFormulaNow(cellIndex),
    setInvalidFormulaValue: (cellIndex) => binding.invalidateFormulaNow(cellIndex),
    markInputChanged: (cellIndex, count) => support.markInputChangedNow(cellIndex, count),
    markFormulaChanged: (cellIndex, count) => support.markFormulaChangedNow(cellIndex, count),
    markVolatileFormulasChanged: (count) => support.markVolatileFormulasChangedNow(count),
    syncDynamicRanges: (formulaChangedCount) => support.syncDynamicRangesNow(formulaChangedCount),
    composeMutationRoots: (changedInputCount, formulaChangedCount) =>
      support.composeMutationRootsNow(changedInputCount, formulaChangedCount),
    getChangedInputBuffer: () => support.getChangedInputBufferNow(),
    rebuildTopoRanks: () => graph.rebuildTopoRanksNow(),
    detectCycles: () => graph.detectCyclesNow(),
    recalculate: (changedRoots, kernelSyncRoots) => requireService(recalc, 'recalc').recalculateNowSync(changedRoots, kernelSyncRoots),
    reconcilePivotOutputs: (baseChanged, forceAllPivots) =>
      requireService(recalc, 'recalc').reconcilePivotOutputsNow(baseChanged, forceAllPivots),
    getBatchMutationDepth: () => args.operation.getBatchMutationDepth(),
    setBatchMutationDepth: (next) => {
      args.operation.setBatchMutationDepth(next)
    },
    flushWasmProgramSync: () => graph.flushWasmProgramSyncNow(),
    writeHydratedFormulaValue: (cellIndex, value) => {
      args.state.workbook.cellStore.flags[cellIndex] =
        (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
      args.state.workbook.cellStore.setValue(cellIndex, value, value.tag === ValueTag.String ? args.state.strings.intern(value.value) : 0)
      args.state.workbook.notifyCellValueWritten(cellIndex)
    },
  })
  operations = createEngineOperationService({
    ...args.operation,
    getSelectionState: () => runEngineEffect(selection.getSelectionState()),
    setSelection: (sheetName, address) => runEngineEffect(selection.setSelection(sheetName, address)),
    rewriteDefinedNamesForSheetRename: (oldSheetName, newSheetName) =>
      runEngineEffect(maintenance.rewriteDefinedNamesForSheetRename(oldSheetName, newSheetName)),
    rewriteCellFormulasForSheetRename: (oldSheetName, newSheetName, formulaChangedCount) =>
      runEngineEffect(binding.rewriteCellFormulasForSheetRename(oldSheetName, newSheetName, formulaChangedCount)),
    rebindDefinedNameDependents: (names, formulaChangedCount) => binding.rebindDefinedNameDependentsNow(names, formulaChangedCount),
    rebindTableDependents: (tableNames, formulaChangedCount) => binding.rebindTableDependentsNow(tableNames, formulaChangedCount),
    rebindFormulaCells: (candidates, formulaChangedCount) => binding.rebindFormulaCellsNow(candidates, formulaChangedCount),
    refreshRangeDependencies: (rangeIndices) => binding.refreshRangeDependenciesNow(rangeIndices),
    rebindFormulasForSheet: (sheetName, formulaChangedCount, candidates) =>
      binding.rebindFormulasForSheetNow(sheetName, formulaChangedCount, candidates),
    removeSheetRuntime: (sheetName, explicitChangedCount) => support.removeSheetRuntimeNow(sheetName, explicitChangedCount),
    applyStructuralAxisOp: (op) => runEngineEffect(structure.applyStructuralAxisOp(op)),
    clearOwnedSpill: (cellIndex) => support.clearOwnedSpillNow(cellIndex),
    clearPivotForCell: (cellIndex) => requireService(pivot, 'pivot').clearPivotForCellNow(cellIndex),
    clearOwnedPivot: (pivotRecord) => requireService(pivot, 'pivot').clearOwnedPivotNow(pivotRecord),
    materializePivot: (pivotRecord) => requireService(pivot, 'pivot').materializePivotNow(pivotRecord),
    removeFormula: (cellIndex) => binding.clearFormulaNow(cellIndex),
    bindFormula: (cellIndex, ownerSheetName, source) => binding.bindFormulaNow(cellIndex, ownerSheetName, source),
    setInvalidFormulaValue: (cellIndex) => binding.invalidateFormulaNow(cellIndex),
    beginMutationCollection: () => support.beginMutationCollectionNow(),
    markInputChanged: (cellIndex, count) => support.markInputChangedNow(cellIndex, count),
    markFormulaChanged: (cellIndex, count) => support.markFormulaChangedNow(cellIndex, count),
    markVolatileFormulasChanged: (count) => support.markVolatileFormulasChangedNow(count),
    markSpillRootsChanged: (cellIndices, count) => support.markSpillRootsChangedNow(cellIndices, count),
    markPivotRootsChanged: (cellIndices, count) => support.markPivotRootsChangedNow(cellIndices, count),
    markExplicitChanged: (cellIndex, count) => support.markExplicitChangedNow(cellIndex, count),
    composeMutationRoots: (changedInputCount, formulaChangedCount) =>
      support.composeMutationRootsNow(changedInputCount, formulaChangedCount),
    composeEventChanges: (recalculated, explicitChangedCount) => support.composeEventChangesNow(recalculated, explicitChangedCount),
    captureChangedCells: (changedCellIndices) => changeSetEmitter.captureChangedCells(changedCellIndices),
    getChangedInputBuffer: () => support.getChangedInputBufferNow(),
    ensureRecalcScratchCapacity: (size) => scratch.ensureRecalcCapacityNow(size),
    estimatePotentialNewCells: (ops) => runEngineEffect(maintenance.estimatePotentialNewCells(ops)),
    ensureCellTracked: (sheetName, address) => support.ensureCellTrackedNow(sheetName, address),
    resetMaterializedCellScratch: (expectedSize) => support.resetMaterializedCellScratchNow(expectedSize),
    syncDynamicRanges: (formulaChangedCount) => support.syncDynamicRangesNow(formulaChangedCount),
    rebuildTopoRanks: () => graph.rebuildTopoRanksNow(),
    detectCycles: () => graph.detectCyclesNow(),
    recalculate: (changedRoots, kernelSyncRoots) => requireService(recalc, 'recalc').recalculateNowSync(changedRoots, kernelSyncRoots),
    evaluateDirectFormula: (cellIndex: number) => evaluation.evaluateDirectLookupFormulaNow(cellIndex),
    reconcilePivotOutputs: (baseChanged, forceAllPivots) =>
      requireService(recalc, 'recalc').reconcilePivotOutputsNow(baseChanged, forceAllPivots),
    flushWasmProgramSync: () => graph.flushWasmProgramSyncNow(),
    getEntityDependents: (entityId) => traversal.getEntityDependentsNow(entityId),
    collectFormulaDependents: (entityId) => traversal.collectFormulaDependentsNow(entityId),
    noteExactLookupLiteralWrite: (request) => exactLookup.recordLiteralWrite(request),
    noteSortedLookupLiteralWrite: (request) => sortedLookup.recordLiteralWrite(request),
    invalidateExactLookupColumn: (request) => exactLookup.invalidateColumn(request),
    invalidateSortedLookupColumn: (request) => sortedLookup.invalidateColumn(request),
  })
  const mutation = createEngineMutationService({
    state: args.state,
    captureSheetCellState: (sheetName) => runEngineEffect(maintenance.captureSheetCellState(sheetName)),
    captureRowRangeCellState: (sheetName, start, count) => runEngineEffect(maintenance.captureRowRangeCellState(sheetName, start, count)),
    captureColumnRangeCellState: (sheetName, start, count) =>
      runEngineEffect(maintenance.captureColumnRangeCellState(sheetName, start, count)),
    captureStoredCellOps: (cellIndex, sheetName, address) => cellState.captureStoredCellOpsNow(cellIndex, sheetName, address),
    restoreCellOps: (sheetName, address) => cellState.restoreCellOpsNow(sheetName, address),
    getCellByIndex: args.getCellByIndex,
    readRangeCells: (range) => cellState.readRangeCellsNow(range),
    toCellStateOps: (sheetName, address, snapshot, sourceSheetName, sourceAddress) =>
      cellState.toCellStateOpsNow(sheetName, address, snapshot, sourceSheetName, sourceAddress),
    applyBatchNow: (batch, source, potentialNewCells, preparedCellAddressesByOpIndex) =>
      runEngineEffect(operations.applyBatch(batch, source, potentialNewCells, preparedCellAddressesByOpIndex)),
    applyCellMutationsAtBatchNow: (refs, batch, source, potentialNewCells) =>
      runEngineEffect(operations.applyCellMutationsAt(refs, batch, source, potentialNewCells)),
    hasExternallyVisibleLocalMutationObservers: () =>
      args.state.events.hasListeners() ||
      args.state.events.hasTrackedListeners() ||
      args.state.events.hasCellListeners() ||
      (args.state.batchListeners?.size ?? 0) > 0 ||
      args.state.getSyncClientConnection() !== null,
  })
  const history = createEngineHistoryService({
    state: args.state,
    executeTransaction: (transaction, source) => runEngineEffect(mutation.executeTransaction(transaction, source)),
  })
  pivot = createEnginePivotService({
    ...args.pivot,
    ensureCellTrackedByCoords: (sheetId, row, col) => support.ensureCellTrackedByCoordsNow(sheetId, row, col),
    forEachSheetCell: (sheetId, fn) => traversal.forEachSheetCellNow(sheetId, fn),
    flushDeferredKernelSync: () => {
      const deferredCount = scratch.getDeferredKernelSyncCountNow()
      if (deferredCount === 0 || !args.state.wasm.ready) {
        return
      }
      args.state.wasm.syncFromStore(args.state.workbook.cellStore, scratch.getPendingKernelSyncNow().subarray(0, deferredCount))
      scratch.setDeferredKernelSyncCountNow(0)
      let nextEpoch = scratch.getDeferredKernelSyncEpochNow() + 1
      if (nextEpoch === 0xffff_ffff) {
        nextEpoch = 1
        scratch.getDeferredKernelSyncSeenNow().fill(0)
      }
      scratch.setDeferredKernelSyncEpochNow(nextEpoch)
    },
    scheduleWasmProgramSync: () => graph.scheduleWasmProgramSyncNow(),
    flushWasmProgramSync: () => graph.flushWasmProgramSyncNow(),
    applyDerivedOp: (op) => runEngineEffect(operations.applyDerivedOp(op)),
  })
  const snapshot = createEngineSnapshotService({
    state: args.state,
    getCellByIndex: args.getCellByIndex,
    resetWorkbook: (workbookName) => runEngineEffect(maintenance.resetWorkbook(workbookName)),
    executeRestoreTransaction: (transaction) => runEngineEffect(mutation.executeTransaction(transaction, 'restore')),
    exportTemplateBank: () => formulaTemplates.listTemplates(),
    exportFormulaInstances: () => formulaInstances.list(),
    hydrateTemplateBank: (templates) => formulaTemplates.hydrateTemplates(templates),
    resolveTemplateById: (templateId, source, row, col) => formulaTemplates.resolveByTemplateId(templateId, source, row, col),
    initializeCellFormulasAt: (refs, potentialNewCells) =>
      requireService(formulaInitialization, 'formulaInitialization').initializeCellFormulasAtNow(refs, potentialNewCells),
    initializePreparedCellFormulasAt: (refs, potentialNewCells) =>
      requireService(formulaInitialization, 'formulaInitialization').initializePreparedCellFormulasAtNow(refs, potentialNewCells),
    initializeHydratedPreparedCellFormulasAt: (refs, potentialNewCells) =>
      requireService(formulaInitialization, 'formulaInitialization').initializeHydratedPreparedCellFormulasAtNow(refs, potentialNewCells),
    materializePivot: (pivotRecord) => requireService(pivot, 'pivot').materializePivotNow(pivotRecord),
  })
  const sync = createEngineReplicaSyncService({
    state: args.state,
    applyRemoteBatchNow: (batch) => runEngineEffect(operations.applyBatch(batch, 'remote')),
    applyRemoteSnapshot: args.applyRemoteSnapshot,
  })

  return {
    cellState,
    maintenance,
    traversal,
    events: createEngineEventService(args.state),
    evaluation,
    selection,
    binding,
    formulaInitialization,
    graph,
    history,
    support,
    read,
    recalc,
    structure,
    mutation,
    operations,
    pivot,
    snapshot,
    sync,
  }
}

export function runEngineEffect<Success, Failure>(effect: Effect.Effect<Success, Failure>): Success {
  const exit = Effect.runSyncExit(effect)
  if (Exit.isSuccess(exit)) {
    return exit.value
  }
  throw Cause.squash(exit.cause)
}

export async function runEngineEffectPromise<Success, Failure>(effect: Effect.Effect<Success, Failure>): Promise<Success> {
  const exit = await Effect.runPromiseExit(effect)
  if (Exit.isSuccess(exit)) {
    return exit.value
  }
  throw Cause.squash(exit.cause)
}
