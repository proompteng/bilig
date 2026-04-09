import { Effect, Exit, Cause } from "effect";
import type { CellSnapshot } from "@bilig/protocol";
import type { EngineRuntimeState } from "./runtime-state.js";
import {
  createEngineCellStateService,
  type EngineCellStateService,
} from "./services/cell-state-service.js";
import {
  createEngineEventService,
  type EngineEventService,
} from "./services/event-service.js";
import {
  createEngineFormulaEvaluationService,
  type EngineFormulaEvaluationService,
} from "./services/formula-evaluation-service.js";
import {
  createEngineFormulaBindingService,
  type EngineFormulaBindingService,
} from "./services/formula-binding-service.js";
import {
  createEngineFormulaGraphService,
  type EngineFormulaGraphService,
} from "./services/formula-graph-service.js";
import {
  createEngineHistoryService,
  type EngineHistoryService,
} from "./services/history-service.js";
import {
  createEngineMutationService,
  type EngineMutationService,
} from "./services/mutation-service.js";
import {
  createEngineMutationSupportService,
  type EngineMutationSupportService,
} from "./services/mutation-support-service.js";
import {
  createEngineOperationService,
  type EngineOperationService,
} from "./services/operation-service.js";
import {
  createEnginePivotService,
  type EnginePivotService,
} from "./services/pivot-service.js";
import {
  createEngineReplicaSyncService,
  type EngineReplicaSyncService,
} from "./services/replica-sync-service.js";
import {
  createEngineReadService,
  type EngineReadService,
} from "./services/read-service.js";
import {
  createEngineRecalcService,
  type EngineRecalcService,
} from "./services/recalc-service.js";
import {
  createEngineSelectionService,
  type EngineSelectionService,
} from "./services/selection-service.js";
import {
  createEngineSnapshotService,
  type EngineSnapshotService,
} from "./services/snapshot-service.js";
import {
  createEngineStructureService,
  type EngineStructureService,
} from "./services/structure-service.js";
import {
  createEngineTraversalService,
  type EngineTraversalService,
} from "./services/traversal-service.js";

export interface EngineServiceRuntime {
  readonly cellState: EngineCellStateService;
  readonly traversal: EngineTraversalService;
  readonly events: EngineEventService;
  readonly evaluation: EngineFormulaEvaluationService;
  readonly selection: EngineSelectionService;
  readonly binding: EngineFormulaBindingService;
  readonly graph: EngineFormulaGraphService;
  readonly history: EngineHistoryService;
  readonly mutation: EngineMutationService;
  readonly support: EngineMutationSupportService;
  readonly operations: EngineOperationService;
  readonly pivot: EnginePivotService;
  readonly read: EngineReadService;
  readonly recalc: EngineRecalcService;
  readonly structure: EngineStructureService;
  readonly snapshot: EngineSnapshotService;
  readonly sync: EngineReplicaSyncService;
}

type EngineMutationSupportRuntimeConfig = Omit<
  Parameters<typeof createEngineMutationSupportService>[0],
  | "removeFormula"
  | "rebindFormulasForSheet"
  | "applyDerivedOp"
  | "scheduleWasmProgramSync"
  | "collectFormulaDependents"
>;

type EngineFormulaBindingRuntimeConfig = Omit<
  Parameters<typeof createEngineFormulaBindingService>[0],
  | "ensureCellTracked"
  | "ensureCellTrackedByCoords"
  | "markFormulaChanged"
  | "forEachSheetCell"
  | "resolveStructuredReference"
  | "resolveSpillReference"
  | "scheduleWasmProgramSync"
>;

type EngineFormulaGraphRuntimeConfig = Omit<
  Parameters<typeof createEngineFormulaGraphService>[0],
  "forEachFormulaDependencyCell" | "collectFormulaDependents"
>;

type EngineRecalcRuntimeConfig = Omit<
  Parameters<typeof createEngineRecalcService>[0],
  | "beginMutationCollection"
  | "markInputChanged"
  | "markFormulaChanged"
  | "markExplicitChanged"
  | "composeMutationRoots"
  | "composeEventChanges"
  | "unionChangedSets"
  | "composeChangedRootsAndOrdered"
  | "emptyChangedSet"
  | "getChangedInputBuffer"
  | "getEntityDependents"
  | "materializeSpill"
  | "clearOwnedSpill"
  | "evaluateUnsupportedFormula"
  | "materializePivot"
>;

type EnginePivotRuntimeConfig = Omit<
  Parameters<typeof createEnginePivotService>[0],
  | "ensureCellTrackedByCoords"
  | "forEachSheetCell"
  | "scheduleWasmProgramSync"
  | "flushWasmProgramSync"
  | "applyDerivedOp"
>;

type EngineOperationRuntimeConfig = Omit<
  Parameters<typeof createEngineOperationService>[0],
  | "getSelectionState"
  | "setSelection"
  | "rewriteCellFormulasForSheetRename"
  | "rebindDefinedNameDependents"
  | "rebindTableDependents"
  | "rebindFormulaCells"
  | "rebindFormulasForSheet"
  | "removeSheetRuntime"
  | "applyStructuralAxisOp"
  | "clearOwnedSpill"
  | "clearPivotForCell"
  | "clearOwnedPivot"
  | "removeFormula"
  | "bindFormula"
  | "setInvalidFormulaValue"
  | "beginMutationCollection"
  | "markInputChanged"
  | "markFormulaChanged"
  | "markVolatileFormulasChanged"
  | "markSpillRootsChanged"
  | "markPivotRootsChanged"
  | "markExplicitChanged"
  | "composeMutationRoots"
  | "composeEventChanges"
  | "getChangedInputBuffer"
  | "ensureCellTracked"
  | "resetMaterializedCellScratch"
  | "syncDynamicRanges"
  | "rebuildTopoRanks"
  | "detectCycles"
  | "recalculate"
  | "reconcilePivotOutputs"
  | "flushWasmProgramSync"
>;

type EngineTraversalRuntimeConfig = Parameters<typeof createEngineTraversalService>[0];

function requireService<Service>(service: Service | undefined, name: string): Service {
  if (service === undefined) {
    throw new Error(`Engine service ${name} is not initialized`);
  }
  return service;
}

export function createEngineServiceRuntime(args: {
  readonly state: EngineRuntimeState;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
  readonly exportSnapshot: () => import("@bilig/protocol").WorkbookSnapshot;
  readonly importSnapshot: (snapshot: import("@bilig/protocol").WorkbookSnapshot) => void;
  readonly resetWorkbook: () => void;
  readonly mutationSupport: EngineMutationSupportRuntimeConfig;
  readonly captureSheetCellState: (
    sheetName: string,
  ) => import("@bilig/workbook-domain").EngineOp[];
  readonly captureRowRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => import("@bilig/workbook-domain").EngineOp[];
  readonly captureColumnRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => import("@bilig/workbook-domain").EngineOp[];
  readonly formulaBinding: EngineFormulaBindingRuntimeConfig;
  readonly formulaGraph: EngineFormulaGraphRuntimeConfig;
  readonly recalc: EngineRecalcRuntimeConfig;
  readonly pivot: EnginePivotRuntimeConfig;
  readonly operation: EngineOperationRuntimeConfig;
  readonly traversal: EngineTraversalRuntimeConfig;
  readonly cellToCsvValue: (cell: CellSnapshot) => string;
  readonly serializeCsv: (rows: string[][]) => string;
  readonly pivotState: {
    readonly pivotOutputOwners: Map<number, string>;
  };
  readonly applyRemoteSnapshot: (snapshot: import("@bilig/protocol").WorkbookSnapshot) => void;
}): EngineServiceRuntime {
  const traversal = createEngineTraversalService(args.traversal);
  const graph = createEngineFormulaGraphService({
    ...args.formulaGraph,
    forEachFormulaDependencyCell: (cellIndex, fn) =>
      traversal.forEachFormulaDependencyCellNow(cellIndex, fn),
    collectFormulaDependents: (entityId) => traversal.collectFormulaDependentsNow(entityId),
  });
  let binding: EngineFormulaBindingService | undefined;
  let operations: EngineOperationService | undefined;
  let pivot: EnginePivotService | undefined;
  let recalc: EngineRecalcService | undefined;
  const selection = createEngineSelectionService(args.state);
  const support = createEngineMutationSupportService({
    ...args.mutationSupport,
    removeFormula: (cellIndex) => runEngineEffect(requireService(binding, "binding").clearFormula(cellIndex)),
    rebindFormulasForSheet: (sheetName, formulaChangedCount, candidates) =>
      runEngineEffect(
        requireService(binding, "binding").rebindFormulasForSheet(
          sheetName,
          formulaChangedCount,
          candidates,
        ),
      ),
    applyDerivedOp: (op) =>
      runEngineEffect(requireService(operations, "operations").applyDerivedOp(op)),
    collectFormulaDependents: (entityId) => traversal.collectFormulaDependentsNow(entityId),
    scheduleWasmProgramSync: () => runEngineEffect(graph.scheduleWasmProgramSync()),
  });
  const evaluation = createEngineFormulaEvaluationService({
    state: args.state,
    materializeSpill: (cellIndex, arrayValue) =>
      runEngineEffect(support.materializeSpill(cellIndex, arrayValue)),
    clearOwnedSpill: (cellIndex) => runEngineEffect(support.clearOwnedSpill(cellIndex)),
    resolvePivotData: (sheetName, address, dataField, filters) =>
      runEngineEffect(
        requireService(pivot, "pivot").resolvePivotData(sheetName, address, dataField, filters),
      ),
  });
  binding = createEngineFormulaBindingService({
    ...args.formulaBinding,
    ensureCellTracked: (sheetName, address) =>
      runEngineEffect(support.ensureCellTracked(sheetName, address)),
    ensureCellTrackedByCoords: (sheetId, row, col) =>
      runEngineEffect(support.ensureCellTrackedByCoords(sheetId, row, col)),
    markFormulaChanged: (cellIndex, count) =>
      runEngineEffect(support.markFormulaChanged(cellIndex, count)),
    resolveStructuredReference: (tableName, columnName) =>
      runEngineEffect(evaluation.resolveStructuredReference(tableName, columnName)),
    resolveSpillReference: (currentSheetName, sheetName, address) =>
      runEngineEffect(evaluation.resolveSpillReference(currentSheetName, sheetName, address)),
    forEachSheetCell: (sheetId, fn) => traversal.forEachSheetCellNow(sheetId, fn),
    scheduleWasmProgramSync: () => runEngineEffect(graph.scheduleWasmProgramSync()),
  });
  const read = createEngineReadService({
    state: args.state,
    forEachFormulaDependencyCell: (cellIndex, fn) =>
      traversal.forEachFormulaDependencyCellNow(cellIndex, fn),
    getEntityDependents: (entityId) => traversal.getEntityDependentsNow(entityId),
    cellToCsvValue: args.cellToCsvValue,
    serializeCsv: args.serializeCsv,
  });
  const cellState = createEngineCellStateService({
    state: args.state,
    getCell: (sheetName, address) => runEngineEffect(read.getCell(sheetName, address)),
    getCellByIndex: (cellIndex) => runEngineEffect(read.getCellByIndex(cellIndex)),
  });
  const structure = createEngineStructureService({
    state: {
      workbook: args.state.workbook,
      formulas: args.state.formulas,
      pivotOutputOwners: args.pivotState.pivotOutputOwners,
    },
    getCellByIndex: (cellIndex) => runEngineEffect(read.getCellByIndex(cellIndex)),
    toCellStateOps: (sheetName, address, snapshot, sourceSheetName, sourceAddress) =>
      runEngineEffect(
        cellState.toCellStateOps(
          sheetName,
          address,
          snapshot,
          sourceSheetName,
          sourceAddress,
        ),
      ),
    removeFormula: (cellIndex) => runEngineEffect(binding.clearFormula(cellIndex)),
    clearOwnedPivot: (pivotRecord) =>
      runEngineEffect(requireService(pivot, "pivot").clearOwnedPivot(pivotRecord)),
    rebuildAllFormulaBindings: () => runEngineEffect(binding.rebuildAllFormulaBindings()),
  });
  recalc = createEngineRecalcService({
    ...args.recalc,
    beginMutationCollection: () => runEngineEffect(support.beginMutationCollection()),
    markInputChanged: (cellIndex, count) => runEngineEffect(support.markInputChanged(cellIndex, count)),
    markFormulaChanged: (cellIndex, count) =>
      runEngineEffect(support.markFormulaChanged(cellIndex, count)),
    markExplicitChanged: (cellIndex, count) =>
      runEngineEffect(support.markExplicitChanged(cellIndex, count)),
    composeMutationRoots: (changedInputCount, formulaChangedCount) =>
      runEngineEffect(support.composeMutationRoots(changedInputCount, formulaChangedCount)),
    composeEventChanges: (recalculated, explicitChangedCount) =>
      runEngineEffect(support.composeEventChanges(recalculated, explicitChangedCount)),
    unionChangedSets: (...sets) => runEngineEffect(support.unionChangedSets(...sets)),
    composeChangedRootsAndOrdered: (changedRoots, ordered, orderedCount) =>
      runEngineEffect(support.composeChangedRootsAndOrdered(changedRoots, ordered, orderedCount)),
    emptyChangedSet: () => runEngineEffect(support.unionChangedSets()),
    getChangedInputBuffer: () => runEngineEffect(support.getChangedInputBuffer()),
    materializeSpill: (cellIndex, arrayValue) =>
      runEngineEffect(support.materializeSpill(cellIndex, arrayValue)),
    clearOwnedSpill: (cellIndex) => runEngineEffect(support.clearOwnedSpill(cellIndex)),
    evaluateUnsupportedFormula: (cellIndex) =>
      runEngineEffect(evaluation.evaluateUnsupportedFormula(cellIndex)),
    materializePivot: (pivotRecord) =>
      runEngineEffect(requireService(pivot, "pivot").materializePivot(pivotRecord)),
    getEntityDependents: (entityId) => traversal.getEntityDependentsNow(entityId),
  });
  operations = createEngineOperationService({
    ...args.operation,
    getSelectionState: () => runEngineEffect(selection.getSelectionState()),
    setSelection: (sheetName, address) =>
      runEngineEffect(selection.setSelection(sheetName, address)),
    rewriteCellFormulasForSheetRename: (oldSheetName, newSheetName, formulaChangedCount) =>
      runEngineEffect(
        binding.rewriteCellFormulasForSheetRename(
          oldSheetName,
          newSheetName,
          formulaChangedCount,
        ),
      ),
    rebindDefinedNameDependents: (names, formulaChangedCount) =>
      runEngineEffect(binding.rebindDefinedNameDependents(names, formulaChangedCount)),
    rebindTableDependents: (tableNames, formulaChangedCount) =>
      runEngineEffect(binding.rebindTableDependents(tableNames, formulaChangedCount)),
    rebindFormulaCells: (candidates, formulaChangedCount) =>
      runEngineEffect(binding.rebindFormulaCells(candidates, formulaChangedCount)),
    rebindFormulasForSheet: (sheetName, formulaChangedCount, candidates) =>
      runEngineEffect(binding.rebindFormulasForSheet(sheetName, formulaChangedCount, candidates)),
    removeSheetRuntime: (sheetName, explicitChangedCount) =>
      runEngineEffect(support.removeSheetRuntime(sheetName, explicitChangedCount)),
    applyStructuralAxisOp: (op) => runEngineEffect(structure.applyStructuralAxisOp(op)),
    clearOwnedSpill: (cellIndex) => runEngineEffect(support.clearOwnedSpill(cellIndex)),
    clearPivotForCell: (cellIndex) =>
      runEngineEffect(requireService(pivot, "pivot").clearPivotForCell(cellIndex)),
    clearOwnedPivot: (pivotRecord) =>
      runEngineEffect(requireService(pivot, "pivot").clearOwnedPivot(pivotRecord)),
    removeFormula: (cellIndex) => runEngineEffect(binding.clearFormula(cellIndex)),
    bindFormula: (cellIndex, ownerSheetName, source) =>
      runEngineEffect(binding.bindFormula(cellIndex, ownerSheetName, source)),
    setInvalidFormulaValue: (cellIndex) => runEngineEffect(binding.invalidateFormula(cellIndex)),
    beginMutationCollection: () => runEngineEffect(support.beginMutationCollection()),
    markInputChanged: (cellIndex, count) => runEngineEffect(support.markInputChanged(cellIndex, count)),
    markFormulaChanged: (cellIndex, count) =>
      runEngineEffect(support.markFormulaChanged(cellIndex, count)),
    markVolatileFormulasChanged: (count) =>
      runEngineEffect(support.markVolatileFormulasChanged(count)),
    markSpillRootsChanged: (cellIndices, count) =>
      runEngineEffect(support.markSpillRootsChanged(cellIndices, count)),
    markPivotRootsChanged: (cellIndices, count) =>
      runEngineEffect(support.markPivotRootsChanged(cellIndices, count)),
    markExplicitChanged: (cellIndex, count) =>
      runEngineEffect(support.markExplicitChanged(cellIndex, count)),
    composeMutationRoots: (changedInputCount, formulaChangedCount) =>
      runEngineEffect(support.composeMutationRoots(changedInputCount, formulaChangedCount)),
    composeEventChanges: (recalculated, explicitChangedCount) =>
      runEngineEffect(support.composeEventChanges(recalculated, explicitChangedCount)),
    getChangedInputBuffer: () => runEngineEffect(support.getChangedInputBuffer()),
    ensureCellTracked: (sheetName, address) =>
      runEngineEffect(support.ensureCellTracked(sheetName, address)),
    resetMaterializedCellScratch: (expectedSize) =>
      runEngineEffect(support.resetMaterializedCellScratch(expectedSize)),
    syncDynamicRanges: (formulaChangedCount) =>
      runEngineEffect(support.syncDynamicRanges(formulaChangedCount)),
    rebuildTopoRanks: () => runEngineEffect(graph.rebuildTopoRanks()),
    detectCycles: () => runEngineEffect(graph.detectCycles()),
    recalculate: (changedRoots, kernelSyncRoots) =>
      runEngineEffect(requireService(recalc, "recalc").recalculate(changedRoots, kernelSyncRoots)),
    reconcilePivotOutputs: (baseChanged, forceAllPivots) =>
      runEngineEffect(
        requireService(recalc, "recalc").reconcilePivotOutputs(baseChanged, forceAllPivots),
      ),
    flushWasmProgramSync: () => runEngineEffect(graph.flushWasmProgramSync()),
  });
  const mutation = createEngineMutationService({
    state: args.state,
    captureSheetCellState: args.captureSheetCellState,
    captureRowRangeCellState: args.captureRowRangeCellState,
    captureColumnRangeCellState: args.captureColumnRangeCellState,
    restoreCellOps: (sheetName, address) =>
      runEngineEffect(cellState.restoreCellOps(sheetName, address)),
    readRangeCells: (range) => runEngineEffect(cellState.readRangeCells(range)),
    toCellStateOps: (sheetName, address, snapshot, sourceSheetName, sourceAddress) =>
      runEngineEffect(
        cellState.toCellStateOps(
          sheetName,
          address,
          snapshot,
          sourceSheetName,
          sourceAddress,
        ),
      ),
    applyBatchNow: (batch, source, potentialNewCells) =>
      runEngineEffect(operations.applyBatch(batch, source, potentialNewCells)),
  });
  const history = createEngineHistoryService({
    state: args.state,
    executeTransaction: (transaction, source) =>
      runEngineEffect(mutation.executeTransaction(transaction, source)),
  });
  pivot = createEnginePivotService({
    ...args.pivot,
    ensureCellTrackedByCoords: (sheetId, row, col) =>
      runEngineEffect(support.ensureCellTrackedByCoords(sheetId, row, col)),
    forEachSheetCell: (sheetId, fn) => traversal.forEachSheetCellNow(sheetId, fn),
    scheduleWasmProgramSync: () => runEngineEffect(graph.scheduleWasmProgramSync()),
    flushWasmProgramSync: () => runEngineEffect(graph.flushWasmProgramSync()),
    applyDerivedOp: (op) => runEngineEffect(operations.applyDerivedOp(op)),
  });
  const snapshot = createEngineSnapshotService({
    state: args.state,
    getCellByIndex: args.getCellByIndex,
    resetWorkbook: args.resetWorkbook,
    executeRestoreTransaction: (transaction) =>
      runEngineEffect(mutation.executeTransaction(transaction, "restore")),
  });
  const sync = createEngineReplicaSyncService({
    state: args.state,
    applyRemoteBatchNow: (batch) => runEngineEffect(operations.applyBatch(batch, "remote")),
    applyRemoteSnapshot: args.applyRemoteSnapshot,
  });

  return {
    cellState,
    traversal,
    events: createEngineEventService(args.state),
    evaluation,
    selection,
    binding,
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
  };
}

export function runEngineEffect<Success, Failure>(
  effect: Effect.Effect<Success, Failure>,
): Success {
  const exit = Effect.runSyncExit(effect);
  if (Exit.isSuccess(exit)) {
    return exit.value;
  }
  throw Cause.squash(exit.cause);
}

export async function runEngineEffectPromise<Success, Failure>(
  effect: Effect.Effect<Success, Failure>,
): Promise<Success> {
  const exit = await Effect.runPromiseExit(effect);
  if (Exit.isSuccess(exit)) {
    return exit.value;
  }
  throw Cause.squash(exit.cause);
}
