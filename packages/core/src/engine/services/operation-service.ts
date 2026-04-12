import { Effect } from "effect";
import { formatAddress } from "@bilig/formula";
import type { EngineOp, EngineOpBatch } from "@bilig/workbook-domain";
import type { CellRangeRef, EngineEvent } from "@bilig/protocol";
import type { EngineCellMutationRef } from "../../cell-mutations-at.js";
import { makeCellEntity } from "../../entity-ids.js";
import {
  batchOpOrder,
  compareOpOrder,
  createBatch,
  markBatchApplied,
  type OpOrder,
} from "../../replica-state.js";
import { CellFlags } from "../../cell-store.js";
import { emptyValue, writeLiteralToCellStore } from "../../engine-value-utils.js";
import { spillDependencyKey, tableDependencyKey } from "../../engine-metadata-utils.js";
import {
  makeCellKey,
  normalizeDefinedName,
  pivotKey,
  type WorkbookPivotRecord,
} from "../../workbook-store.js";
import type { EngineRuntimeState, PreparedCellAddress, U32 } from "../runtime-state.js";
import { EngineMutationError } from "../errors.js";

type MutationSource = "local" | "remote" | "restore" | "history";

type StructuralAxisOp = Extract<
  EngineOp,
  {
    kind:
      | "insertRows"
      | "deleteRows"
      | "moveRows"
      | "insertColumns"
      | "deleteColumns"
      | "moveColumns";
  }
>;

type DerivedOp = Extract<
  EngineOp,
  { kind: "upsertSpillRange" | "deleteSpillRange" | "upsertPivotTable" | "deletePivotTable" }
>;

export interface EngineOperationService {
  readonly applyBatch: (
    batch: EngineOpBatch,
    source: MutationSource,
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
  ) => Effect.Effect<void, EngineMutationError>;
  readonly applyCellMutationsAt: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch,
    source: "local" | "restore",
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>;
  readonly applyDerivedOp: (op: DerivedOp) => Effect.Effect<number[], EngineMutationError>;
}

interface VersionStore {
  get(key: string): OpOrder | undefined;
  set(key: string, value: OpOrder): void;
}

const noopVersionStore: VersionStore = {
  get() {
    return undefined;
  },
  set() {
    return;
  },
};

const FAST_LITERAL_OVERWRITE_FLAGS =
  CellFlags.HasFormula |
  CellFlags.JsOnly |
  CellFlags.InCycle |
  CellFlags.SpillChild |
  CellFlags.PivotOutput;

function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

function assertNever(value: never): never {
  throw new Error(`Unexpected value: ${String(value)}`);
}

function collectTrackedDependents(
  registry: Map<string, Set<number>>,
  keys: readonly string[],
): number[] {
  const candidates = new Set<number>();
  keys.forEach((key) => {
    registry.get(key)?.forEach((cellIndex) => {
      candidates.add(cellIndex);
    });
  });
  return [...candidates];
}

export function createEngineOperationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | "workbook"
    | "strings"
    | "events"
    | "formulas"
    | "replicaState"
    | "entityVersions"
    | "sheetDeleteVersions"
    | "batchListeners"
    | "redoStack"
    | "trackReplicaVersions"
    | "getSyncClientConnection"
    | "getLastMetrics"
    | "setLastMetrics"
  >;
  readonly reverseState: {
    readonly reverseSpillEdges: Map<string, Set<number>>;
  };
  readonly getSelectionState: () => import("@bilig/protocol").SelectionState;
  readonly setSelection: (sheetName: string, address: string) => void;
  readonly rewriteDefinedNamesForSheetRename: (oldSheetName: string, newSheetName: string) => void;
  readonly rewriteCellFormulasForSheetRename: (
    oldSheetName: string,
    newSheetName: string,
    formulaChangedCount: number,
  ) => number;
  readonly rebindDefinedNameDependents: (
    names: readonly string[],
    formulaChangedCount: number,
  ) => number;
  readonly rebindTableDependents: (
    tableNames: readonly string[],
    formulaChangedCount: number,
  ) => number;
  readonly rebindFormulaCells: (
    candidates: readonly number[],
    formulaChangedCount: number,
  ) => number;
  readonly rebindFormulasForSheet: (
    sheetName: string,
    formulaChangedCount: number,
    candidates?: readonly number[] | U32,
  ) => number;
  readonly removeSheetRuntime: (
    sheetName: string,
    explicitChangedCount: number,
  ) => { changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number };
  readonly applyStructuralAxisOp: (op: StructuralAxisOp) => {
    changedCellIndices: number[];
    formulaCellIndices: number[];
  };
  readonly clearOwnedSpill: (cellIndex: number) => number[];
  readonly clearPivotForCell: (cellIndex: number) => number[];
  readonly clearOwnedPivot: (pivot: WorkbookPivotRecord) => number[];
  readonly removeFormula: (cellIndex: number) => boolean;
  readonly bindFormula: (cellIndex: number, ownerSheetName: string, source: string) => boolean;
  readonly setInvalidFormulaValue: (cellIndex: number) => void;
  readonly beginMutationCollection: () => void;
  readonly markInputChanged: (cellIndex: number, count: number) => number;
  readonly markFormulaChanged: (cellIndex: number, count: number) => number;
  readonly markVolatileFormulasChanged: (count: number) => number;
  readonly markSpillRootsChanged: (cellIndices: readonly number[], count: number) => number;
  readonly markPivotRootsChanged: (cellIndices: readonly number[], count: number) => number;
  readonly markExplicitChanged: (cellIndex: number, count: number) => number;
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32;
  readonly composeEventChanges: (recalculated: U32, explicitChangedCount: number) => U32;
  readonly getChangedInputBuffer: () => U32;
  readonly ensureCellTracked: (sheetName: string, address: string) => number;
  readonly estimatePotentialNewCells: (ops: readonly EngineOp[]) => number;
  readonly resetMaterializedCellScratch: (expectedSize: number) => void;
  readonly syncDynamicRanges: (formulaChangedCount: number) => number;
  readonly rebuildTopoRanks: () => void;
  readonly detectCycles: () => void;
  readonly recalculate: (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots?: readonly number[] | U32,
  ) => U32;
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32;
  readonly flushWasmProgramSync: () => void;
  readonly getBatchMutationDepth: () => number;
  readonly setBatchMutationDepth: (next: number) => void;
  readonly collectFormulaDependents: (entityId: number) => Uint32Array;
}): EngineOperationService {
  const emitBatch = (batch: EngineOpBatch): void => {
    args.state.batchListeners.forEach((listener) => listener(batch));
  };
  const entityVersions: VersionStore = args.state.trackReplicaVersions
    ? args.state.entityVersions
    : noopVersionStore;
  const sheetDeleteVersions: VersionStore = args.state.trackReplicaVersions
    ? args.state.sheetDeleteVersions
    : noopVersionStore;
  const setEntityVersionForOp = (op: EngineOp, order: OpOrder): void => {
    if (!args.state.trackReplicaVersions) {
      return;
    }
    entityVersions.set(entityKeyForOp(op), order);
  };
  const setCellEntityVersion = (sheetName: string, address: string, order: OpOrder): void => {
    if (!args.state.trackReplicaVersions) {
      return;
    }
    entityVersions.set(`cell:${sheetName}!${address}`, order);
  };
  const setSheetDeleteVersion = (sheetName: string, order: OpOrder): void => {
    if (!args.state.trackReplicaVersions) {
      return;
    }
    sheetDeleteVersions.set(sheetName, order);
  };

  const pruneCellIfOrphaned = (cellIndex: number): void => {
    if (args.collectFormulaDependents(makeCellEntity(cellIndex)).length > 0) {
      return;
    }
    args.state.workbook.pruneCellIfEmpty(cellIndex);
  };

  const entityKeyForOp = (op: EngineOp): string => {
    switch (op.kind) {
      case "upsertWorkbook":
        return "workbook";
      case "setWorkbookMetadata":
        return `workbook-meta:${op.key}`;
      case "setCalculationSettings":
        return "workbook-calc";
      case "setVolatileContext":
        return "workbook-volatile";
      case "upsertSheet":
      case "deleteSheet":
        return `sheet:${op.name}`;
      case "renameSheet":
        return `sheet:${op.oldName}`;
      case "insertRows":
      case "deleteRows":
      case "moveRows":
        return `row-structure:${op.sheetName}`;
      case "insertColumns":
      case "deleteColumns":
      case "moveColumns":
        return `column-structure:${op.sheetName}`;
      case "updateRowMetadata":
        return `row-meta:${op.sheetName}:${op.start}:${op.count}`;
      case "updateColumnMetadata":
        return `column-meta:${op.sheetName}:${op.start}:${op.count}`;
      case "setFreezePane":
      case "clearFreezePane":
        return `freeze:${op.sheetName}`;
      case "setFilter":
      case "clearFilter":
        return `filter:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
      case "setSort":
      case "clearSort":
        return `sort:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
      case "setCellFormat":
        return `format:${op.sheetName}!${op.address}`;
      case "upsertCellStyle":
        return `style:${op.style.id}`;
      case "upsertCellNumberFormat":
        return `number-format:${op.format.id}`;
      case "setStyleRange":
        return `style-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
      case "setFormatRange":
        return `format-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
        return `cell:${op.sheetName}!${op.address}`;
      case "upsertDefinedName":
      case "deleteDefinedName":
        return `defined-name:${normalizeDefinedName(op.name)}`;
      case "upsertTable":
        return `table:${normalizeDefinedName(op.table.name)}`;
      case "deleteTable":
        return `table:${normalizeDefinedName(op.name)}`;
      case "upsertSpillRange":
      case "deleteSpillRange":
        return `spill:${op.sheetName}!${op.address}`;
      case "upsertPivotTable":
      case "deletePivotTable":
        return `pivot:${pivotKey(op.sheetName, op.address)}`;
      default:
        return assertNever(op);
    }
  };
  const canFastPathLiteralOverwrite = (cellIndex: number): boolean => {
    const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0;
    return (
      (flags & FAST_LITERAL_OVERWRITE_FLAGS) === 0 &&
      args.state.formulas.get(cellIndex) === undefined
    );
  };

  const getPreparedExistingCellIndex = (
    sheetName: string,
    address: string,
    preparedCellAddress: PreparedCellAddress | null,
  ): number | undefined => {
    if (!preparedCellAddress) {
      return args.state.workbook.getCellIndex(sheetName, address);
    }
    const sheet = args.state.workbook.getSheet(sheetName);
    if (!sheet) {
      return undefined;
    }
    return args.state.workbook.cellKeyToIndex.get(
      makeCellKey(sheet.id, preparedCellAddress.row, preparedCellAddress.col),
    );
  };

  const ensurePreparedCellTracked = (
    sheetName: string,
    address: string,
    preparedCellAddress: PreparedCellAddress | null,
  ): number => {
    if (!preparedCellAddress) {
      return args.ensureCellTracked(sheetName, address);
    }
    const sheet = args.state.workbook.getOrCreateSheet(sheetName);
    return args.state.workbook.ensureCellAt(
      sheet.id,
      preparedCellAddress.row,
      preparedCellAddress.col,
    ).cellIndex;
  };

  const sheetDeleteBarrierForOp = (op: EngineOp): OpOrder | undefined => {
    switch (op.kind) {
      case "upsertWorkbook":
      case "setWorkbookMetadata":
      case "setCalculationSettings":
      case "setVolatileContext":
      case "deleteSheet":
      case "upsertDefinedName":
      case "deleteDefinedName":
      case "upsertTable":
      case "deleteTable":
        return undefined;
      case "updateRowMetadata":
      case "updateColumnMetadata":
      case "insertRows":
      case "deleteRows":
      case "moveRows":
      case "insertColumns":
      case "deleteColumns":
      case "moveColumns":
      case "setFreezePane":
      case "clearFreezePane":
      case "setFilter":
      case "clearFilter":
      case "setSort":
      case "clearSort":
      case "setCellFormat":
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
      case "upsertSpillRange":
      case "deleteSpillRange":
      case "deletePivotTable":
        return sheetDeleteVersions.get(op.sheetName);
      case "setStyleRange":
      case "setFormatRange":
        return sheetDeleteVersions.get(op.range.sheetName);
      case "upsertCellNumberFormat":
      case "upsertCellStyle":
        return undefined;
      case "upsertSheet":
        return sheetDeleteVersions.get(op.name);
      case "renameSheet":
        return sheetDeleteVersions.get(op.oldName);
      case "upsertPivotTable":
        return (
          sheetDeleteVersions.get(op.sheetName) ?? sheetDeleteVersions.get(op.source.sheetName)
        );
      default:
        return assertNever(op);
    }
  };

  const shouldApplyOp = (op: EngineOp, order: OpOrder): boolean => {
    const sheetDeleteOrder = sheetDeleteBarrierForOp(op);
    if (sheetDeleteOrder && compareOpOrder(order, sheetDeleteOrder) <= 0) {
      return false;
    }
    const existingOrder = entityVersions.get(entityKeyForOp(op));
    if (existingOrder && compareOpOrder(order, existingOrder) <= 0) {
      return false;
    }
    return true;
  };

  const applySpillRangeOp = (
    op: Extract<EngineOp, { kind: "upsertSpillRange" | "deleteSpillRange" }>,
    order: OpOrder,
  ): number[] => {
    if (op.kind === "upsertSpillRange") {
      args.state.workbook.setSpill(op.sheetName, op.address, op.rows, op.cols);
    } else {
      args.state.workbook.deleteSpill(op.sheetName, op.address);
    }
    setEntityVersionForOp(op, order);
    return collectTrackedDependents(args.reverseState.reverseSpillEdges, [
      spillDependencyKey(op.sheetName, op.address),
    ]);
  };

  const applyPivotUpsertOp = (
    op: Extract<EngineOp, { kind: "upsertPivotTable" }>,
    order: OpOrder,
  ): void => {
    args.state.workbook.setPivot({
      name: op.name,
      sheetName: op.sheetName,
      address: op.address,
      source: op.source,
      groupBy: op.groupBy,
      values: op.values,
      rows: op.rows,
      cols: op.cols,
    });
    setEntityVersionForOp(op, order);
  };

  const applyPivotDeleteOp = (
    op: Extract<EngineOp, { kind: "deletePivotTable" }>,
    order: OpOrder,
  ): number[] => {
    const pivot = args.state.workbook.getPivot(op.sheetName, op.address);
    if (!pivot) {
      setEntityVersionForOp(op, order);
      return [];
    }
    const changedPivotOutputs = args.clearOwnedPivot(pivot);
    args.state.workbook.deletePivot(op.sheetName, op.address);
    setEntityVersionForOp(op, order);
    return changedPivotOutputs;
  };

  const applyDerivedOpNow = (op: DerivedOp): number[] => {
    const batch = createBatch(args.state.replicaState, [op]);
    const order = batchOpOrder(batch, 0);
    switch (op.kind) {
      case "upsertSpillRange":
      case "deleteSpillRange": {
        const candidates = applySpillRangeOp(op, order);
        args.rebindFormulaCells(candidates, 0);
        return candidates;
      }
      case "upsertPivotTable":
        applyPivotUpsertOp(op, order);
        return [];
      case "deletePivotTable":
        return applyPivotDeleteOp(op, order);
      default:
        return assertNever(op);
    }
  };

  const applyBatchNow = (
    batch: EngineOpBatch,
    source: MutationSource,
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
  ): void => {
    if (
      preparedCellAddressesByOpIndex &&
      preparedCellAddressesByOpIndex.length !== batch.ops.length
    ) {
      throw new Error("Prepared cell addresses must align with batch operations");
    }
    const isRestore = source === "restore";
    args.beginMutationCollection();
    let changedInputCount = 0;
    let formulaChangedCount = 0;
    let explicitChangedCount = 0;
    let topologyChanged = false;
    let sheetDeleted = false;
    let structuralInvalidation = false;
    let compileMs = 0;
    const invalidatedRanges: CellRangeRef[] = [];
    const invalidatedRows: { sheetName: string; startIndex: number; endIndex: number }[] = [];
    const invalidatedColumns: { sheetName: string; startIndex: number; endIndex: number }[] = [];
    let refreshAllPivots = false;
    let appliedOps = 0;
    const canSkipOrderChecks = source !== "remote";

    const reservedNewCells = potentialNewCells ?? args.estimatePotentialNewCells(batch.ops);
    args.state.workbook.cellStore.ensureCapacity(
      args.state.workbook.cellStore.size + reservedNewCells,
    );
    args.resetMaterializedCellScratch(reservedNewCells);

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1);
    try {
      batch.ops.forEach((op, opIndex) => {
        const order = batchOpOrder(batch, opIndex);
        const preparedCellAddress = preparedCellAddressesByOpIndex?.[opIndex] ?? null;
        if (!canSkipOrderChecks && !shouldApplyOp(op, order)) {
          return;
        }

        switch (op.kind) {
          case "upsertWorkbook":
            args.state.workbook.workbookName = op.name;
            setEntityVersionForOp(op, order);
            break;
          case "setWorkbookMetadata":
            args.state.workbook.setWorkbookProperty(op.key, op.value);
            setEntityVersionForOp(op, order);
            break;
          case "setCalculationSettings":
            args.state.workbook.setCalculationSettings(op.settings);
            setEntityVersionForOp(op, order);
            break;
          case "setVolatileContext":
            args.state.workbook.setVolatileContext(op.context);
            setEntityVersionForOp(op, order);
            break;
          case "upsertSheet": {
            args.state.workbook.createSheet(op.name, op.order, op.id);
            setEntityVersionForOp(op, order);
            const tombstone = sheetDeleteVersions.get(op.name);
            if (!tombstone || compareOpOrder(order, tombstone) > 0) {
              args.state.sheetDeleteVersions.delete(op.name);
            }
            const reboundCount = formulaChangedCount;
            formulaChangedCount = args.rebindFormulasForSheet(op.name, formulaChangedCount);
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
            refreshAllPivots = true;
            break;
          }
          case "renameSheet": {
            const renamedSheet = args.state.workbook.renameSheet(op.oldName, op.newName);
            if (args.state.trackReplicaVersions) {
              entityVersions.set(`sheet:${op.oldName}`, order);
              entityVersions.set(`sheet:${op.newName}`, order);
            }
            setSheetDeleteVersion(op.oldName, order);
            const renamedTombstone = sheetDeleteVersions.get(op.newName);
            if (!renamedTombstone || compareOpOrder(order, renamedTombstone) > 0) {
              args.state.sheetDeleteVersions.delete(op.newName);
            }
            if (!renamedSheet) {
              break;
            }
            const selection = args.getSelectionState();
            if (selection.sheetName === op.oldName) {
              args.setSelection(op.newName, selection.address ?? "A1");
            }
            args.rewriteDefinedNamesForSheetRename(op.oldName, op.newName);
            formulaChangedCount = args.rewriteCellFormulasForSheetRename(
              op.oldName,
              op.newName,
              formulaChangedCount,
            );
            topologyChanged = true;
            sheetDeleted = true;
            structuralInvalidation = true;
            refreshAllPivots = true;
            break;
          }
          case "deleteSheet": {
            const removal = args.removeSheetRuntime(op.name, explicitChangedCount);
            changedInputCount += removal.changedInputCount;
            formulaChangedCount += removal.formulaChangedCount;
            explicitChangedCount = removal.explicitChangedCount;
            setEntityVersionForOp(op, order);
            setSheetDeleteVersion(op.name, order);
            topologyChanged = true;
            sheetDeleted = true;
            structuralInvalidation = true;
            refreshAllPivots = true;
            break;
          }
          case "insertRows":
          case "deleteRows":
          case "moveRows":
          case "insertColumns":
          case "deleteColumns":
          case "moveColumns": {
            const structural = args.applyStructuralAxisOp(op);
            structural.changedCellIndices.forEach((cellIndex) => {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
            });
            structural.formulaCellIndices.forEach((cellIndex) => {
              formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
            });
            topologyChanged = true;
            structuralInvalidation = true;
            refreshAllPivots = true;
            setEntityVersionForOp(op, order);
            break;
          }
          case "updateRowMetadata":
            args.state.workbook.setRowMetadata(
              op.sheetName,
              op.start,
              op.count,
              op.size,
              op.hidden,
            );
            invalidatedRows.push({
              sheetName: op.sheetName,
              startIndex: op.start,
              endIndex: op.start + op.count - 1,
            });
            setEntityVersionForOp(op, order);
            break;
          case "updateColumnMetadata":
            args.state.workbook.setColumnMetadata(
              op.sheetName,
              op.start,
              op.count,
              op.size,
              op.hidden,
            );
            invalidatedColumns.push({
              sheetName: op.sheetName,
              startIndex: op.start,
              endIndex: op.start + op.count - 1,
            });
            setEntityVersionForOp(op, order);
            break;
          case "setFreezePane":
            args.state.workbook.setFreezePane(op.sheetName, op.rows, op.cols);
            structuralInvalidation = true;
            setEntityVersionForOp(op, order);
            break;
          case "clearFreezePane":
            args.state.workbook.clearFreezePane(op.sheetName);
            structuralInvalidation = true;
            setEntityVersionForOp(op, order);
            break;
          case "setFilter":
            args.state.workbook.setFilter(op.sheetName, op.range);
            structuralInvalidation = true;
            setEntityVersionForOp(op, order);
            break;
          case "clearFilter":
            args.state.workbook.deleteFilter(op.sheetName, op.range);
            structuralInvalidation = true;
            setEntityVersionForOp(op, order);
            break;
          case "setSort":
            args.state.workbook.setSort(op.sheetName, op.range, op.keys);
            structuralInvalidation = true;
            setEntityVersionForOp(op, order);
            break;
          case "clearSort":
            args.state.workbook.deleteSort(op.sheetName, op.range);
            structuralInvalidation = true;
            setEntityVersionForOp(op, order);
            break;
          case "upsertTable": {
            args.state.workbook.setTable(op.table);
            const reboundCount = formulaChangedCount;
            formulaChangedCount = args.rebindTableDependents(
              [tableDependencyKey(op.table.name)],
              formulaChangedCount,
            );
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
            setEntityVersionForOp(op, order);
            break;
          }
          case "deleteTable": {
            args.state.workbook.deleteTable(op.name);
            const reboundCount = formulaChangedCount;
            formulaChangedCount = args.rebindTableDependents(
              [tableDependencyKey(op.name)],
              formulaChangedCount,
            );
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
            setEntityVersionForOp(op, order);
            break;
          }
          case "upsertSpillRange":
          case "deleteSpillRange": {
            const reboundCount = formulaChangedCount;
            formulaChangedCount = args.rebindFormulaCells(
              applySpillRangeOp(op, order),
              formulaChangedCount,
            );
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
            break;
          }
          case "setCellValue": {
            if (!isRestore) {
              const existingIndex = getPreparedExistingCellIndex(
                op.sheetName,
                op.address,
                preparedCellAddress,
              );
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(
                  args.clearPivotForCell(existingIndex),
                  changedInputCount,
                );
              }
            }
            const cellIndex = ensurePreparedCellTracked(
              op.sheetName,
              op.address,
              preparedCellAddress,
            );
            if (!isRestore) {
              changedInputCount = args.markSpillRootsChanged(
                args.clearOwnedSpill(cellIndex),
                changedInputCount,
              );
              topologyChanged = args.removeFormula(cellIndex) || topologyChanged;
            }
            writeLiteralToCellStore(
              args.state.workbook.cellStore,
              cellIndex,
              op.value,
              args.state.strings,
            );
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(
                CellFlags.HasFormula |
                CellFlags.JsOnly |
                CellFlags.InCycle |
                CellFlags.SpillChild |
                CellFlags.PivotOutput
              );
            if (!isRestore && op.value === null) {
              pruneCellIfOrphaned(cellIndex);
            }
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount);
            if (!isRestore) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
              setEntityVersionForOp(op, order);
            }
            break;
          }
          case "setCellFormula": {
            if (!isRestore) {
              const existingIndex = getPreparedExistingCellIndex(
                op.sheetName,
                op.address,
                preparedCellAddress,
              );
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(
                  args.clearPivotForCell(existingIndex),
                  changedInputCount,
                );
              }
            }
            const cellIndex = ensurePreparedCellTracked(
              op.sheetName,
              op.address,
              preparedCellAddress,
            );
            if (!isRestore) {
              changedInputCount = args.markSpillRootsChanged(
                args.clearOwnedSpill(cellIndex),
                changedInputCount,
              );
            }
            const compileStarted = isRestore ? 0 : performance.now();
            try {
              const changedTopology = args.bindFormula(cellIndex, op.sheetName, op.formula);
              if (!isRestore) {
                compileMs += performance.now() - compileStarted;
              }
              formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
              topologyChanged = topologyChanged || changedTopology;
            } catch {
              if (!isRestore) {
                compileMs += performance.now() - compileStarted;
              }
              topologyChanged = args.removeFormula(cellIndex) || topologyChanged;
              args.setInvalidFormulaValue(cellIndex);
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount);
            }
            if (!isRestore) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
              setEntityVersionForOp(op, order);
            }
            break;
          }
          case "setCellFormat": {
            const cellIndex = args.ensureCellTracked(op.sheetName, op.address);
            args.state.workbook.setCellFormat(cellIndex, op.format);
            if (!isRestore) {
              pruneCellIfOrphaned(cellIndex);
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
              setEntityVersionForOp(op, order);
            }
            break;
          }
          case "upsertCellStyle":
            args.state.workbook.upsertCellStyle(op.style);
            setEntityVersionForOp(op, order);
            break;
          case "upsertCellNumberFormat":
            args.state.workbook.upsertCellNumberFormat(op.format);
            setEntityVersionForOp(op, order);
            break;
          case "setStyleRange":
            args.state.workbook.setStyleRange(op.range, op.styleId);
            invalidatedRanges.push(op.range);
            setEntityVersionForOp(op, order);
            break;
          case "setFormatRange":
            args.state.workbook.setFormatRange(op.range, op.formatId);
            invalidatedRanges.push(op.range);
            setEntityVersionForOp(op, order);
            break;
          case "clearCell": {
            const cellIndex = getPreparedExistingCellIndex(
              op.sheetName,
              op.address,
              preparedCellAddress,
            );
            if (cellIndex === undefined) {
              setEntityVersionForOp(op, order);
              break;
            }
            changedInputCount = args.markPivotRootsChanged(
              args.clearPivotForCell(cellIndex),
              changedInputCount,
            );
            changedInputCount = args.markSpillRootsChanged(
              args.clearOwnedSpill(cellIndex),
              changedInputCount,
            );
            topologyChanged = args.removeFormula(cellIndex) || topologyChanged;
            args.state.workbook.cellStore.setValue(cellIndex, emptyValue());
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(
                CellFlags.HasFormula |
                CellFlags.JsOnly |
                CellFlags.InCycle |
                CellFlags.SpillChild |
                CellFlags.PivotOutput
              );
            pruneCellIfOrphaned(cellIndex);
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount);
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
            setEntityVersionForOp(op, order);
            break;
          }
          case "upsertDefinedName": {
            const normalizedName = normalizeDefinedName(op.name);
            args.state.workbook.setDefinedName(op.name, op.value);
            const reboundCount = formulaChangedCount;
            formulaChangedCount = args.rebindDefinedNameDependents(
              [normalizedName],
              formulaChangedCount,
            );
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
            setEntityVersionForOp(op, order);
            break;
          }
          case "deleteDefinedName": {
            const normalizedName = normalizeDefinedName(op.name);
            args.state.workbook.deleteDefinedName(op.name);
            const reboundCount = formulaChangedCount;
            formulaChangedCount = args.rebindDefinedNameDependents(
              [normalizedName],
              formulaChangedCount,
            );
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
            setEntityVersionForOp(op, order);
            break;
          }
          case "upsertPivotTable":
            applyPivotUpsertOp(op, order);
            refreshAllPivots = true;
            break;
          case "deletePivotTable": {
            const changedPivotOutputs = applyPivotDeleteOp(op, order);
            changedInputCount = args.markPivotRootsChanged(changedPivotOutputs, changedInputCount);
            changedPivotOutputs.forEach((cellIndex) => {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
            });
            refreshAllPivots = true;
            break;
          }
          default:
            assertNever(op);
        }
        appliedOps += 1;
      });

      const reboundCount = formulaChangedCount;
      formulaChangedCount = args.syncDynamicRanges(formulaChangedCount);
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1);
      args.flushWasmProgramSync();
    }

    markBatchApplied(args.state.replicaState, batch);
    if (appliedOps === 0) {
      if (source === "local") {
        emitBatch(batch);
      }
      return;
    }

    if (topologyChanged) {
      args.rebuildTopoRanks();
      args.detectCycles();
    }
    formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount);
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount);
    let recalculated = args.recalculate(
      args.composeMutationRoots(changedInputCount, formulaChangedCount),
      changedInputArray,
    );
    recalculated = args.reconcilePivotOutputs(recalculated, refreshAllPivots);
    const hasEventListeners =
      args.state.events.hasListeners() || args.state.events.hasCellListeners();
    const changed =
      isRestore || !hasEventListeners
        ? new Uint32Array()
        : args.composeEventChanges(recalculated, explicitChangedCount);
    const lastMetrics = {
      ...args.state.getLastMetrics(),
      batchId: args.state.getLastMetrics().batchId + 1,
      changedInputCount: changedInputCount + formulaChangedCount,
      compileMs,
    };
    args.state.setLastMetrics(lastMetrics);
    if (hasEventListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: "batch",
        invalidation: isRestore || sheetDeleted || structuralInvalidation ? "full" : "cells",
        changedCellIndices: changed,
        invalidatedRanges,
        invalidatedRows,
        invalidatedColumns,
        metrics: lastMetrics,
        explicitChangedCount,
      };
      if (event.invalidation === "full") {
        args.state.events.emitAllWatched(event);
      } else {
        args.state.events.emit(event, changed, (cellIndex) =>
          args.state.workbook.getQualifiedAddress(cellIndex),
        );
      }
    }
    if (source === "local") {
      void args.state.getSyncClientConnection()?.send(batch);
      emitBatch(batch);
    } else if (source === "remote" && args.state.redoStack.length > 0) {
      args.state.redoStack.length = 0;
    }
  };

  const applyCellMutationsAtNow = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch,
    source: "local" | "restore",
    potentialNewCells?: number,
  ): void => {
    const isRestore = source === "restore";
    args.beginMutationCollection();
    let changedInputCount = 0;
    let formulaChangedCount = 0;
    let explicitChangedCount = 0;
    let topologyChanged = false;
    let compileMs = 0;
    const reservedNewCells = potentialNewCells ?? refs.length;
    args.state.workbook.cellStore.ensureCapacity(
      args.state.workbook.cellStore.size + reservedNewCells,
    );
    args.resetMaterializedCellScratch(reservedNewCells);

    const sheetNameById = new Map<number, string>();
    const resolveSheetName = (sheetId: number): string => {
      const cached = sheetNameById.get(sheetId);
      if (cached !== undefined) {
        return cached;
      }
      const sheet = args.state.workbook.getSheetById(sheetId);
      if (!sheet) {
        throw new Error(`Unknown sheet id: ${sheetId}`);
      }
      sheetNameById.set(sheetId, sheet.name);
      return sheet.name;
    };

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1);
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        refs.forEach((ref, refIndex) => {
          const { sheetId, mutation } = ref;
          const order = args.state.trackReplicaVersions ? batchOpOrder(batch, refIndex) : undefined;
          const existingIndex = args.state.workbook.cellKeyToIndex.get(
            makeCellKey(sheetId, mutation.row, mutation.col),
          );

          switch (mutation.kind) {
            case "setCellValue": {
              if (existingIndex !== undefined && canFastPathLiteralOverwrite(existingIndex)) {
                writeLiteralToCellStore(
                  args.state.workbook.cellStore,
                  existingIndex,
                  mutation.value,
                  args.state.strings,
                );
                changedInputCount = args.markInputChanged(existingIndex, changedInputCount);
                if (!isRestore) {
                  explicitChangedCount = args.markExplicitChanged(
                    existingIndex,
                    explicitChangedCount,
                  );
                }
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(
                    resolveSheetName(sheetId),
                    formatAddress(mutation.row, mutation.col),
                    order!,
                  );
                }
                break;
              }
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(
                  args.clearPivotForCell(existingIndex),
                  changedInputCount,
                );
              }
              const cellIndex = args.state.workbook.ensureCellAt(
                sheetId,
                mutation.row,
                mutation.col,
              ).cellIndex;
              if (!isRestore) {
                changedInputCount = args.markSpillRootsChanged(
                  args.clearOwnedSpill(cellIndex),
                  changedInputCount,
                );
                topologyChanged = args.removeFormula(cellIndex) || topologyChanged;
              }
              writeLiteralToCellStore(
                args.state.workbook.cellStore,
                cellIndex,
                mutation.value,
                args.state.strings,
              );
              args.state.workbook.cellStore.flags[cellIndex] =
                (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
                ~(
                  CellFlags.HasFormula |
                  CellFlags.JsOnly |
                  CellFlags.InCycle |
                  CellFlags.SpillChild |
                  CellFlags.PivotOutput
                );
              if (!isRestore && mutation.value === null) {
                pruneCellIfOrphaned(cellIndex);
              }
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount);
              if (!isRestore) {
                explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(
                  resolveSheetName(sheetId),
                  formatAddress(mutation.row, mutation.col),
                  order!,
                );
              }
              break;
            }
            case "setCellFormula": {
              const sheetName = resolveSheetName(sheetId);
              if (!isRestore && existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(
                  args.clearPivotForCell(existingIndex),
                  changedInputCount,
                );
              }
              const cellIndex = args.state.workbook.ensureCellAt(
                sheetId,
                mutation.row,
                mutation.col,
              ).cellIndex;
              if (!isRestore) {
                changedInputCount = args.markSpillRootsChanged(
                  args.clearOwnedSpill(cellIndex),
                  changedInputCount,
                );
              }
              const compileStarted = isRestore ? 0 : performance.now();
              try {
                const changedTopology = args.bindFormula(cellIndex, sheetName, mutation.formula);
                if (!isRestore) {
                  compileMs += performance.now() - compileStarted;
                }
                formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
                topologyChanged = topologyChanged || changedTopology;
              } catch {
                if (!isRestore) {
                  compileMs += performance.now() - compileStarted;
                }
                topologyChanged = args.removeFormula(cellIndex) || topologyChanged;
                args.setInvalidFormulaValue(cellIndex);
                changedInputCount = args.markInputChanged(cellIndex, changedInputCount);
              }
              if (!isRestore) {
                explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), order!);
              }
              break;
            }
            case "clearCell": {
              if (existingIndex !== undefined && canFastPathLiteralOverwrite(existingIndex)) {
                args.state.workbook.cellStore.setValue(existingIndex, emptyValue());
                changedInputCount = args.markInputChanged(existingIndex, changedInputCount);
                if (!isRestore) {
                  explicitChangedCount = args.markExplicitChanged(
                    existingIndex,
                    explicitChangedCount,
                  );
                }
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(
                    resolveSheetName(sheetId),
                    formatAddress(mutation.row, mutation.col),
                    order!,
                  );
                }
                break;
              }
              if (existingIndex === undefined) {
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(
                    resolveSheetName(sheetId),
                    formatAddress(mutation.row, mutation.col),
                    order!,
                  );
                }
                break;
              }
              changedInputCount = args.markPivotRootsChanged(
                args.clearPivotForCell(existingIndex),
                changedInputCount,
              );
              changedInputCount = args.markSpillRootsChanged(
                args.clearOwnedSpill(existingIndex),
                changedInputCount,
              );
              topologyChanged = args.removeFormula(existingIndex) || topologyChanged;
              args.state.workbook.cellStore.setValue(existingIndex, emptyValue());
              args.state.workbook.cellStore.flags[existingIndex] =
                (args.state.workbook.cellStore.flags[existingIndex] ?? 0) &
                ~(
                  CellFlags.HasFormula |
                  CellFlags.JsOnly |
                  CellFlags.InCycle |
                  CellFlags.SpillChild |
                  CellFlags.PivotOutput
                );
              if (!isRestore) {
                pruneCellIfOrphaned(existingIndex);
              }
              changedInputCount = args.markInputChanged(existingIndex, changedInputCount);
              if (!isRestore) {
                explicitChangedCount = args.markExplicitChanged(
                  existingIndex,
                  explicitChangedCount,
                );
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(
                  resolveSheetName(sheetId),
                  formatAddress(mutation.row, mutation.col),
                  order!,
                );
              }
              break;
            }
            default:
              assertNever(mutation);
          }
        });
      });

      const reboundCount = formulaChangedCount;
      formulaChangedCount = args.syncDynamicRanges(formulaChangedCount);
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1);
      args.flushWasmProgramSync();
    }

    markBatchApplied(args.state.replicaState, batch);
    if (refs.length === 0) {
      if (!isRestore) {
        emitBatch(batch);
      }
      return;
    }

    if (topologyChanged) {
      args.rebuildTopoRanks();
      args.detectCycles();
    }
    formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount);
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount);
    let recalculated = args.recalculate(
      args.composeMutationRoots(changedInputCount, formulaChangedCount),
      changedInputArray,
    );
    recalculated = args.reconcilePivotOutputs(recalculated, false);
    const hasEventListeners =
      args.state.events.hasListeners() || args.state.events.hasCellListeners();
    const changed =
      isRestore || !hasEventListeners
        ? new Uint32Array()
        : args.composeEventChanges(recalculated, explicitChangedCount);
    const lastMetrics = {
      ...args.state.getLastMetrics(),
      batchId: args.state.getLastMetrics().batchId + 1,
      changedInputCount: changedInputCount + formulaChangedCount,
      compileMs,
    };
    args.state.setLastMetrics(lastMetrics);
    if (hasEventListeners) {
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: "batch",
        invalidation: isRestore ? "full" : "cells",
        changedCellIndices: changed,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      };
      if (isRestore) {
        args.state.events.emitAllWatched(event);
        return;
      }
      args.state.events.emit(event, changed, (cellIndex) =>
        args.state.workbook.getQualifiedAddress(cellIndex),
      );
    } else if (isRestore) {
      return;
    }
    void args.state.getSyncClientConnection()?.send(batch);
    emitBatch(batch);
  };

  return {
    applyBatch(batch, source, potentialNewCells, preparedCellAddressesByOpIndex) {
      return Effect.try({
        try: () => {
          applyBatchNow(batch, source, potentialNewCells, preparedCellAddressesByOpIndex);
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply ${source} batch`, cause),
            cause,
          }),
      });
    },
    applyCellMutationsAt(refs, batch, source, potentialNewCells) {
      return Effect.try({
        try: () => {
          applyCellMutationsAtNow(refs, batch, source, potentialNewCells);
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply ${source} cell mutations`, cause),
            cause,
          }),
      });
    },
    applyDerivedOp(op) {
      return Effect.try({
        try: () => applyDerivedOpNow(op),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply derived operation ${op.kind}`, cause),
            cause,
          }),
      });
    },
  };
}
