import { Effect } from "effect";
import { formatAddress } from "@bilig/formula";
import type { EngineOp, EngineOpBatch } from "@bilig/workbook-domain";
import { ValueTag, type CellRangeRef, type CellSnapshot, type LiteralInput } from "@bilig/protocol";
import { normalizeRange } from "../../engine-range-utils.js";
import { cloneCellStyleRecord } from "../../engine-style-utils.js";
import { restoreFormatRangeOps, restoreStyleRangeOps } from "../../engine-range-format-ops.js";
import { sheetMetadataToOps } from "../../engine-snapshot-utils.js";
import { parseCsv, parseCsvCellInput } from "../../csv.js";
import { createBatch } from "../../replica-state.js";
import { makeCellKey, type WorkbookStore } from "../../workbook-store.js";
import {
  cellMutationRefToEngineOp,
  cloneCellMutationRef,
  countPotentialNewCellsForMutationRefs,
  type EngineCellMutationRef,
} from "../../cell-mutations-at.js";
import type { CommitOp, EngineRuntimeState, TransactionRecord } from "../runtime-state.js";
import { EngineMutationError } from "../errors.js";
import {
  tryBuildFastMutationHistory,
  type FastMutationHistoryResult,
} from "./mutation-history-fast-path.js";

function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

function getMatrixCell(
  matrix: readonly (readonly CellSnapshot[])[],
  rowIndex: number,
  colIndex: number,
): CellSnapshot {
  const row = matrix[rowIndex];
  if (row === undefined) {
    throw new RangeError(`Missing source row at index ${rowIndex}`);
  }
  const cell = row[colIndex];
  if (cell === undefined) {
    throw new RangeError(`Missing source cell at row ${rowIndex}, column ${colIndex}`);
  }
  return cell;
}

export interface EngineMutationService {
  readonly executeTransactionNow: (
    record: TransactionRecord,
    source: "local" | "restore" | "history",
  ) => void;
  readonly executeTransaction: (
    record: TransactionRecord,
    source: "local" | "restore" | "history",
  ) => Effect.Effect<void, EngineMutationError>;
  readonly executeLocalNow: (
    ops: EngineOp[],
    potentialNewCells?: number,
  ) => readonly EngineOp[] | null;
  readonly executeLocalCellMutationsAtNow: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ) => readonly EngineOp[] | null;
  readonly applyCellMutationsAtNow: (
    refs: readonly EngineCellMutationRef[],
    options?: {
      captureUndo?: boolean;
      potentialNewCells?: number;
      source?: "local" | "restore";
    },
  ) => readonly EngineOp[] | null;
  readonly applyCellMutationsAt: (
    refs: readonly EngineCellMutationRef[],
    options?: {
      captureUndo?: boolean;
      potentialNewCells?: number;
      source?: "local" | "restore";
    },
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>;
  readonly executeLocal: (
    ops: EngineOp[],
    potentialNewCells?: number,
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>;
  readonly applyOpsNow: (
    ops: readonly EngineOp[],
    options?: {
      captureUndo?: boolean;
      potentialNewCells?: number;
      source?: "local" | "restore";
      trusted?: boolean;
    },
  ) => readonly EngineOp[] | null;
  readonly applyOps: (
    ops: readonly EngineOp[],
    options?: {
      captureUndo?: boolean;
      potentialNewCells?: number;
      source?: "local" | "restore";
      trusted?: boolean;
    },
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>;
  readonly captureUndoOps: <Result>(mutate: () => Result) => Effect.Effect<
    {
      result: Result;
      undoOps: readonly EngineOp[] | null;
    },
    EngineMutationError
  >;
  readonly setRangeValues: (
    range: CellRangeRef,
    values: readonly (readonly LiteralInput[])[],
  ) => Effect.Effect<void, EngineMutationError>;
  readonly setRangeFormulas: (
    range: CellRangeRef,
    formulas: readonly (readonly string[])[],
  ) => Effect.Effect<void, EngineMutationError>;
  readonly clearRange: (range: CellRangeRef) => Effect.Effect<void, EngineMutationError>;
  readonly fillRange: (
    source: CellRangeRef,
    target: CellRangeRef,
  ) => Effect.Effect<void, EngineMutationError>;
  readonly copyRange: (
    source: CellRangeRef,
    target: CellRangeRef,
  ) => Effect.Effect<void, EngineMutationError>;
  readonly moveRange: (
    source: CellRangeRef,
    target: CellRangeRef,
  ) => Effect.Effect<void, EngineMutationError>;
  readonly importSheetCsv: (
    sheetName: string,
    csv: string,
  ) => Effect.Effect<void, EngineMutationError>;
  readonly renderCommit: (ops: CommitOp[]) => Effect.Effect<void, EngineMutationError>;
}

export function createEngineMutationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | "replicaState"
    | "undoStack"
    | "redoStack"
    | "getTransactionReplayDepth"
    | "setTransactionReplayDepth"
  > & {
    readonly workbook: WorkbookStore;
  };
  readonly captureSheetCellState: (sheetName: string) => EngineOp[];
  readonly captureRowRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => EngineOp[];
  readonly captureColumnRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => EngineOp[];
  readonly restoreCellOps: (sheetName: string, address: string) => EngineOp[];
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
  readonly readRangeCells: (range: CellRangeRef) => CellSnapshot[][];
  readonly toCellStateOps: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[];
  readonly applyBatchNow: (
    batch: EngineOpBatch,
    source: "local" | "restore" | "history",
    potentialNewCells?: number,
  ) => void;
  readonly applyCellMutationsAtBatchNow: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch,
    source: "local" | "restore",
    potentialNewCells?: number,
  ) => void;
}): EngineMutationService {
  const restoreCellOpFromRef = (ref: EngineCellMutationRef): EngineOp => {
    const sheet = args.state.workbook.getSheetById(ref.sheetId);
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${ref.sheetId}`);
    }
    const address = formatAddress(ref.mutation.row, ref.mutation.col);
    const cellIndex = args.state.workbook.cellKeyToIndex.get(
      makeCellKey(ref.sheetId, ref.mutation.row, ref.mutation.col),
    );
    if (cellIndex === undefined) {
      return { kind: "clearCell", sheetName: sheet.name, address };
    }
    const snapshot = args.getCellByIndex(cellIndex);
    if (snapshot.formula !== undefined) {
      return {
        kind: "setCellFormula",
        sheetName: sheet.name,
        address,
        formula: snapshot.formula,
      };
    }
    switch (snapshot.value.tag) {
      case ValueTag.Empty:
      case ValueTag.Error:
        return { kind: "clearCell", sheetName: sheet.name, address };
      case ValueTag.Number:
      case ValueTag.Boolean:
      case ValueTag.String:
        return {
          kind: "setCellValue",
          sheetName: sheet.name,
          address,
          value: snapshot.value.value,
        };
    }
  };

  const buildFastMutationHistoryFromRefs = (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells: number,
  ): FastMutationHistoryResult => {
    const forwardOps: EngineOp[] = Array.from({ length: refs.length });
    for (let index = 0; index < refs.length; index += 1) {
      const ref = refs[index]!;
      forwardOps[index] = cellMutationRefToEngineOp(args.state.workbook, ref);
    }

    const inverseOps: EngineOp[] = [];
    for (let index = refs.length - 1; index >= 0; index -= 1) {
      inverseOps.push(restoreCellOpFromRef(refs[index]!));
    }

    return {
      forward: { ops: forwardOps, potentialNewCells },
      inverse: { ops: inverseOps, potentialNewCells: refs.length },
      undoOps: structuredClone(inverseOps),
    };
  };

  const inverseOpsFor = (op: EngineOp): EngineOp[] => {
    switch (op.kind) {
      case "upsertWorkbook":
        return [{ kind: "upsertWorkbook", name: args.state.workbook.workbookName }];
      case "setWorkbookMetadata": {
        const existing = args.state.workbook.getWorkbookProperty(op.key);
        return [{ kind: "setWorkbookMetadata", key: op.key, value: existing?.value ?? null }];
      }
      case "setCalculationSettings":
        return [
          {
            kind: "setCalculationSettings",
            settings: args.state.workbook.getCalculationSettings(),
          },
        ];
      case "setVolatileContext":
        return [{ kind: "setVolatileContext", context: args.state.workbook.getVolatileContext() }];
      case "upsertSheet": {
        const existing = args.state.workbook.getSheet(op.name);
        if (!existing) {
          return [{ kind: "deleteSheet", name: op.name }];
        }
        return [{ kind: "upsertSheet", name: existing.name, order: existing.order }];
      }
      case "renameSheet": {
        const existing = args.state.workbook.getSheet(op.newName);
        if (!existing) {
          return [];
        }
        return [{ kind: "renameSheet", oldName: op.newName, newName: op.oldName }];
      }
      case "deleteSheet": {
        const sheet = args.state.workbook.getSheet(op.name);
        if (!sheet) {
          return [];
        }
        const restoredOps: EngineOp[] = [
          { kind: "upsertSheet", name: sheet.name, order: sheet.order },
        ];
        restoredOps.push(...sheetMetadataToOps(args.state.workbook, sheet.name));
        args.state.workbook
          .listTables()
          .filter((table) => table.sheetName === sheet.name)
          .forEach((table) => {
            restoredOps.push({
              kind: "upsertTable",
              table: {
                name: table.name,
                sheetName: table.sheetName,
                startAddress: table.startAddress,
                endAddress: table.endAddress,
                columnNames: [...table.columnNames],
                headerRow: table.headerRow,
                totalsRow: table.totalsRow,
              },
            });
          });
        args.state.workbook
          .listSpills()
          .filter((spill) => spill.sheetName === sheet.name)
          .forEach((spill) => {
            restoredOps.push({
              kind: "upsertSpillRange",
              sheetName: spill.sheetName,
              address: spill.address,
              rows: spill.rows,
              cols: spill.cols,
            });
          });
        args.state.workbook
          .listPivots()
          .filter((pivot) => pivot.sheetName === sheet.name)
          .forEach((pivot) => {
            restoredOps.push({
              kind: "upsertPivotTable",
              name: pivot.name,
              sheetName: pivot.sheetName,
              address: pivot.address,
              source: { ...pivot.source },
              groupBy: [...pivot.groupBy],
              values: pivot.values.map((value) => Object.assign({}, value)),
              rows: pivot.rows,
              cols: pivot.cols,
            });
          });
        restoredOps.push(...args.captureSheetCellState(sheet.name));
        return restoredOps;
      }
      case "insertRows":
        return [{ kind: "deleteRows", sheetName: op.sheetName, start: op.start, count: op.count }];
      case "deleteRows": {
        const entries = args.state.workbook.snapshotRowAxisEntries(
          op.sheetName,
          op.start,
          op.count,
        );
        return [
          {
            kind: "insertRows",
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
          ...args.captureRowRangeCellState(op.sheetName, op.start, op.count),
        ];
      }
      case "moveRows":
        return [
          {
            kind: "moveRows",
            sheetName: op.sheetName,
            start: op.target,
            count: op.count,
            target: op.start,
          },
        ];
      case "insertColumns":
        return [
          { kind: "deleteColumns", sheetName: op.sheetName, start: op.start, count: op.count },
        ];
      case "deleteColumns": {
        const entries = args.state.workbook.snapshotColumnAxisEntries(
          op.sheetName,
          op.start,
          op.count,
        );
        return [
          {
            kind: "insertColumns",
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
          ...args.captureColumnRangeCellState(op.sheetName, op.start, op.count),
        ];
      }
      case "moveColumns":
        return [
          {
            kind: "moveColumns",
            sheetName: op.sheetName,
            start: op.target,
            count: op.count,
            target: op.start,
          },
        ];
      case "updateRowMetadata": {
        const existing = args.state.workbook.getRowMetadata(op.sheetName, op.start, op.count);
        return [
          {
            kind: "updateRowMetadata",
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            size: existing?.size ?? null,
            hidden: existing?.hidden ?? null,
          },
        ];
      }
      case "updateColumnMetadata": {
        const existing = args.state.workbook.getColumnMetadata(op.sheetName, op.start, op.count);
        return [
          {
            kind: "updateColumnMetadata",
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            size: existing?.size ?? null,
            hidden: existing?.hidden ?? null,
          },
        ];
      }
      case "setFreezePane": {
        const existing = args.state.workbook.getFreezePane(op.sheetName);
        if (!existing) {
          return [{ kind: "clearFreezePane", sheetName: op.sheetName }];
        }
        return [
          {
            kind: "setFreezePane",
            sheetName: op.sheetName,
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "clearFreezePane": {
        const existing = args.state.workbook.getFreezePane(op.sheetName);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "setFreezePane",
            sheetName: op.sheetName,
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "setFilter": {
        const existing = args.state.workbook.getFilter(op.sheetName, op.range);
        if (!existing) {
          return [{ kind: "clearFilter", sheetName: op.sheetName, range: { ...op.range } }];
        }
        return [{ kind: "setFilter", sheetName: op.sheetName, range: { ...existing.range } }];
      }
      case "clearFilter": {
        const existing = args.state.workbook.getFilter(op.sheetName, op.range);
        if (!existing) {
          return [];
        }
        return [{ kind: "setFilter", sheetName: op.sheetName, range: { ...existing.range } }];
      }
      case "setSort": {
        const existing = args.state.workbook.getSort(op.sheetName, op.range);
        if (!existing) {
          return [{ kind: "clearSort", sheetName: op.sheetName, range: { ...op.range } }];
        }
        return [
          {
            kind: "setSort",
            sheetName: op.sheetName,
            range: { ...existing.range },
            keys: existing.keys.map((key) => Object.assign({}, key)),
          },
        ];
      }
      case "clearSort": {
        const existing = args.state.workbook.getSort(op.sheetName, op.range);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "setSort",
            sheetName: op.sheetName,
            range: { ...existing.range },
            keys: existing.keys.map((key) => Object.assign({}, key)),
          },
        ];
      }
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
        return args.restoreCellOps(op.sheetName, op.address);
      case "setCellFormat": {
        const cellIndex = args.state.workbook.getCellIndex(op.sheetName, op.address);
        return [
          {
            kind: "setCellFormat",
            sheetName: op.sheetName,
            address: op.address,
            format:
              cellIndex === undefined
                ? null
                : (args.state.workbook.getCellFormat(cellIndex) ?? null),
          },
        ];
      }
      case "upsertCellStyle": {
        const existing = args.state.workbook.getCellStyle(op.style.id);
        if (!existing || existing.id !== op.style.id) {
          return [];
        }
        return [{ kind: "upsertCellStyle", style: cloneCellStyleRecord(existing) }];
      }
      case "upsertCellNumberFormat": {
        const existing = args.state.workbook.getCellNumberFormat(op.format.id);
        if (!existing || existing.id !== op.format.id) {
          return [];
        }
        return [{ kind: "upsertCellNumberFormat", format: { ...existing } }];
      }
      case "setStyleRange":
        return restoreStyleRangeOps(args.state.workbook, op.range);
      case "setFormatRange":
        return restoreFormatRangeOps(args.state.workbook, op.range);
      case "upsertDefinedName": {
        const existing = args.state.workbook.getDefinedName(op.name);
        if (!existing) {
          return [{ kind: "deleteDefinedName", name: op.name }];
        }
        return [{ kind: "upsertDefinedName", name: existing.name, value: existing.value }];
      }
      case "deleteDefinedName": {
        const existing = args.state.workbook.getDefinedName(op.name);
        if (!existing) {
          return [];
        }
        return [{ kind: "upsertDefinedName", name: existing.name, value: existing.value }];
      }
      case "upsertTable": {
        const existing = args.state.workbook.getTable(op.table.name);
        if (!existing) {
          return [{ kind: "deleteTable", name: op.table.name }];
        }
        return [
          {
            kind: "upsertTable",
            table: {
              name: existing.name,
              sheetName: existing.sheetName,
              startAddress: existing.startAddress,
              endAddress: existing.endAddress,
              columnNames: [...existing.columnNames],
              headerRow: existing.headerRow,
              totalsRow: existing.totalsRow,
            },
          },
        ];
      }
      case "deleteTable": {
        const existing = args.state.workbook.getTable(op.name);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "upsertTable",
            table: {
              name: existing.name,
              sheetName: existing.sheetName,
              startAddress: existing.startAddress,
              endAddress: existing.endAddress,
              columnNames: [...existing.columnNames],
              headerRow: existing.headerRow,
              totalsRow: existing.totalsRow,
            },
          },
        ];
      }
      case "upsertSpillRange": {
        const existing = args.state.workbook.getSpill(op.sheetName, op.address);
        if (!existing) {
          return [{ kind: "deleteSpillRange", sheetName: op.sheetName, address: op.address }];
        }
        return [
          {
            kind: "upsertSpillRange",
            sheetName: existing.sheetName,
            address: existing.address,
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "deleteSpillRange": {
        const existing = args.state.workbook.getSpill(op.sheetName, op.address);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "upsertSpillRange",
            sheetName: existing.sheetName,
            address: existing.address,
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "upsertPivotTable": {
        const existing = args.state.workbook.getPivot(op.sheetName, op.address);
        if (!existing) {
          return [{ kind: "deletePivotTable", sheetName: op.sheetName, address: op.address }];
        }
        return [
          {
            kind: "upsertPivotTable",
            name: existing.name,
            sheetName: existing.sheetName,
            address: existing.address,
            source: { ...existing.source },
            groupBy: [...existing.groupBy],
            values: existing.values.map((v) => Object.assign({}, v)),
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "deletePivotTable": {
        const existing = args.state.workbook.getPivot(op.sheetName, op.address);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "upsertPivotTable",
            name: existing.name,
            sheetName: existing.sheetName,
            address: existing.address,
            source: { ...existing.source },
            groupBy: [...existing.groupBy],
            values: existing.values.map((value) => Object.assign({}, value)),
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      default: {
        const exhaustive: never = op;
        return exhaustive;
      }
    }
  };

  const buildInverseOps = (ops: readonly EngineOp[]): EngineOp[] => {
    const inverseOps: EngineOp[] = [];
    for (let index = ops.length - 1; index >= 0; index -= 1) {
      const op = ops[index];
      if (op !== undefined) {
        inverseOps.push(...inverseOpsFor(op));
      }
    }
    return inverseOps;
  };

  const canonicalizeForwardOps = (ops: readonly EngineOp[]): EngineOp[] =>
    ops.map((op) => {
      if (op.kind === "insertRows") {
        return op.entries
          ? { ...op, entries: op.entries.map((entry) => ({ ...entry })) }
          : {
              ...op,
              entries: args.state.workbook.snapshotRowAxisEntries(op.sheetName, op.start, op.count),
            };
      }

      if (op.kind === "insertColumns") {
        return op.entries
          ? { ...op, entries: op.entries.map((entry) => ({ ...entry })) }
          : {
              ...op,
              entries: args.state.workbook.snapshotColumnAxisEntries(
                op.sheetName,
                op.start,
                op.count,
              ),
            };
      }

      return structuredClone(op);
    });

  const executeTransactionNow = (
    record: TransactionRecord,
    source: "local" | "restore" | "history",
  ): void => {
    if (record.ops.length === 0) {
      return;
    }
    const batch = createBatch(args.state.replicaState, record.ops);
    args.applyBatchNow(batch, source, record.potentialNewCells);
  };

  const executeLocalCellMutationsAtNow = (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ): readonly EngineOp[] | null => {
    if (refs.length === 0) {
      return null;
    }
    const nextRefs = refs.map((ref) => cloneCellMutationRef(ref));
    const nextPotentialNewCells =
      potentialNewCells ?? countPotentialNewCellsForMutationRefs(nextRefs);
    const fastHistory = buildFastMutationHistoryFromRefs(nextRefs, nextPotentialNewCells);
    const inverse: TransactionRecord = fastHistory?.inverse ?? {
      ops: buildInverseOps(fastHistory.forward.ops),
      potentialNewCells: fastHistory.forward.ops.length,
    };
    args.applyCellMutationsAtBatchNow(
      nextRefs,
      createBatch(args.state.replicaState, fastHistory.forward.ops),
      "local",
      nextPotentialNewCells,
    );
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward: fastHistory?.forward ?? {
          ops: canonicalizeForwardOps(fastHistory.forward.ops),
          potentialNewCells: nextPotentialNewCells,
        },
        inverse,
      });
      args.state.redoStack.length = 0;
    }
    return fastHistory?.undoOps ?? structuredClone(inverse.ops);
  };

  const applyCellMutationsAtNow = (
    refs: readonly EngineCellMutationRef[],
    options: {
      captureUndo?: boolean;
      potentialNewCells?: number;
      source?: "local" | "restore";
    } = {},
  ): readonly EngineOp[] | null => {
    const source = options.source ?? "restore";
    const captureUndo = options.captureUndo ?? source === "local";
    if (captureUndo) {
      return executeLocalCellMutationsAtNow(refs, options.potentialNewCells);
    }
    if (refs.length === 0) {
      return null;
    }
    const nextRefs = refs.map((ref) => cloneCellMutationRef(ref));
    const nextPotentialNewCells =
      options.potentialNewCells ?? countPotentialNewCellsForMutationRefs(nextRefs);
    const forwardOps = nextRefs.map((ref) => cellMutationRefToEngineOp(args.state.workbook, ref));
    args.applyCellMutationsAtBatchNow(
      nextRefs,
      createBatch(args.state.replicaState, forwardOps),
      source,
      nextPotentialNewCells,
    );
    return null;
  };

  return {
    executeTransactionNow: executeTransactionNow,
    executeTransaction(record, source) {
      return Effect.try({
        try: () => {
          executeTransactionNow(record, source);
        },
        catch: (cause) =>
          new EngineMutationError({
            message: `Failed to execute ${source} transaction`,
            cause,
          }),
      });
    },
    executeLocalNow(ops, potentialNewCells) {
      if (ops.length === 0) {
        return null;
      }
      const forward: TransactionRecord =
        potentialNewCells === undefined ? { ops } : { ops, potentialNewCells };
      const fastHistory = tryBuildFastMutationHistory(
        potentialNewCells === undefined
          ? {
              workbook: args.state.workbook,
              getCellByIndex: args.getCellByIndex,
              ops,
            }
          : {
              workbook: args.state.workbook,
              getCellByIndex: args.getCellByIndex,
              ops,
              potentialNewCells,
            },
      );
      const inverse: TransactionRecord = fastHistory?.inverse ?? {
        ops: buildInverseOps(ops),
        potentialNewCells: ops.length,
      };
      executeTransactionNow(forward, "local");
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward:
            fastHistory?.forward ??
            (potentialNewCells === undefined
              ? { ops: canonicalizeForwardOps(ops) }
              : { ops: canonicalizeForwardOps(ops), potentialNewCells }),
          inverse,
        });
        args.state.redoStack.length = 0;
      }
      return fastHistory?.undoOps ?? structuredClone(inverse.ops);
    },
    executeLocalCellMutationsAtNow(refs, potentialNewCells) {
      return executeLocalCellMutationsAtNow(refs, potentialNewCells);
    },
    applyCellMutationsAtNow(refs, options = {}) {
      return applyCellMutationsAtNow(refs, options);
    },
    applyCellMutationsAt(refs, options = {}) {
      return Effect.try({
        try: () => applyCellMutationsAtNow(refs, options),
        catch: (cause) =>
          new EngineMutationError({
            message: "Failed to apply cell mutations",
            cause,
          }),
      });
    },
    executeLocal(ops, potentialNewCells) {
      return Effect.try({
        try: () => this.executeLocalNow(ops, potentialNewCells),
        catch: (cause) =>
          new EngineMutationError({
            message: "Failed to execute local transaction",
            cause,
          }),
      });
    },
    applyOpsNow(ops, options = {}) {
      const nextOps = options.trusted ? Array.from(ops) : structuredClone([...ops]);
      if (nextOps.length === 0) {
        return null;
      }
      if (options.captureUndo) {
        return this.executeLocalNow(nextOps, options.potentialNewCells);
      }
      executeTransactionNow(
        options.potentialNewCells === undefined
          ? { ops: nextOps }
          : { ops: nextOps, potentialNewCells: options.potentialNewCells },
        options.source ?? "restore",
      );
      return null;
    },
    applyOps(ops, options = {}) {
      return Effect.try({
        try: () => this.applyOpsNow(ops, options),
        catch: (cause) =>
          new EngineMutationError({
            message: "Failed to apply engine operations",
            cause,
          }),
      });
    },
    captureUndoOps(mutate) {
      return Effect.try({
        try: () => {
          const previousUndoDepth = args.state.undoStack.length;
          const result = mutate();
          if (args.state.undoStack.length === previousUndoDepth) {
            return {
              result,
              undoOps: null,
            };
          }
          if (args.state.undoStack.length === previousUndoDepth + 1) {
            return {
              result,
              undoOps: structuredClone(args.state.undoStack.at(-1)?.inverse.ops ?? null),
            };
          }
          throw new Error("Expected a single local transaction while capturing undo ops");
        },
        catch: (cause) =>
          new EngineMutationError({
            message: "Failed to capture undo ops",
            cause,
          }),
      });
    },
    setRangeValues(range, values) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range);
          const expectedHeight = bounds.endRow - bounds.startRow + 1;
          const expectedWidth = bounds.endCol - bounds.startCol + 1;
          if (
            values.length !== expectedHeight ||
            values.some((row) => row.length !== expectedWidth)
          ) {
            throw new Error(
              "setRangeValues requires a value matrix that exactly matches the target range",
            );
          }

          const opCount = expectedHeight * expectedWidth;
          const ops = Array.from<EngineOp>({ length: opCount });
          let opIndex = 0;
          for (let rowOffset = 0; rowOffset < expectedHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < expectedWidth; colOffset += 1) {
              ops[opIndex] = {
                kind: "setCellValue",
                sheetName: range.sheetName,
                address: formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset),
                value: values[rowOffset]![colOffset] ?? null,
              };
              opIndex += 1;
            }
          }
          Effect.runSync(this.executeLocal(ops, opCount));
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage("Failed to set range values", cause),
            cause,
          }),
      });
    },
    setRangeFormulas(range, formulas) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range);
          const expectedHeight = bounds.endRow - bounds.startRow + 1;
          const expectedWidth = bounds.endCol - bounds.startCol + 1;
          if (
            formulas.length !== expectedHeight ||
            formulas.some((row) => row.length !== expectedWidth)
          ) {
            throw new Error(
              "setRangeFormulas requires a formula matrix that exactly matches the target range",
            );
          }

          const opCount = expectedHeight * expectedWidth;
          const ops = Array.from<EngineOp>({ length: opCount });
          let opIndex = 0;
          for (let rowOffset = 0; rowOffset < expectedHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < expectedWidth; colOffset += 1) {
              ops[opIndex] = {
                kind: "setCellFormula",
                sheetName: range.sheetName,
                address: formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset),
                formula: formulas[rowOffset]![colOffset] ?? "",
              };
              opIndex += 1;
            }
          }
          Effect.runSync(this.executeLocal(ops, opCount));
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage("Failed to set range formulas", cause),
            cause,
          }),
      });
    },
    clearRange(range) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range);
          const opCount =
            (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1);
          const ops = Array.from<EngineOp>({ length: opCount });
          let opIndex = 0;
          for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
            for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
              ops[opIndex] = {
                kind: "clearCell",
                sheetName: range.sheetName,
                address: formatAddress(row, col),
              };
              opIndex += 1;
            }
          }
          Effect.runSync(this.executeLocal(ops, opCount));
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage("Failed to clear range", cause),
            cause,
          }),
      });
    },
    fillRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source);
          const targetBounds = normalizeRange(target);
          const sourceBounds = normalizeRange(source);
          const sourceHeight = sourceMatrix.length;
          const sourceWidth = sourceMatrix[0]?.length ?? 0;
          if (sourceHeight === 0 || sourceWidth === 0) {
            return;
          }

          const ops: EngineOp[] = [];
          for (let row = targetBounds.startRow; row <= targetBounds.endRow; row += 1) {
            for (let col = targetBounds.startCol; col <= targetBounds.endCol; col += 1) {
              const sourceRowOffset = (row - targetBounds.startRow) % sourceHeight;
              const sourceColOffset = (col - targetBounds.startCol) % sourceWidth;
              const sourceCell = getMatrixCell(sourceMatrix, sourceRowOffset, sourceColOffset);
              const sourceAddress = formatAddress(
                sourceBounds.startRow + sourceRowOffset,
                sourceBounds.startCol + sourceColOffset,
              );
              ops.push(
                ...args.toCellStateOps(
                  target.sheetName,
                  formatAddress(row, col),
                  sourceCell,
                  source.sheetName,
                  sourceAddress,
                ),
              );
            }
          }
          Effect.runSync(this.executeLocal(ops, ops.length));
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage("Failed to fill range", cause),
            cause,
          }),
      });
    },
    copyRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source);
          const targetBounds = normalizeRange(target);
          const sourceBounds = normalizeRange(source);
          const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1;
          const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1;
          const targetHeight = targetBounds.endRow - targetBounds.startRow + 1;
          const targetWidth = targetBounds.endCol - targetBounds.startCol + 1;
          if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
            throw new Error("copyRange requires source and target dimensions to match exactly");
          }

          const ops: EngineOp[] = [];
          for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
              const nextAddress = formatAddress(
                targetBounds.startRow + rowOffset,
                targetBounds.startCol + colOffset,
              );
              const sourceAddress = formatAddress(
                sourceBounds.startRow + rowOffset,
                sourceBounds.startCol + colOffset,
              );
              ops.push(
                ...args.toCellStateOps(
                  target.sheetName,
                  nextAddress,
                  getMatrixCell(sourceMatrix, rowOffset, colOffset),
                  source.sheetName,
                  sourceAddress,
                ),
              );
            }
          }
          Effect.runSync(this.executeLocal(ops, ops.length));
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage("Failed to copy range", cause),
            cause,
          }),
      });
    },
    moveRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source);
          const targetBounds = normalizeRange(target);
          const sourceBounds = normalizeRange(source);
          const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1;
          const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1;
          const targetHeight = targetBounds.endRow - targetBounds.startRow + 1;
          const targetWidth = targetBounds.endCol - targetBounds.startCol + 1;
          if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
            throw new Error("moveRange requires source and target dimensions to match exactly");
          }

          const ops: EngineOp[] = [];
          for (let row = sourceBounds.startRow; row <= sourceBounds.endRow; row += 1) {
            for (let col = sourceBounds.startCol; col <= sourceBounds.endCol; col += 1) {
              ops.push({
                kind: "clearCell",
                sheetName: source.sheetName,
                address: formatAddress(row, col),
              });
            }
          }
          for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
              const nextAddress = formatAddress(
                targetBounds.startRow + rowOffset,
                targetBounds.startCol + colOffset,
              );
              const sourceAddress = formatAddress(
                sourceBounds.startRow + rowOffset,
                sourceBounds.startCol + colOffset,
              );
              ops.push(
                ...args.toCellStateOps(
                  target.sheetName,
                  nextAddress,
                  getMatrixCell(sourceMatrix, rowOffset, colOffset),
                  source.sheetName,
                  sourceAddress,
                ),
              );
            }
          }
          Effect.runSync(this.executeLocal(ops, ops.length));
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage("Failed to move range", cause),
            cause,
          }),
      });
    },
    importSheetCsv(sheetName, csv) {
      return Effect.try({
        try: () => {
          const rows = parseCsv(csv);
          const existingSheet = args.state.workbook.getSheet(sheetName);
          const order = existingSheet?.order ?? args.state.workbook.sheetsByName.size;
          const ops: EngineOp[] = [];
          let potentialNewCells = 0;

          if (existingSheet) {
            ops.push({ kind: "deleteSheet", name: sheetName });
          }
          ops.push({ kind: "upsertSheet", name: sheetName, order });

          rows.forEach((row, rowIndex) => {
            row.forEach((raw, colIndex) => {
              const parsed = parseCsvCellInput(raw);
              if (!parsed) {
                return;
              }
              const address = formatAddress(rowIndex, colIndex);
              if (parsed.formula !== undefined) {
                ops.push({ kind: "setCellFormula", sheetName, address, formula: parsed.formula });
                potentialNewCells += 1;
                return;
              }
              ops.push({ kind: "setCellValue", sheetName, address, value: parsed.value ?? null });
              potentialNewCells += 1;
            });
          });

          Effect.runSync(this.executeLocal(ops, potentialNewCells));
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage("Failed to import sheet CSV", cause),
            cause,
          }),
      });
    },
    renderCommit(ops) {
      return Effect.flatMap(
        Effect.try({
          try: () => {
            const engineOps: EngineOp[] = [];
            let potentialNewCells = 0;
            ops.forEach((op) => {
              switch (op.kind) {
                case "upsertWorkbook":
                  if (op.name) {
                    engineOps.push({ kind: "upsertWorkbook", name: op.name });
                  }
                  break;
                case "upsertSheet":
                  if (op.name) {
                    engineOps.push({ kind: "upsertSheet", name: op.name, order: op.order ?? 0 });
                  }
                  break;
                case "renameSheet":
                  if (op.oldName && op.newName) {
                    engineOps.push({
                      kind: "renameSheet",
                      oldName: op.oldName,
                      newName: op.newName,
                    });
                  }
                  break;
                case "deleteSheet":
                  if (op.name) {
                    engineOps.push({ kind: "deleteSheet", name: op.name });
                  }
                  break;
                case "upsertCell":
                  if (!op.sheetName || !op.addr) {
                    break;
                  }
                  if (op.formula !== undefined) {
                    engineOps.push({
                      kind: "setCellFormula",
                      sheetName: op.sheetName,
                      address: op.addr,
                      formula: op.formula,
                    });
                  } else {
                    engineOps.push({
                      kind: "setCellValue",
                      sheetName: op.sheetName,
                      address: op.addr,
                      value: op.value ?? null,
                    });
                  }
                  potentialNewCells += 1;
                  if (op.format !== undefined) {
                    engineOps.push({
                      kind: "setCellFormat",
                      sheetName: op.sheetName,
                      address: op.addr,
                      format: op.format,
                    });
                  }
                  break;
                case "deleteCell":
                  if (op.sheetName && op.addr) {
                    engineOps.push({
                      kind: "clearCell",
                      sheetName: op.sheetName,
                      address: op.addr,
                    });
                    engineOps.push({
                      kind: "setCellFormat",
                      sheetName: op.sheetName,
                      address: op.addr,
                      format: null,
                    });
                  }
                  break;
              }
            });
            return { engineOps, potentialNewCells };
          },
          catch: (cause) =>
            new EngineMutationError({
              message: "Failed to normalize render commit operations",
              cause,
            }),
        }),
        ({ engineOps, potentialNewCells }) => this.executeLocal(engineOps, potentialNewCells),
      ).pipe(Effect.asVoid);
    },
  };
}
