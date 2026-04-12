import { ValueTag, type CellSnapshot } from "@bilig/protocol";
import type { EngineOp } from "@bilig/workbook-domain";
import type { WorkbookStore } from "../../workbook-store.js";
import type { TransactionRecord } from "../runtime-state.js";

type FastHistoryOp = Extract<
  EngineOp,
  | { kind: "upsertWorkbook" }
  | { kind: "upsertSheet" }
  | { kind: "renameSheet" }
  | { kind: "setCellValue" }
  | { kind: "setCellFormula" }
  | { kind: "clearCell" }
  | { kind: "setCellFormat" }
>;

type FastHistoryCloneOp = FastHistoryOp | Extract<EngineOp, { kind: "deleteSheet" }>;

export interface FastMutationHistoryResult {
  forward: TransactionRecord;
  inverse: TransactionRecord;
  undoOps: EngineOp[];
}

interface FastMutationHistoryArgs {
  readonly workbook: WorkbookStore;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
  readonly ops: readonly EngineOp[];
  readonly potentialNewCells?: number;
}

function isFastHistoryOp(op: EngineOp): op is FastHistoryOp {
  return (
    op.kind === "upsertWorkbook" ||
    op.kind === "upsertSheet" ||
    op.kind === "renameSheet" ||
    op.kind === "setCellValue" ||
    op.kind === "setCellFormula" ||
    op.kind === "clearCell" ||
    op.kind === "setCellFormat"
  );
}

function cloneFastHistoryForwardOp(op: FastHistoryOp): FastHistoryOp {
  switch (op.kind) {
    case "upsertWorkbook":
      return {
        kind: "upsertWorkbook",
        name: op.name,
      };
    case "upsertSheet":
      return {
        kind: "upsertSheet",
        name: op.name,
        order: op.order,
      };
    case "renameSheet":
      return {
        kind: "renameSheet",
        oldName: op.oldName,
        newName: op.newName,
      };
    case "setCellValue":
      return {
        kind: "setCellValue",
        sheetName: op.sheetName,
        address: op.address,
        value: op.value,
      };
    case "setCellFormula":
      return {
        kind: "setCellFormula",
        sheetName: op.sheetName,
        address: op.address,
        formula: op.formula,
      };
    case "clearCell":
      return {
        kind: "clearCell",
        sheetName: op.sheetName,
        address: op.address,
      };
    case "setCellFormat":
      return {
        kind: "setCellFormat",
        sheetName: op.sheetName,
        address: op.address,
        format: op.format,
      };
  }
}

function cloneFastHistoryUndoOp(op: FastHistoryCloneOp): FastHistoryCloneOp {
  if (op.kind === "deleteSheet") {
    return {
      kind: "deleteSheet",
      name: op.name,
    };
  }
  return cloneFastHistoryForwardOp(op);
}

function restoreCellOpFromSnapshot(
  workbook: WorkbookStore,
  getCellByIndex: (cellIndex: number) => CellSnapshot,
  sheetName: string,
  address: string,
): Extract<FastHistoryOp, { kind: "setCellValue" | "setCellFormula" | "clearCell" }> {
  const cellIndex = workbook.getCellIndex(sheetName, address);
  if (cellIndex === undefined) {
    return { kind: "clearCell", sheetName, address };
  }

  const snapshot = getCellByIndex(cellIndex);
  if (snapshot.formula !== undefined) {
    return {
      kind: "setCellFormula",
      sheetName,
      address,
      formula: snapshot.formula,
    };
  }

  switch (snapshot.value.tag) {
    case ValueTag.Empty:
    case ValueTag.Error:
      return { kind: "clearCell", sheetName, address };
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return {
        kind: "setCellValue",
        sheetName,
        address,
        value: snapshot.value.value,
      };
  }
}

function buildFastInverseOp(
  workbook: WorkbookStore,
  getCellByIndex: (cellIndex: number) => CellSnapshot,
  op: FastHistoryOp,
): FastHistoryCloneOp | null {
  switch (op.kind) {
    case "upsertWorkbook":
      return {
        kind: "upsertWorkbook",
        name: workbook.workbookName,
      };
    case "upsertSheet": {
      const existing = workbook.getSheet(op.name);
      if (!existing) {
        return { kind: "deleteSheet", name: op.name };
      }
      return {
        kind: "upsertSheet",
        name: existing.name,
        order: existing.order,
      };
    }
    case "renameSheet": {
      const existing = workbook.getSheet(op.newName);
      if (!existing) {
        return null;
      }
      return {
        kind: "renameSheet",
        oldName: op.newName,
        newName: op.oldName,
      };
    }
    case "setCellValue":
    case "setCellFormula":
    case "clearCell":
      return restoreCellOpFromSnapshot(workbook, getCellByIndex, op.sheetName, op.address);
    case "setCellFormat": {
      const cellIndex = workbook.getCellIndex(op.sheetName, op.address);
      return {
        kind: "setCellFormat",
        sheetName: op.sheetName,
        address: op.address,
        format: cellIndex === undefined ? null : (workbook.getCellFormat(cellIndex) ?? null),
      };
    }
    default:
      return null;
  }
}

export function tryBuildFastMutationHistory(
  args: FastMutationHistoryArgs,
): FastMutationHistoryResult | null {
  const forwardOps = Array<FastHistoryOp>(args.ops.length);
  for (let index = 0; index < args.ops.length; index += 1) {
    const op = args.ops[index];
    if (op === undefined || !isFastHistoryOp(op)) {
      return null;
    }
    forwardOps[index] = cloneFastHistoryForwardOp(op);
  }

  const inverseOps: EngineOp[] = [];
  for (let index = forwardOps.length - 1; index >= 0; index -= 1) {
    const op = forwardOps[index];
    if (op === undefined) {
      return null;
    }
    const inverse = buildFastInverseOp(args.workbook, args.getCellByIndex, op);
    if (!inverse) {
      return null;
    }
    inverseOps.push(inverse);
  }

  return {
    forward:
      args.potentialNewCells === undefined
        ? { ops: forwardOps }
        : { ops: forwardOps, potentialNewCells: args.potentialNewCells },
    inverse: {
      ops: inverseOps,
      potentialNewCells: args.ops.length,
    },
    undoOps: inverseOps.map((op) => {
      if (op.kind === "deleteSheet") {
        return cloneFastHistoryUndoOp(op);
      }
      if (!isFastHistoryOp(op)) {
        throw new TypeError(`Unsupported fast-path undo op: ${op.kind}`);
      }
      return cloneFastHistoryUndoOp(op);
    }),
  };
}
