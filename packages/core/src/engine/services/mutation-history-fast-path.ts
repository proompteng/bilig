import { ValueTag, type CellSnapshot } from "@bilig/protocol";
import type { EngineOp } from "@bilig/workbook-domain";
import type { WorkbookStore } from "../../workbook-store.js";
import type { TransactionRecord } from "../runtime-state.js";

type FastSimpleCellOp = Extract<
  EngineOp,
  | { kind: "setCellValue" }
  | { kind: "setCellFormula" }
  | { kind: "clearCell" }
  | { kind: "setCellFormat" }
>;

interface FastMutationHistoryResult {
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

function isFastSimpleCellOp(op: EngineOp): op is FastSimpleCellOp {
  return (
    op.kind === "setCellValue" ||
    op.kind === "setCellFormula" ||
    op.kind === "clearCell" ||
    op.kind === "setCellFormat"
  );
}

function cloneSimpleCellOp(op: FastSimpleCellOp): FastSimpleCellOp {
  switch (op.kind) {
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

function restoreCellOpFromSnapshot(
  workbook: WorkbookStore,
  getCellByIndex: (cellIndex: number) => CellSnapshot,
  sheetName: string,
  address: string,
): FastSimpleCellOp {
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

function buildSimpleCellInverseOp(
  workbook: WorkbookStore,
  getCellByIndex: (cellIndex: number) => CellSnapshot,
  op: FastSimpleCellOp,
): FastSimpleCellOp | null {
  switch (op.kind) {
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
  const forwardOps = Array<FastSimpleCellOp>(args.ops.length);
  for (let index = 0; index < args.ops.length; index += 1) {
    const op = args.ops[index];
    if (op === undefined || !isFastSimpleCellOp(op)) {
      return null;
    }
    forwardOps[index] = cloneSimpleCellOp(op);
  }

  const inverseOps: FastSimpleCellOp[] = [];
  for (let index = forwardOps.length - 1; index >= 0; index -= 1) {
    const op = forwardOps[index];
    if (op === undefined) {
      return null;
    }
    const inverse = buildSimpleCellInverseOp(args.workbook, args.getCellByIndex, op);
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
    undoOps: inverseOps.map((op) => cloneSimpleCellOp(op)),
  };
}
