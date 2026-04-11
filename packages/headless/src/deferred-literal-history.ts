import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag } from "@bilig/protocol";
import type { EngineOp } from "@bilig/workbook-domain";

type DeferredLiteralOp = Extract<EngineOp, { kind: "setCellValue" } | { kind: "clearCell" }>;

export interface DeferredLiteralHistoryRecord {
  forward: {
    ops: unknown[];
    potentialNewCells?: number;
  };
  inverse: {
    ops: unknown[];
    potentialNewCells?: number;
  };
}

function cloneDeferredLiteralOp(op: DeferredLiteralOp): EngineOp {
  switch (op.kind) {
    case "setCellValue":
      return {
        kind: "setCellValue",
        sheetName: op.sheetName,
        address: op.address,
        value: op.value,
      };
    case "clearCell":
      return {
        kind: "clearCell",
        sheetName: op.sheetName,
        address: op.address,
      };
  }
}

function restoreCellOp(engine: SpreadsheetEngine, sheetName: string, address: string): EngineOp {
  const cellIndex = engine.workbook.getCellIndex(sheetName, address);
  if (cellIndex === undefined) {
    return { kind: "clearCell", sheetName, address };
  }

  const cell = engine.getCellByIndex(cellIndex);
  if (cell.formula !== undefined) {
    return {
      kind: "setCellFormula",
      sheetName,
      address,
      formula: cell.formula,
    };
  }

  switch (cell.value.tag) {
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
        value: cell.value.value,
      };
  }
}

export function buildDeferredLiteralHistoryRecord(
  engine: SpreadsheetEngine,
  ops: readonly EngineOp[],
  potentialNewCells?: number,
): DeferredLiteralHistoryRecord | null {
  const literalOps: DeferredLiteralOp[] = [];
  for (let index = 0; index < ops.length; index += 1) {
    const op = ops[index];
    if (!op || (op.kind !== "setCellValue" && op.kind !== "clearCell")) {
      return null;
    }
    literalOps.push(op);
  }

  return {
    forward:
      potentialNewCells === undefined
        ? { ops: literalOps.map((op) => cloneDeferredLiteralOp(op)) }
        : { ops: literalOps.map((op) => cloneDeferredLiteralOp(op)), potentialNewCells },
    inverse: {
      ops: literalOps.toReversed().map((op) => restoreCellOp(engine, op.sheetName, op.address)),
      potentialNewCells: literalOps.length,
    },
  };
}
