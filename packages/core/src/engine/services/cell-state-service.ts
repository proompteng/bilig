import { Effect } from "effect";
import {
  ValueTag,
  type CellRangeRef,
  type CellSnapshot,
} from "@bilig/protocol";
import {
  formatAddress,
  parseCellAddress,
  translateFormulaReferences,
} from "@bilig/formula";
import type { EngineOp } from "@bilig/workbook-domain";
import { normalizeRange } from "../../engine-range-utils.js";
import type { EngineRuntimeState } from "../runtime-state.js";
import { EngineCellStateError } from "../errors.js";

export interface EngineCellStateService {
  readonly restoreCellOps: (
    sheetName: string,
    address: string,
  ) => Effect.Effect<EngineOp[], EngineCellStateError>;
  readonly readRangeCells: (
    range: CellRangeRef,
  ) => Effect.Effect<CellSnapshot[][], EngineCellStateError>;
  readonly toCellStateOps: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => Effect.Effect<EngineOp[], EngineCellStateError>;
}

function cellStateErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

function translateFormulaForTarget(
  formula: string,
  sourceSheetName: string,
  sourceAddress: string,
  targetSheetName: string,
  targetAddress: string,
): string {
  const source = parseCellAddress(sourceAddress, sourceSheetName);
  const target = parseCellAddress(targetAddress, targetSheetName);
  return translateFormulaReferences(formula, target.row - source.row, target.col - source.col);
}

export function createEngineCellStateService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook">;
  readonly getCell: (sheetName: string, address: string) => CellSnapshot;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
}): EngineCellStateService {
  const toCellStateOpsNow = (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): EngineOp[] => {
    const ops: EngineOp[] = [];
    if (snapshot.formula !== undefined) {
      const translatedFormula =
        sourceSheetName && sourceAddress
          ? translateFormulaForTarget(
              snapshot.formula,
              sourceSheetName,
              sourceAddress,
              sheetName,
              address,
            )
          : snapshot.formula;
      ops.push({ kind: "setCellFormula", sheetName, address, formula: translatedFormula });
    } else {
      switch (snapshot.value.tag) {
        case ValueTag.Empty:
          ops.push({ kind: "clearCell", sheetName, address });
          break;
        case ValueTag.Number:
        case ValueTag.Boolean:
        case ValueTag.String:
          ops.push({ kind: "setCellValue", sheetName, address, value: snapshot.value.value });
          break;
        case ValueTag.Error:
          ops.push({ kind: "clearCell", sheetName, address });
          break;
      }
    }
    ops.push({
      kind: "setCellFormat",
      sheetName,
      address,
      format: snapshot.format ?? null,
    });
    return ops;
  };

  const restoreCellOpsNow = (sheetName: string, address: string): EngineOp[] => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return [{ kind: "clearCell", sheetName, address }];
    }
    return toCellStateOpsNow(sheetName, address, args.getCellByIndex(cellIndex)).filter(
      (op) => op.kind !== "setCellFormat",
    );
  };

  const readRangeCellsNow = (range: CellRangeRef): CellSnapshot[][] => {
    const bounds = normalizeRange(range);
    const rows: CellSnapshot[][] = [];
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      const cells: CellSnapshot[] = [];
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        cells.push(args.getCell(range.sheetName, formatAddress(row, col)));
      }
      rows.push(cells);
    }
    return rows;
  };

  return {
    restoreCellOps(sheetName, address) {
      return Effect.try({
        try: () => restoreCellOpsNow(sheetName, address),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(
              `Failed to restore cell ops for ${sheetName}!${address}`,
              cause,
            ),
            cause,
          }),
      });
    },
    readRangeCells(range) {
      return Effect.try({
        try: () => readRangeCellsNow(range),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(
              `Failed to read range ${range.sheetName}!${range.startAddress}:${range.endAddress}`,
              cause,
            ),
            cause,
          }),
      });
    },
    toCellStateOps(sheetName, address, snapshot, sourceSheetName, sourceAddress) {
      return Effect.try({
        try: () =>
          toCellStateOpsNow(
            sheetName,
            address,
            snapshot,
            sourceSheetName,
            sourceAddress,
          ),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(
              `Failed to materialize cell state ops for ${sheetName}!${address}`,
              cause,
            ),
            cause,
          }),
      });
    },
  };
}
