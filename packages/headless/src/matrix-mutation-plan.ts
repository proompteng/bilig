import { formatAddress } from "@bilig/formula";
import type { LiteralInput } from "@bilig/protocol";
import type { HeadlessCellAddress, HeadlessSheet, RawCellContent } from "./types.js";

export type MatrixMutationOp =
  | { kind: "clearCell"; sheetName: string; address: string }
  | { kind: "setCellValue"; sheetName: string; address: string; value: LiteralInput }
  | { kind: "setCellFormula"; sheetName: string; address: string; formula: string };

export interface MatrixMutationPlan {
  ops: MatrixMutationOp[];
  potentialNewCells: number;
}

interface BuildMatrixMutationPlanArgs {
  target: HeadlessCellAddress;
  targetSheetName: string;
  content: HeadlessSheet;
  rewriteFormula: (
    formula: string,
    destination: HeadlessCellAddress,
    rowOffset: number,
    columnOffset: number,
  ) => string;
  skipNulls?: boolean;
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === "string" && content.trim().startsWith("=");
}

export function buildMatrixMutationPlan(args: BuildMatrixMutationPlanArgs): MatrixMutationPlan {
  const ops: MatrixMutationOp[] = [];
  let potentialNewCells = 0;

  args.content.forEach((row, rowOffset) => {
    row.forEach((raw, columnOffset) => {
      const destination: HeadlessCellAddress = {
        sheet: args.target.sheet,
        row: args.target.row + rowOffset,
        col: args.target.col + columnOffset,
      };
      const address = formatAddress(destination.row, destination.col);

      if (raw === null) {
        if (!args.skipNulls) {
          ops.push({ kind: "clearCell", sheetName: args.targetSheetName, address });
        }
        return;
      }

      potentialNewCells += 1;

      if (isFormulaContent(raw)) {
        ops.push({
          kind: "setCellFormula",
          sheetName: args.targetSheetName,
          address,
          formula: args.rewriteFormula(raw, destination, rowOffset, columnOffset),
        });
        return;
      }

      ops.push({
        kind: "setCellValue",
        sheetName: args.targetSheetName,
        address,
        value: raw,
      });
    });
  });

  return { ops, potentialNewCells };
}
