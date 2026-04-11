import { formatAddress } from "@bilig/formula";
import type { LiteralInput } from "@bilig/protocol";
import type { HeadlessCellAddress, HeadlessSheet, RawCellContent } from "./types.js";

export type MatrixMutationOp =
  | { kind: "clearCell"; sheetName: string; address: string }
  | { kind: "setCellValue"; sheetName: string; address: string; value: LiteralInput }
  | { kind: "setCellFormula"; sheetName: string; address: string; formula: string };

export interface MatrixMutationPlan {
  leadingOps: MatrixMutationOp[];
  formulaOps: MatrixMutationOp[];
  ops: MatrixMutationOp[];
  potentialNewCells: number;
  trailingLiteralOps: MatrixMutationOp[];
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
  deferLiteralAddresses?: ReadonlySet<string>;
  skipNulls?: boolean;
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === "string" && content.trim().startsWith("=");
}

export function buildMatrixMutationPlan(args: BuildMatrixMutationPlanArgs): MatrixMutationPlan {
  const leadingOps: MatrixMutationOp[] = [];
  const formulaOps: MatrixMutationOp[] = [];
  const trailingLiteralOps: MatrixMutationOp[] = [];
  let potentialNewCells = 0;
  const formulaDestinations: Array<{ row: number; col: number }> = [];

  args.content.forEach((row, rowOffset) => {
    row.forEach((raw, columnOffset) => {
      if (!isFormulaContent(raw)) {
        return;
      }
      formulaDestinations.push({
        row: args.target.row + rowOffset,
        col: args.target.col + columnOffset,
      });
    });
  });

  const shouldDeferLiteral = (address: string, row: number, col: number): boolean =>
    args.deferLiteralAddresses?.has(address) === true ||
    formulaDestinations.some(
      (formulaDestination) => row > formulaDestination.row && col === formulaDestination.col,
    );

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
          const op = { kind: "clearCell", sheetName: args.targetSheetName, address } as const;
          if (shouldDeferLiteral(address, destination.row, destination.col)) {
            trailingLiteralOps.push(op);
          } else {
            leadingOps.push(op);
          }
        }
        return;
      }

      potentialNewCells += 1;

      if (isFormulaContent(raw)) {
        formulaOps.push({
          kind: "setCellFormula",
          sheetName: args.targetSheetName,
          address,
          formula: args.rewriteFormula(raw, destination, rowOffset, columnOffset),
        });
        return;
      }

      const op = {
        kind: "setCellValue",
        sheetName: args.targetSheetName,
        address,
        value: raw,
      } as const;
      if (shouldDeferLiteral(address, destination.row, destination.col)) {
        trailingLiteralOps.push(op);
      } else {
        leadingOps.push(op);
      }
    });
  });

  return {
    leadingOps,
    formulaOps,
    ops: [...leadingOps, ...formulaOps, ...trailingLiteralOps],
    potentialNewCells,
    trailingLiteralOps,
  };
}
