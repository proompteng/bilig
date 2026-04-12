import {
  loadLiteralSheetIntoEmptySheet,
  type SpreadsheetEngine,
  type EngineCellMutationRef,
} from "@bilig/core";
import type { WorkPaperCellAddress, WorkPaperSheet } from "./work-paper-types.js";

export function tryLoadInitialLiteralSheet(
  engine: SpreadsheetEngine,
  sheetId: number,
  content: WorkPaperSheet,
): boolean {
  if (sheetContainsFormulaContent(content)) {
    return false;
  }
  loadLiteralSheetIntoEmptySheet(engine.workbook, engine.strings, sheetId, content);
  return true;
}

function sheetContainsFormulaContent(content: WorkPaperSheet): boolean {
  return content.some((row) =>
    row.some((value) => typeof value === "string" && value.trim().startsWith("=")),
  );
}

export function loadInitialMixedSheet(args: {
  engine: SpreadsheetEngine;
  sheetId: number;
  content: WorkPaperSheet;
  rewriteFormula: (formula: string, destination: WorkPaperCellAddress) => string;
}): void {
  loadLiteralSheetIntoEmptySheet(
    args.engine.workbook,
    args.engine.strings,
    args.sheetId,
    args.content,
    (value: WorkPaperSheet[number][number]) =>
      !(typeof value === "string" && value.trim().startsWith("=")),
  );

  const formulaRefs: EngineCellMutationRef[] = [];
  args.content.forEach((row, rowOffset) => {
    row.forEach((raw, colOffset) => {
      if (typeof raw !== "string" || !raw.trim().startsWith("=")) {
        return;
      }
      formulaRefs.push({
        sheetId: args.sheetId,
        mutation: {
          kind: "setCellFormula",
          row: rowOffset,
          col: colOffset,
          formula: args.rewriteFormula(raw.trim().slice(1), {
            sheet: args.sheetId,
            row: rowOffset,
            col: colOffset,
          }),
        },
      });
    });
  });
  if (formulaRefs.length === 0) {
    return;
  }
  args.engine.applyCellMutationsAtWithOptions(formulaRefs, {
    captureUndo: false,
    source: "restore",
    potentialNewCells: formulaRefs.length,
    returnUndoOps: false,
    reuseRefs: true,
  });
}
