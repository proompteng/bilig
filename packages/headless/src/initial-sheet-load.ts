import { loadLiteralSheetIntoEmptySheet, type SpreadsheetEngine } from "@bilig/core";
import type { WorkPaperSheet } from "./work-paper-types.js";

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
