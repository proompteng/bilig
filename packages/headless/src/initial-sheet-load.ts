import { loadLiteralSheetIntoEmptySheet, type SpreadsheetEngine } from "@bilig/core";
import type { HeadlessSheet } from "./types.js";

export function tryLoadInitialLiteralSheet(
  engine: SpreadsheetEngine,
  sheetId: number,
  content: HeadlessSheet,
): boolean {
  if (sheetContainsFormulaContent(content)) {
    return false;
  }
  loadLiteralSheetIntoEmptySheet(engine.workbook, engine.strings, sheetId, content);
  return true;
}

function sheetContainsFormulaContent(content: HeadlessSheet): boolean {
  return content.some((row) =>
    row.some((value) => typeof value === "string" && value.trim().startsWith("=")),
  );
}
