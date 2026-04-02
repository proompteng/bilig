import type { CellRangeRef, CellStyleRecord, LiteralInput } from "@bilig/protocol";

export type ClassACommand =
  | { kind: "EditCellsBatch"; sheetId: string; range: CellRangeRef; values: LiteralInput[][] }
  | { kind: "SetFormulasBatch"; sheetId: string; range: CellRangeRef; formulas: string[][] }
  | {
      kind: "ApplyStyleBatch";
      sheetId: string;
      range: CellRangeRef;
      style: Partial<CellStyleRecord>;
    }
  | { kind: "RenameSheet"; sheetId: string; newName: string }
  | { kind: "ResizeColumn"; sheetId: string; colNum: number; size: number };

export type ClassBCommand =
  | { kind: "PasteCells"; sheetId: string; target: CellRangeRef; sourceData: unknown }
  | { kind: "InsertRows"; sheetId: string; startIndex: number; count: number }
  | { kind: "DeleteColumns"; sheetId: string; startIndex: number; count: number }
  | { kind: "SortRange"; sheetId: string; range: CellRangeRef; config: unknown };

export type ClassCCommand =
  | { kind: "ImportXlsx"; fileId: string }
  | { kind: "CreatePivot"; sheetId: string; config: unknown };

export type CanonicalCommand = ClassACommand | ClassBCommand | ClassCCommand;

export interface CommandBundle {
  idempotencyKey: string;
  workbookId: string;
  scope: "selection" | "sheet" | "workbook";
  commands: CanonicalCommand[];
  undoLabel?: string;
  userId: string;
}
