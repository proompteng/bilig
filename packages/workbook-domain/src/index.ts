export * from "./guards.js";

import type {
  CellRangeRef,
  CellNumberFormatRecord,
  CellStyleRecord,
  CellStylePatch,
  LiteralInput,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookNoteSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookVolatileContextSnapshot,
} from "@bilig/protocol";

export type ReplicaId = string;
export type OpId = string;

export interface Clock {
  counter: number;
}

export type WorkbookStructuralAxis = "row" | "column";
export type WorkbookSortDirection = "asc" | "desc";

export interface WorkbookTableOp {
  name: string;
  sheetName: string;
  startAddress: string;
  endAddress: string;
  columnNames: string[];
  headerRow: boolean;
  totalsRow: boolean;
}

export interface WorkbookSortKey {
  keyAddress: string;
  direction: WorkbookSortDirection;
}

export interface WorkbookAxisEntryOp extends WorkbookAxisEntrySnapshot {}
export interface WorkbookCellStyleOp extends CellStyleRecord {}
export interface WorkbookCellNumberFormatOp extends CellNumberFormatRecord {}
export interface WorkbookDataValidationOp extends WorkbookDataValidationSnapshot {}
export interface WorkbookConditionalFormatOp extends WorkbookConditionalFormatSnapshot {
  style: CellStylePatch;
}
export interface WorkbookCommentThreadOp extends WorkbookCommentThreadSnapshot {}
export interface WorkbookNoteOp extends WorkbookNoteSnapshot {}

export type WorkbookOp =
  | { kind: "upsertWorkbook"; name: string }
  | { kind: "setWorkbookMetadata"; key: string; value: LiteralInput }
  | { kind: "setCalculationSettings"; settings: WorkbookCalculationSettingsSnapshot }
  | { kind: "setVolatileContext"; context: WorkbookVolatileContextSnapshot }
  | { kind: "upsertSheet"; name: string; order: number; id?: number }
  | { kind: "renameSheet"; oldName: string; newName: string }
  | { kind: "deleteSheet"; name: string }
  | {
      kind: "insertRows";
      sheetName: string;
      start: number;
      count: number;
      entries?: WorkbookAxisEntryOp[];
    }
  | { kind: "deleteRows"; sheetName: string; start: number; count: number }
  | { kind: "moveRows"; sheetName: string; start: number; count: number; target: number }
  | {
      kind: "insertColumns";
      sheetName: string;
      start: number;
      count: number;
      entries?: WorkbookAxisEntryOp[];
    }
  | { kind: "deleteColumns"; sheetName: string; start: number; count: number }
  | { kind: "moveColumns"; sheetName: string; start: number; count: number; target: number }
  | {
      kind: "updateRowMetadata";
      sheetName: string;
      start: number;
      count: number;
      size: number | null;
      hidden: boolean | null;
    }
  | {
      kind: "updateColumnMetadata";
      sheetName: string;
      start: number;
      count: number;
      size: number | null;
      hidden: boolean | null;
    }
  | { kind: "setFreezePane"; sheetName: string; rows: number; cols: number }
  | { kind: "clearFreezePane"; sheetName: string }
  | { kind: "setFilter"; sheetName: string; range: CellRangeRef }
  | { kind: "clearFilter"; sheetName: string; range: CellRangeRef }
  | { kind: "setSort"; sheetName: string; range: CellRangeRef; keys: WorkbookSortKey[] }
  | { kind: "clearSort"; sheetName: string; range: CellRangeRef }
  | { kind: "setDataValidation"; validation: WorkbookDataValidationOp }
  | { kind: "clearDataValidation"; sheetName: string; range: CellRangeRef }
  | { kind: "upsertConditionalFormat"; format: WorkbookConditionalFormatOp }
  | { kind: "deleteConditionalFormat"; id: string; sheetName: string }
  | { kind: "upsertCommentThread"; thread: WorkbookCommentThreadOp }
  | { kind: "deleteCommentThread"; sheetName: string; address: string }
  | { kind: "upsertNote"; note: WorkbookNoteOp }
  | { kind: "deleteNote"; sheetName: string; address: string }
  | { kind: "setCellValue"; sheetName: string; address: string; value: LiteralInput }
  | { kind: "setCellFormula"; sheetName: string; address: string; formula: string }
  | { kind: "setCellFormat"; sheetName: string; address: string; format: string | null }
  | { kind: "upsertCellStyle"; style: WorkbookCellStyleOp }
  | { kind: "upsertCellNumberFormat"; format: WorkbookCellNumberFormatOp }
  | { kind: "setStyleRange"; range: CellRangeRef; styleId: string }
  | { kind: "setFormatRange"; range: CellRangeRef; formatId: string }
  | { kind: "clearCell"; sheetName: string; address: string }
  | { kind: "upsertDefinedName"; name: string; value: WorkbookDefinedNameValueSnapshot }
  | { kind: "deleteDefinedName"; name: string }
  | { kind: "upsertTable"; table: WorkbookTableOp }
  | { kind: "deleteTable"; name: string }
  | { kind: "upsertSpillRange"; sheetName: string; address: string; rows: number; cols: number }
  | { kind: "deleteSpillRange"; sheetName: string; address: string }
  | {
      kind: "upsertPivotTable";
      name: string;
      sheetName: string;
      address: string;
      source: CellRangeRef;
      groupBy: string[];
      values: WorkbookPivotValueSnapshot[];
      rows: number;
      cols: number;
    }
  | { kind: "deletePivotTable"; sheetName: string; address: string };

export type EngineOp = WorkbookOp;

export interface WorkbookOpBatch {
  id: OpId;
  replicaId: ReplicaId;
  clock: Clock;
  ops: WorkbookOp[];
}

export type EngineOpBatch = WorkbookOpBatch;

export interface WorkbookTxn {
  ops: WorkbookOp[];
  potentialNewCells?: number;
}
