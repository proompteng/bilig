import type { ErrorCode, FormulaMode, ValueTag } from "./enums.js";

export type CellIndex = number;
export type FormulaId = number;
export type RangeIndex = number;
export type EntityId = number;
export type LiteralInput = number | string | boolean | null;

export type EmptyValue = { tag: ValueTag.Empty };
export type NumberValue = { tag: ValueTag.Number; value: number };
export type BooleanValue = { tag: ValueTag.Boolean; value: boolean };
export type StringValue = { tag: ValueTag.String; value: string; stringId: number };
export type ErrorValue = { tag: ValueTag.Error; code: ErrorCode };

export type CellValue = EmptyValue | NumberValue | BooleanValue | StringValue | ErrorValue;

export interface CellSnapshot {
  sheetName: string;
  address: string;
  formula?: string;
  format?: string;
  input?: LiteralInput;
  value: CellValue;
  flags: number;
}

export interface DependencySnapshot {
  directPrecedents: string[];
  directDependents: string[];
}

export interface ExplainCellSnapshot {
  sheetName: string;
  address: string;
  formula?: string;
  format?: string;
  mode?: FormulaMode;
  value: CellValue;
  flags: number;
  version: number;
  topoRank?: number;
  inCycle: boolean;
  directPrecedents: string[];
  directDependents: string[];
}

export interface RecalcMetrics {
  batchId: number;
  changedInputCount: number;
  dirtyFormulaCount: number;
  wasmFormulaCount: number;
  jsFormulaCount: number;
  rangeNodeVisits: number;
  recalcMs: number;
  compileMs: number;
}

export interface EngineEvent {
  kind: "batch";
  changedCellIndices: Uint32Array | number[];
  metrics: RecalcMetrics;
}

export interface SelectionState {
  sheetName: string;
  address: string | null;
}

export interface Viewport {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

export interface WorkbookSnapshot {
  version: 1;
  workbook: {
    name: string;
  };
  sheets: Array<{
    name: string;
    order: number;
    cells: Array<{
      address: string;
      value?: LiteralInput;
      formula?: string;
      format?: string;
    }>;
  }>;
}

export interface FormulaRecord {
  id: FormulaId;
  source: string;
  mode: FormulaMode;
  programOffset: number;
  programLength: number;
  constNumberOffset: number;
  constNumberLength: number;
  rangeListOffset: number;
  rangeListLength: number;
  program: Uint32Array;
  constants: number[];
  symbolicRefs: string[];
  symbolicRanges: string[];
  maxStackDepth: number;
}
