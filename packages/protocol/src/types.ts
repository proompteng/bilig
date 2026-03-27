import { ErrorCode, type FormulaMode, type ValueTag } from "./enums.js";

export type CellIndex = number;
export type FormulaId = number;
export type RangeIndex = number;
export type EntityId = number;
export type LiteralInput = number | string | boolean | null;
export type CompatibilityMode = "excel-modern" | "odf-1.4";

export type EmptyValue = { tag: ValueTag.Empty };
export type NumberValue = { tag: ValueTag.Number; value: number };
export type BooleanValue = { tag: ValueTag.Boolean; value: boolean };
export type StringValue = { tag: ValueTag.String; value: string; stringId: number };
export type ErrorValue = { tag: ValueTag.Error; code: ErrorCode };

export type CellValue = EmptyValue | NumberValue | BooleanValue | StringValue | ErrorValue;

export function formatErrorCode(code: ErrorCode): string {
  switch (code) {
    case ErrorCode.None:
      return "#ERROR!";
    case ErrorCode.Div0:
      return "#DIV/0!";
    case ErrorCode.Ref:
      return "#REF!";
    case ErrorCode.Value:
      return "#VALUE!";
    case ErrorCode.Name:
      return "#NAME?";
    case ErrorCode.NA:
      return "#N/A";
    case ErrorCode.Cycle:
      return "#CYCLE!";
    case ErrorCode.Spill:
      return "#SPILL!";
    case ErrorCode.Blocked:
      return "#BLOCKED!";
    default:
      return "#ERROR!";
  }
}

export interface CellSnapshot {
  sheetName: string;
  address: string;
  formula?: string;
  format?: string;
  input?: LiteralInput;
  value: CellValue;
  flags: number;
  version: number;
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

export interface CellRangeRef {
  sheetName: string;
  startAddress: string;
  endAddress: string;
}

export interface SelectionRange {
  startAddress: string;
  endAddress: string;
}

export type SelectionEditMode = "idle" | "cell" | "formula";
export type SyncState = "local-only" | "syncing" | "live" | "behind" | "reconnecting";

export interface SelectionState {
  sheetName: string;
  address: string | null;
  anchorAddress: string | null;
  range: SelectionRange | null;
  editMode: SelectionEditMode;
}

export interface Viewport {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

export interface WorkbookDefinedNameSnapshot {
  name: string;
  value: WorkbookDefinedNameValueSnapshot;
}

export interface WorkbookDefinedNameScalarValueSnapshot {
  kind: "scalar";
  value: LiteralInput;
}

export interface WorkbookDefinedNameCellRefValueSnapshot {
  kind: "cell-ref";
  sheetName: string;
  address: string;
}

export interface WorkbookDefinedNameRangeRefValueSnapshot {
  kind: "range-ref";
  sheetName: string;
  startAddress: string;
  endAddress: string;
}

export interface WorkbookDefinedNameStructuredRefValueSnapshot {
  kind: "structured-ref";
  tableName: string;
  columnName: string;
}

export interface WorkbookDefinedNameFormulaValueSnapshot {
  kind: "formula";
  formula: string;
}

export type WorkbookDefinedNameValueSnapshot =
  | LiteralInput
  | WorkbookDefinedNameScalarValueSnapshot
  | WorkbookDefinedNameCellRefValueSnapshot
  | WorkbookDefinedNameRangeRefValueSnapshot
  | WorkbookDefinedNameStructuredRefValueSnapshot
  | WorkbookDefinedNameFormulaValueSnapshot;

export interface WorkbookPropertySnapshot {
  key: string;
  value: LiteralInput;
}

export interface WorkbookSpillSnapshot {
  sheetName: string;
  address: string;
  rows: number;
  cols: number;
}

export type PivotAggregation = "sum" | "count";

export interface WorkbookPivotValueSnapshot {
  sourceColumn: string;
  summarizeBy: PivotAggregation;
  outputLabel?: string;
}

export interface WorkbookPivotSnapshot {
  name: string;
  sheetName: string;
  address: string;
  source: CellRangeRef;
  groupBy: string[];
  values: WorkbookPivotValueSnapshot[];
  rows: number;
  cols: number;
}

export interface WorkbookTableSnapshot {
  name: string;
  sheetName: string;
  startAddress: string;
  endAddress: string;
  columnNames: string[];
  headerRow: boolean;
  totalsRow: boolean;
}

export interface WorkbookAxisMetadataSnapshot {
  start: number;
  count: number;
  size?: number | null;
  hidden?: boolean | null;
}

export interface WorkbookAxisEntrySnapshot {
  id: string;
  index: number;
  size?: number | null;
  hidden?: boolean | null;
}

export type WorkbookCalculationMode = "automatic" | "manual";

export interface WorkbookCalculationSettingsSnapshot {
  mode: WorkbookCalculationMode;
  compatibilityMode?: CompatibilityMode;
}

export interface WorkbookVolatileContextSnapshot {
  recalcEpoch: number;
}

export interface WorkbookFreezePaneSnapshot {
  rows: number;
  cols: number;
}

export interface WorkbookSortKeySnapshot {
  keyAddress: string;
  direction: "asc" | "desc";
}

export interface WorkbookSortSnapshot {
  range: CellRangeRef;
  keys: WorkbookSortKeySnapshot[];
}

export interface WorkbookMetadataSnapshot {
  properties?: WorkbookPropertySnapshot[];
  definedNames?: WorkbookDefinedNameSnapshot[];
  tables?: WorkbookTableSnapshot[];
  spills?: WorkbookSpillSnapshot[];
  pivots?: WorkbookPivotSnapshot[];
  calculationSettings?: WorkbookCalculationSettingsSnapshot;
  volatileContext?: WorkbookVolatileContextSnapshot;
}

export interface SheetMetadataSnapshot {
  rows?: WorkbookAxisEntrySnapshot[];
  columns?: WorkbookAxisEntrySnapshot[];
  rowMetadata?: WorkbookAxisMetadataSnapshot[];
  columnMetadata?: WorkbookAxisMetadataSnapshot[];
  freezePane?: WorkbookFreezePaneSnapshot;
  filters?: CellRangeRef[];
  sorts?: WorkbookSortSnapshot[];
}

export interface WorkbookSnapshot {
  version: 1;
  workbook: {
    name: string;
    metadata?: WorkbookMetadataSnapshot;
  };
  sheets: Array<{
    name: string;
    order: number;
    metadata?: SheetMetadataSnapshot;
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
  depsPtr: number;
  depsLen: number;
  programOffset: number;
  programLength: number;
  constNumberOffset: number;
  constNumberLength: number;
  rangeListOffset: number;
  rangeListLength: number;
  maxStackDepth: number;
}
