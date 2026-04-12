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
  numberFormatId?: string;
  styleId?: string;
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
  numberFormatId?: string;
  styleId?: string;
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

export interface AxisInvalidation {
  sheetName: string;
  startIndex: number;
  endIndex: number;
}

export interface EngineEvent {
  kind: "batch";
  invalidation: "cells" | "full";
  changedCellIndices: Uint32Array | number[];
  invalidatedRanges: CellRangeRef[];
  invalidatedRows: AxisInvalidation[];
  invalidatedColumns: AxisInvalidation[];
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

export type WorkbookChartType = "column" | "bar" | "line" | "area" | "pie" | "scatter";
export type WorkbookChartSeriesOrientation = "rows" | "columns";
export type WorkbookChartLegendPosition = "top" | "right" | "bottom" | "left" | "hidden";

export interface WorkbookChartSnapshot {
  id: string;
  sheetName: string;
  address: string;
  source: CellRangeRef;
  chartType: WorkbookChartType;
  seriesOrientation?: WorkbookChartSeriesOrientation;
  firstRowAsHeaders?: boolean;
  firstColumnAsLabels?: boolean;
  title?: string;
  legendPosition?: WorkbookChartLegendPosition;
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

export interface CellStyleFillSnapshot {
  backgroundColor: string;
}

export interface CellStyleFontSnapshot {
  family?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
}

export type CellHorizontalAlignment = "general" | "left" | "center" | "right";
export type CellVerticalAlignment = "top" | "middle" | "bottom";
export type CellBorderStyle = "solid" | "dashed" | "dotted" | "double";
export type CellBorderWeight = "thin" | "medium" | "thick";

export interface CellStyleAlignmentSnapshot {
  horizontal?: CellHorizontalAlignment;
  vertical?: CellVerticalAlignment;
  wrap?: boolean;
  indent?: number;
}

export interface CellBorderSideSnapshot {
  style: CellBorderStyle;
  weight: CellBorderWeight;
  color: string;
}

export interface CellStyleBordersSnapshot {
  top?: CellBorderSideSnapshot;
  right?: CellBorderSideSnapshot;
  bottom?: CellBorderSideSnapshot;
  left?: CellBorderSideSnapshot;
}

export interface CellStyleRecord {
  id: string;
  fill?: CellStyleFillSnapshot;
  font?: CellStyleFontSnapshot;
  alignment?: CellStyleAlignmentSnapshot;
  borders?: CellStyleBordersSnapshot;
}

export interface CellStyleFillPatch {
  backgroundColor?: string | null;
}

export interface CellStyleFontPatch {
  family?: string | null;
  size?: number | null;
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | null;
  color?: string | null;
}

export interface CellStyleAlignmentPatch {
  horizontal?: CellHorizontalAlignment | null;
  vertical?: CellVerticalAlignment | null;
  wrap?: boolean | null;
  indent?: number | null;
}

export interface CellBorderSidePatch {
  style?: CellBorderStyle | null;
  weight?: CellBorderWeight | null;
  color?: string | null;
}

export interface CellStyleBordersPatch {
  top?: CellBorderSidePatch | null;
  right?: CellBorderSidePatch | null;
  bottom?: CellBorderSidePatch | null;
  left?: CellBorderSidePatch | null;
}

export interface CellStylePatch {
  fill?: CellStyleFillPatch | null;
  font?: CellStyleFontPatch | null;
  alignment?: CellStyleAlignmentPatch | null;
  borders?: CellStyleBordersPatch | null;
}

export type CellStyleField =
  | "backgroundColor"
  | "fontFamily"
  | "fontSize"
  | "fontBold"
  | "fontItalic"
  | "fontUnderline"
  | "fontColor"
  | "alignmentHorizontal"
  | "alignmentVertical"
  | "alignmentWrap"
  | "alignmentIndent"
  | "borderTop"
  | "borderRight"
  | "borderBottom"
  | "borderLeft";

export type CellNumberFormatKind =
  | "general"
  | "number"
  | "currency"
  | "accounting"
  | "percent"
  | "date"
  | "time"
  | "datetime"
  | "text";

export type CellNumberNegativeStyle = "minus" | "parentheses";
export type CellNumberZeroStyle = "zero" | "dash";
export type CellDateStyle = "short" | "iso";

export interface CellNumberFormatPreset {
  kind: CellNumberFormatKind;
  currency?: string;
  decimals?: number;
  useGrouping?: boolean;
  negativeStyle?: CellNumberNegativeStyle;
  zeroStyle?: CellNumberZeroStyle;
  dateStyle?: CellDateStyle;
}

export type CellNumberFormatInput = string | CellNumberFormatPreset;

export interface CellNumberFormatRecord {
  id: string;
  code: string;
  kind: CellNumberFormatKind;
}

export interface SheetStyleRangeSnapshot {
  range: CellRangeRef;
  styleId: string;
}

export interface SheetFormatRangeSnapshot {
  range: CellRangeRef;
  formatId: string;
}

export interface WorkbookSortKeySnapshot {
  keyAddress: string;
  direction: "asc" | "desc";
}

export interface WorkbookSortSnapshot {
  range: CellRangeRef;
  keys: WorkbookSortKeySnapshot[];
}

export type WorkbookValidationComparisonOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual";

export type WorkbookValidationErrorStyle = "stop" | "warning" | "information";

export interface WorkbookValidationNamedRangeSourceSnapshot {
  kind: "named-range";
  name: string;
}

export interface WorkbookValidationCellRefSourceSnapshot {
  kind: "cell-ref";
  sheetName: string;
  address: string;
}

export interface WorkbookValidationRangeRefSourceSnapshot {
  kind: "range-ref";
  sheetName: string;
  startAddress: string;
  endAddress: string;
}

export interface WorkbookValidationStructuredRefSourceSnapshot {
  kind: "structured-ref";
  tableName: string;
  columnName: string;
}

export type WorkbookValidationListSourceSnapshot =
  | WorkbookValidationNamedRangeSourceSnapshot
  | WorkbookValidationCellRefSourceSnapshot
  | WorkbookValidationRangeRefSourceSnapshot
  | WorkbookValidationStructuredRefSourceSnapshot;

export interface WorkbookListValidationRuleSnapshot {
  kind: "list";
  values?: LiteralInput[];
  source?: WorkbookValidationListSourceSnapshot;
}

export interface WorkbookCheckboxValidationRuleSnapshot {
  kind: "checkbox";
  checkedValue?: LiteralInput;
  uncheckedValue?: LiteralInput;
}

export interface WorkbookScalarValidationRuleSnapshot {
  kind: "whole" | "decimal" | "date" | "time" | "textLength";
  operator: WorkbookValidationComparisonOperator;
  values: LiteralInput[];
}

export type WorkbookDataValidationRuleSnapshot =
  | WorkbookListValidationRuleSnapshot
  | WorkbookCheckboxValidationRuleSnapshot
  | WorkbookScalarValidationRuleSnapshot;

export interface WorkbookDataValidationSnapshot {
  range: CellRangeRef;
  rule: WorkbookDataValidationRuleSnapshot;
  allowBlank?: boolean;
  showDropdown?: boolean;
  promptTitle?: string;
  promptMessage?: string;
  errorStyle?: WorkbookValidationErrorStyle;
  errorTitle?: string;
  errorMessage?: string;
}

export interface WorkbookConditionalFormatCellIsRuleSnapshot {
  kind: "cellIs";
  operator: WorkbookValidationComparisonOperator;
  values: LiteralInput[];
}

export interface WorkbookConditionalFormatTextContainsRuleSnapshot {
  kind: "textContains";
  text: string;
  caseSensitive?: boolean;
}

export interface WorkbookConditionalFormatFormulaRuleSnapshot {
  kind: "formula";
  formula: string;
}

export interface WorkbookConditionalFormatBlanksRuleSnapshot {
  kind: "blanks";
}

export interface WorkbookConditionalFormatNotBlanksRuleSnapshot {
  kind: "notBlanks";
}

export type WorkbookConditionalFormatRuleSnapshot =
  | WorkbookConditionalFormatCellIsRuleSnapshot
  | WorkbookConditionalFormatTextContainsRuleSnapshot
  | WorkbookConditionalFormatFormulaRuleSnapshot
  | WorkbookConditionalFormatBlanksRuleSnapshot
  | WorkbookConditionalFormatNotBlanksRuleSnapshot;

export interface WorkbookConditionalFormatSnapshot {
  id: string;
  range: CellRangeRef;
  rule: WorkbookConditionalFormatRuleSnapshot;
  style: CellStylePatch;
  stopIfTrue?: boolean;
  priority?: number;
}

export interface WorkbookSheetProtectionSnapshot {
  sheetName: string;
  hideFormulas?: boolean;
}

export interface WorkbookRangeProtectionSnapshot {
  id: string;
  range: CellRangeRef;
  hideFormulas?: boolean;
}

export interface WorkbookCommentEntrySnapshot {
  id: string;
  body: string;
  authorUserId?: string;
  authorDisplayName?: string;
  createdAtUnixMs?: number;
}

export interface WorkbookCommentThreadSnapshot {
  threadId: string;
  sheetName: string;
  address: string;
  comments: WorkbookCommentEntrySnapshot[];
  resolved?: boolean;
  resolvedByUserId?: string;
  resolvedAtUnixMs?: number;
}

export interface WorkbookNoteSnapshot {
  sheetName: string;
  address: string;
  text: string;
}

export interface WorkbookMetadataSnapshot {
  properties?: WorkbookPropertySnapshot[];
  definedNames?: WorkbookDefinedNameSnapshot[];
  tables?: WorkbookTableSnapshot[];
  spills?: WorkbookSpillSnapshot[];
  pivots?: WorkbookPivotSnapshot[];
  charts?: WorkbookChartSnapshot[];
  styles?: CellStyleRecord[];
  formats?: CellNumberFormatRecord[];
  calculationSettings?: WorkbookCalculationSettingsSnapshot;
  volatileContext?: WorkbookVolatileContextSnapshot;
}

export interface SheetMetadataSnapshot {
  rows?: WorkbookAxisEntrySnapshot[];
  columns?: WorkbookAxisEntrySnapshot[];
  rowMetadata?: WorkbookAxisMetadataSnapshot[];
  columnMetadata?: WorkbookAxisMetadataSnapshot[];
  styleRanges?: SheetStyleRangeSnapshot[];
  formatRanges?: SheetFormatRangeSnapshot[];
  freezePane?: WorkbookFreezePaneSnapshot;
  sheetProtection?: WorkbookSheetProtectionSnapshot;
  filters?: CellRangeRef[];
  sorts?: WorkbookSortSnapshot[];
  validations?: WorkbookDataValidationSnapshot[];
  conditionalFormats?: WorkbookConditionalFormatSnapshot[];
  protectedRanges?: WorkbookRangeProtectionSnapshot[];
  commentThreads?: WorkbookCommentThreadSnapshot[];
  notes?: WorkbookNoteSnapshot[];
}

export interface WorkbookSnapshot {
  version: 1;
  workbook: {
    name: string;
    metadata?: WorkbookMetadataSnapshot;
  };
  sheets: Array<{
    id?: number;
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
