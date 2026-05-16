import { ErrorCode, type FormulaMode, type ValueTag } from './enums.js'
import type { CellNumberFormatRecord, CellStylePatch, CellStyleRecord } from './cell-format-types.js'
import type {
  WorkbookDataModelArtifactsSnapshot,
  WorkbookDocumentPropertiesArtifactsSnapshot,
  WorkbookControlArtifactsSnapshot,
  WorkbookDrawingArtifactsSnapshot,
  WorkbookExternalLinkArtifactsSnapshot,
  WorkbookSheetArrayFormulasSnapshot,
  WorkbookSheetDataTableFormulasSnapshot,
  WorkbookSheetControlArtifactsSnapshot,
  WorkbookSheetDrawingArtifactsSnapshot,
  WorkbookSheetThreadedCommentArtifactsSnapshot,
  WorkbookSlicerConnectionArtifactsSnapshot,
  WorkbookStyleArtifactsSnapshot,
  WorkbookThreadedCommentArtifactsSnapshot,
} from './package-artifacts.js'
import type {
  WorkbookPivotArtifactsSnapshot,
  WorkbookPivotSnapshot,
  WorkbookSheetPivotArtifactsSnapshot,
  WorkbookExternalWorkbookReferenceSnapshot,
  WorkbookUnsupportedFormulaDependencySnapshot,
  WorkbookUnsupportedPivotSnapshot,
} from './workbook-pivot-types.js'
export type {
  CellBorderSidePatch,
  CellBorderSideSnapshot,
  CellBorderStyle,
  CellBorderWeight,
  CellDateStyle,
  CellHorizontalAlignment,
  CellNumberFormatInput,
  CellNumberFormatKind,
  CellNumberFormatPreset,
  CellNumberFormatRecord,
  CellNumberNegativeStyle,
  CellNumberZeroStyle,
  CellStyleAlignmentPatch,
  CellStyleAlignmentSnapshot,
  CellStyleBordersPatch,
  CellStyleBordersSnapshot,
  CellStyleField,
  CellStyleFillPatch,
  CellStyleFillSnapshot,
  CellStyleFontPatch,
  CellStyleFontSnapshot,
  CellStylePatch,
  CellStyleProtectionSnapshot,
  CellStyleRecord,
  CellVerticalAlignment,
} from './cell-format-types.js'
export {
  CELL_BORDER_STYLE_VALUES,
  CELL_BORDER_WEIGHT_VALUES,
  CELL_DATE_STYLE_VALUES,
  CELL_HORIZONTAL_ALIGNMENT_VALUES,
  CELL_NUMBER_FORMAT_KIND_VALUES,
  CELL_NUMBER_NEGATIVE_STYLE_VALUES,
  CELL_NUMBER_ZERO_STYLE_VALUES,
  CELL_STYLE_FIELD_VALUES,
  CELL_VERTICAL_ALIGNMENT_VALUES,
} from './cell-format-types.js'
export type {
  PivotAggregation,
  WorkbookExternalWorkbookReferenceSnapshot,
  WorkbookPivotArtifactsSnapshot,
  WorkbookPivotPackagePartSnapshot,
  WorkbookPivotSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookSheetPivotArtifactsSnapshot,
  WorkbookUnsupportedFormulaDependencySnapshot,
  WorkbookUnsupportedPivotSnapshot,
} from './workbook-pivot-types.js'

export type CellIndex = number
export type FormulaId = number
export type RangeIndex = number
export type EntityId = number
export type LiteralInput = number | string | boolean | null
export type CompatibilityMode = 'excel-modern' | 'odf-1.4'

export type EmptyValue = { tag: ValueTag.Empty }
export type NumberValue = { tag: ValueTag.Number; value: number }
export type BooleanValue = { tag: ValueTag.Boolean; value: boolean }
export type StringValue = { tag: ValueTag.String; value: string; stringId: number }
export type ErrorValue = { tag: ValueTag.Error; code: ErrorCode }

export type CellValue = EmptyValue | NumberValue | BooleanValue | StringValue | ErrorValue

export function formatErrorCode(code: ErrorCode): string {
  switch (code) {
    case ErrorCode.None:
      return '#ERROR!'
    case ErrorCode.Div0:
      return '#DIV/0!'
    case ErrorCode.Ref:
      return '#REF!'
    case ErrorCode.Value:
      return '#VALUE!'
    case ErrorCode.Name:
      return '#NAME?'
    case ErrorCode.NA:
      return '#N/A'
    case ErrorCode.Cycle:
      return '#CYCLE!'
    case ErrorCode.Spill:
      return '#SPILL!'
    case ErrorCode.Blocked:
      return '#BLOCKED!'
    default:
      return '#ERROR!'
  }
}

export interface CellSnapshot {
  sheetName: string
  address: string
  formula?: string
  format?: string
  numberFormatId?: string
  styleId?: string
  input?: LiteralInput
  value: CellValue
  flags: number
  version: number
}

export interface DependencySnapshot {
  directPrecedents: string[]
  directDependents: string[]
}

export interface ExplainCellSnapshot {
  sheetName: string
  address: string
  formula?: string
  format?: string
  numberFormatId?: string
  styleId?: string
  mode?: FormulaMode
  value: CellValue
  flags: number
  version: number
  topoRank?: number
  inCycle: boolean
  directPrecedents: string[]
  directDependents: string[]
}

export interface RecalcMetrics {
  batchId: number
  changedInputCount: number
  dirtyFormulaCount: number
  wasmFormulaCount: number
  jsFormulaCount: number
  rangeNodeVisits: number
  recalcMs: number
  compileMs: number
}

export interface AxisInvalidation {
  sheetName: string
  startIndex: number
  endIndex: number
}

export interface EngineChangedCell {
  kind: 'cell'
  cellIndex: number
  address: {
    sheet: number
    row: number
    col: number
  }
  sheetName: string
  a1: string
  newValue: CellValue
}

export interface EngineEvent {
  kind: 'batch'
  invalidation: 'cells' | 'full'
  changedCellIndices: Uint32Array | number[]
  changedCells: readonly EngineChangedCell[]
  invalidatedRanges: CellRangeRef[]
  invalidatedRows: AxisInvalidation[]
  invalidatedColumns: AxisInvalidation[]
  metrics: RecalcMetrics
}

export interface CellRangeRef {
  sheetName: string
  startAddress: string
  endAddress: string
}

export type WorkbookAutoFilterCustomOperator = 'equal' | 'lessThan' | 'lessThanOrEqual' | 'notEqual' | 'greaterThanOrEqual' | 'greaterThan'

export interface WorkbookAutoFilterValueCriteriaSnapshot {
  blank?: boolean
  values: string[]
}

export interface WorkbookAutoFilterCustomCriterionSnapshot {
  operator?: WorkbookAutoFilterCustomOperator
  value: string
}

export interface WorkbookAutoFilterCustomCriteriaSnapshot {
  and?: boolean
  filters: WorkbookAutoFilterCustomCriterionSnapshot[]
}

export interface WorkbookAutoFilterColumnSnapshot {
  colId: number
  hiddenButton?: boolean
  showButton?: boolean
  filters?: WorkbookAutoFilterValueCriteriaSnapshot
  customFilters?: WorkbookAutoFilterCustomCriteriaSnapshot
}

export interface WorkbookAutoFilterSnapshot extends CellRangeRef {
  criteria?: WorkbookAutoFilterColumnSnapshot[]
}

export interface SelectionRange {
  startAddress: string
  endAddress: string
}

export type SelectionEditMode = 'idle' | 'cell' | 'formula'
export type SyncState = 'local-only' | 'syncing' | 'live' | 'behind' | 'reconnecting'

export interface SelectionState {
  sheetName: string
  address: string | null
  anchorAddress: string | null
  range: SelectionRange | null
  editMode: SelectionEditMode
}

export interface Viewport {
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
}

export interface WorkbookDefinedNameSnapshot {
  name: string
  scopeSheetName?: string
  value: WorkbookDefinedNameValueSnapshot
}

export interface WorkbookDefinedNameScalarValueSnapshot {
  kind: 'scalar'
  value: LiteralInput
}

export interface WorkbookDefinedNameCellRefValueSnapshot {
  kind: 'cell-ref'
  sheetName: string
  address: string
}

export interface WorkbookDefinedNameRangeRefValueSnapshot {
  kind: 'range-ref'
  sheetName: string
  startAddress: string
  endAddress: string
}

export interface WorkbookDefinedNameStructuredRefValueSnapshot {
  kind: 'structured-ref'
  tableName: string
  columnName: string
}

export interface WorkbookDefinedNameFormulaValueSnapshot {
  kind: 'formula'
  formula: string
}

export type WorkbookDefinedNameValueSnapshot =
  | LiteralInput
  | WorkbookDefinedNameScalarValueSnapshot
  | WorkbookDefinedNameCellRefValueSnapshot
  | WorkbookDefinedNameRangeRefValueSnapshot
  | WorkbookDefinedNameStructuredRefValueSnapshot
  | WorkbookDefinedNameFormulaValueSnapshot

export interface WorkbookPropertySnapshot {
  key: string
  value: LiteralInput
}

export interface WorkbookProtectionXmlAttributeSnapshot {
  name: string
  value: string
}

export interface WorkbookProtectionSnapshot {
  lockStructure?: boolean
  lockWindows?: boolean
  lockRevision?: boolean
  xmlAttributes?: WorkbookProtectionXmlAttributeSnapshot[]
}

export interface WorkbookSpillSnapshot {
  sheetName: string
  address: string
  rows: number
  cols: number
}

export interface WorkbookSheetCellStyleIndexSnapshot {
  address: string
  styleIndex: number
}

export interface WorkbookSheetStyleArtifactsSnapshot {
  cellStyleIndexes: WorkbookSheetCellStyleIndexSnapshot[]
  blankCellAddresses?: string[]
}

export type WorkbookChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter'
export type WorkbookChartSeriesOrientation = 'rows' | 'columns'
export type WorkbookChartLegendPosition = 'top' | 'right' | 'bottom' | 'left' | 'hidden'

export interface WorkbookChartSnapshot {
  id: string
  sheetName: string
  address: string
  source: CellRangeRef
  chartType: WorkbookChartType
  seriesOrientation?: WorkbookChartSeriesOrientation
  firstRowAsHeaders?: boolean
  firstColumnAsLabels?: boolean
  title?: string
  legendPosition?: WorkbookChartLegendPosition
  rows: number
  cols: number
}

export interface WorkbookImageSnapshot {
  id: string
  sheetName: string
  address: string
  sourceUrl: string
  rows: number
  cols: number
  altText?: string
}

export type WorkbookShapeType = 'rectangle' | 'roundedRectangle' | 'ellipse' | 'line' | 'arrow' | 'textBox'

export interface WorkbookShapeSnapshot {
  id: string
  sheetName: string
  address: string
  shapeType: WorkbookShapeType
  rows: number
  cols: number
  text?: string
  fillColor?: string
  strokeColor?: string
}

export interface WorkbookTableSnapshot {
  name: string
  sheetName: string
  startAddress: string
  endAddress: string
  columnNames: string[]
  columns?: WorkbookTableColumnSnapshot[]
  headerRow: boolean
  totalsRow: boolean
  style?: WorkbookTableStyleSnapshot
  sortState?: string
}

export interface WorkbookTableColumnSnapshot {
  name: string
  totalsRowLabel?: string
  totalsRowFunction?: string
}

export interface WorkbookTableStyleSnapshot {
  name?: string
  showFirstColumn?: boolean
  showLastColumn?: boolean
  showRowStripes?: boolean
  showColumnStripes?: boolean
}

export interface WorkbookAxisMetadataSnapshot {
  start: number
  count: number
  size?: number | null
  hidden?: boolean | null
  styleIndex?: number | null
  xlsxWidth?: number | null
  xlsxHeight?: number | null
  customFormat?: boolean | null
  customWidth?: boolean | null
  bestFit?: boolean | null
  outlineLevel?: number | null
  collapsed?: boolean | null
  customHeight?: boolean | null
  thickTop?: boolean | null
  thickBottom?: boolean | null
}

export interface WorkbookAxisEntrySnapshot {
  id: string
  index: number
  size?: number | null
  hidden?: boolean | null
  styleIndex?: number | null
  xlsxWidth?: number | null
  xlsxHeight?: number | null
  customFormat?: boolean | null
  customWidth?: boolean | null
  bestFit?: boolean | null
  outlineLevel?: number | null
  collapsed?: boolean | null
  customHeight?: boolean | null
  thickTop?: boolean | null
  thickBottom?: boolean | null
}

export interface WorkbookSheetFormatPrSnapshot {
  baseColWidth?: number | null
  defaultColWidth?: number | null
  defaultRowHeight?: number | null
  customHeight?: boolean | null
  outlineLevelRow?: number | null
  outlineLevelCol?: number | null
  thickTop?: boolean | null
  thickBottom?: boolean | null
}

export type WorkbookCalculationMode = 'automatic' | 'manual'
export type WorkbookDateSystem = '1900' | '1904'

export interface WorkbookCalculationSettingsSnapshot {
  mode: WorkbookCalculationMode
  compatibilityMode?: CompatibilityMode
  dateSystem?: WorkbookDateSystem
  iterate?: boolean | null
  iterateCount?: number | null
  iterateDelta?: string | null
  fullPrecision?: boolean | null
  fullCalcOnLoad?: boolean | null
  concurrentCalc?: boolean | null
}

export interface WorkbookVolatileContextSnapshot {
  recalcEpoch: number
}

export interface WorkbookMacroSheetCodeNameSnapshot {
  sheetName: string
  codeName: string
}

export interface WorkbookMacroPayloadSnapshot {
  kind: 'vbaProject'
  storage: 'base64'
  dataBase64: string
  byteLength: number
  preservedWithoutExecution: true
  workbookCodeName?: string
  sheetCodeNames?: WorkbookMacroSheetCodeNameSnapshot[]
}

export type WorkbookFreezePaneActivePane = 'bottomRight' | 'bottomLeft' | 'topRight' | 'topLeft'

export interface WorkbookFreezePaneSnapshot {
  rows: number
  cols: number
  topLeftCell?: string
  activePane?: WorkbookFreezePaneActivePane
}

export interface WorkbookSheetTabColorSnapshot {
  rgb?: string
  theme?: string
  tint?: string
  indexed?: string
  auto?: string
}

export interface WorkbookMergeRangeSnapshot extends CellRangeRef {}

export interface SheetStyleRangeSnapshot {
  range: CellRangeRef
  styleId: string
}

export interface SheetFormatRangeSnapshot {
  range: CellRangeRef
  formatId: string
}

export interface WorkbookSortKeySnapshot {
  keyAddress: string
  direction: 'asc' | 'desc'
}

export interface WorkbookSortSnapshot {
  range: CellRangeRef
  keys: WorkbookSortKeySnapshot[]
}

export type WorkbookValidationComparisonOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual'

export type WorkbookValidationErrorStyle = 'stop' | 'warning' | 'information'

export interface WorkbookValidationNamedRangeSourceSnapshot {
  kind: 'named-range'
  name: string
}

export interface WorkbookValidationCellRefSourceSnapshot {
  kind: 'cell-ref'
  sheetName: string
  address: string
}

export interface WorkbookValidationRangeRefSourceSnapshot {
  kind: 'range-ref'
  sheetName: string
  startAddress: string
  endAddress: string
}

export interface WorkbookValidationStructuredRefSourceSnapshot {
  kind: 'structured-ref'
  tableName: string
  columnName: string
}

export type WorkbookValidationListSourceSnapshot =
  | WorkbookValidationNamedRangeSourceSnapshot
  | WorkbookValidationCellRefSourceSnapshot
  | WorkbookValidationRangeRefSourceSnapshot
  | WorkbookValidationStructuredRefSourceSnapshot

export interface WorkbookListValidationRuleSnapshot {
  kind: 'list'
  values?: LiteralInput[]
  source?: WorkbookValidationListSourceSnapshot
}

export interface WorkbookCheckboxValidationRuleSnapshot {
  kind: 'checkbox'
  checkedValue?: LiteralInput
  uncheckedValue?: LiteralInput
}

export interface WorkbookAnyValidationRuleSnapshot {
  kind: 'any'
}

export interface WorkbookScalarValidationRuleSnapshot {
  kind: 'whole' | 'decimal' | 'date' | 'time' | 'textLength'
  operator: WorkbookValidationComparisonOperator
  values: LiteralInput[]
}

export type WorkbookDataValidationRuleSnapshot =
  | WorkbookListValidationRuleSnapshot
  | WorkbookCheckboxValidationRuleSnapshot
  | WorkbookAnyValidationRuleSnapshot
  | WorkbookScalarValidationRuleSnapshot

export interface WorkbookDataValidationSnapshot {
  range: CellRangeRef
  rule: WorkbookDataValidationRuleSnapshot
  allowBlank?: boolean
  showDropdown?: boolean
  promptTitle?: string
  promptMessage?: string
  errorStyle?: WorkbookValidationErrorStyle
  errorTitle?: string
  errorMessage?: string
}

export interface WorkbookConditionalFormatCellIsRuleSnapshot {
  kind: 'cellIs'
  operator: WorkbookValidationComparisonOperator
  values: LiteralInput[]
}

export interface WorkbookConditionalFormatTextContainsRuleSnapshot {
  kind: 'textContains'
  text: string
  caseSensitive?: boolean
}

export interface WorkbookConditionalFormatFormulaRuleSnapshot {
  kind: 'formula'
  formula: string
}

export interface WorkbookConditionalFormatBlanksRuleSnapshot {
  kind: 'blanks'
}

export interface WorkbookConditionalFormatNotBlanksRuleSnapshot {
  kind: 'notBlanks'
}

export type WorkbookConditionalFormatRuleSnapshot =
  | WorkbookConditionalFormatCellIsRuleSnapshot
  | WorkbookConditionalFormatTextContainsRuleSnapshot
  | WorkbookConditionalFormatFormulaRuleSnapshot
  | WorkbookConditionalFormatBlanksRuleSnapshot
  | WorkbookConditionalFormatNotBlanksRuleSnapshot

export interface WorkbookConditionalFormatSnapshot {
  id: string
  range: CellRangeRef
  rule: WorkbookConditionalFormatRuleSnapshot
  style: CellStylePatch
  stopIfTrue?: boolean
  priority?: number
}

export interface WorkbookSheetProtectionXmlAttributeSnapshot {
  name: string
  value: string
}

export interface WorkbookSheetProtectionSnapshot {
  sheetName: string
  hideFormulas?: boolean
  xmlAttributes?: WorkbookSheetProtectionXmlAttributeSnapshot[]
}

export interface WorkbookRangeProtectionSnapshot {
  id: string
  range: CellRangeRef
  hideFormulas?: boolean
}

export interface WorkbookCommentEntrySnapshot {
  id: string
  body: string
  authorUserId?: string
  authorDisplayName?: string
  createdAtUnixMs?: number
}

export interface WorkbookCommentThreadSnapshot {
  threadId: string
  sheetName: string
  address: string
  comments: WorkbookCommentEntrySnapshot[]
  resolved?: boolean
  resolvedByUserId?: string
  resolvedAtUnixMs?: number
}

export interface WorkbookNoteSnapshot {
  sheetName: string
  address: string
  text: string
}

export interface WorkbookHyperlinkSnapshot {
  sheetName: string
  address: string
  target: string
  tooltip?: string
  display?: string
}

export interface WorkbookLegacyCommentVmlSnapshot {
  relationshipTarget: string
  vmlXml: string
  commentsRelationshipTarget?: string
  commentsXml?: string
  commentSignature: string
}

export interface WorkbookPrinterSettingsSnapshot {
  relationshipTarget: string
  storage: 'base64'
  dataBase64: string
  byteLength: number
  pageSetupXml?: string
}

export interface WorkbookSheetPrSnapshot {
  xml: string
}

export interface WorkbookIgnoredErrorsSnapshot {
  xml: string
}

export interface WorkbookSparklinesSnapshot {
  xml: string
}

export interface WorkbookSheetConditionalFormatArtifactsSnapshot {
  xml: string
}

export interface WorkbookChartSheetArtifactsSnapshot {
  name: string
  relationshipTarget: string
  sheetId?: number
  state?: 'hidden' | 'veryHidden'
}

export type WorkbookSheetVisibilitySnapshot = 'hidden' | 'veryHidden'

export interface WorkbookCellMetadataSnapshot {
  relationshipTarget: string
  metadataXml: string
}

export interface WorkbookCellMetadataReferenceSnapshot {
  address: string
  cellSignature: string
  cm?: string
  vm?: string
}

export interface WorkbookRichTextCellSnapshot {
  address: string
  text: string
  storage: 'sharedString' | 'inlineString'
  xml: string
}

export interface WorkbookSheetRichTextArtifactsSnapshot {
  cells: WorkbookRichTextCellSnapshot[]
}

export interface WorkbookViewStateSnapshot {
  bookViewsXml: string
}

export interface WorkbookSheetViewStateSnapshot {
  sheetViewsXml: string
}

export interface WorkbookSheetPrintPageSetupSnapshot {
  printOptionsXml?: string
  pageMarginsXml?: string
  pageSetupXml?: string
  headerFooterXml?: string
  rowBreaksXml?: string
  colBreaksXml?: string
}

export interface WorkbookMetadataSnapshot {
  properties?: WorkbookPropertySnapshot[]
  documentPropertyArtifacts?: WorkbookDocumentPropertiesArtifactsSnapshot
  workbookProtection?: WorkbookProtectionSnapshot
  definedNames?: WorkbookDefinedNameSnapshot[]
  tables?: WorkbookTableSnapshot[]
  spills?: WorkbookSpillSnapshot[]
  pivots?: WorkbookPivotSnapshot[]
  externalWorkbookReferences?: WorkbookExternalWorkbookReferenceSnapshot[]
  unsupportedFormulaDependencies?: WorkbookUnsupportedFormulaDependencySnapshot[]
  unsupportedPivots?: WorkbookUnsupportedPivotSnapshot[]
  pivotArtifacts?: WorkbookPivotArtifactsSnapshot
  drawingArtifacts?: WorkbookDrawingArtifactsSnapshot
  chartArtifacts?: WorkbookDrawingArtifactsSnapshot
  chartSheetArtifacts?: WorkbookChartSheetArtifactsSnapshot[]
  controlArtifacts?: WorkbookControlArtifactsSnapshot
  dataModelArtifacts?: WorkbookDataModelArtifactsSnapshot
  externalLinkArtifacts?: WorkbookExternalLinkArtifactsSnapshot
  slicerConnectionArtifacts?: WorkbookSlicerConnectionArtifactsSnapshot
  threadedCommentArtifacts?: WorkbookThreadedCommentArtifactsSnapshot
  viewState?: WorkbookViewStateSnapshot
  charts?: WorkbookChartSnapshot[]
  images?: WorkbookImageSnapshot[]
  shapes?: WorkbookShapeSnapshot[]
  styles?: CellStyleRecord[]
  styleArtifacts?: WorkbookStyleArtifactsSnapshot
  formats?: CellNumberFormatRecord[]
  macroPayloads?: WorkbookMacroPayloadSnapshot[]
  calculationSettings?: WorkbookCalculationSettingsSnapshot
  volatileContext?: WorkbookVolatileContextSnapshot
  cellMetadata?: WorkbookCellMetadataSnapshot
}

export interface SheetMetadataSnapshot {
  rows?: WorkbookAxisEntrySnapshot[]
  columns?: WorkbookAxisEntrySnapshot[]
  rowMetadata?: WorkbookAxisMetadataSnapshot[]
  columnMetadata?: WorkbookAxisMetadataSnapshot[]
  sheetFormatPr?: WorkbookSheetFormatPrSnapshot
  styleRanges?: SheetStyleRangeSnapshot[]
  formatRanges?: SheetFormatRangeSnapshot[]
  freezePane?: WorkbookFreezePaneSnapshot
  tabColor?: WorkbookSheetTabColorSnapshot
  merges?: WorkbookMergeRangeSnapshot[]
  sheetProtection?: WorkbookSheetProtectionSnapshot
  filters?: WorkbookAutoFilterSnapshot[]
  sorts?: WorkbookSortSnapshot[]
  validations?: WorkbookDataValidationSnapshot[]
  conditionalFormats?: WorkbookConditionalFormatSnapshot[]
  conditionalFormatArtifacts?: WorkbookSheetConditionalFormatArtifactsSnapshot
  protectedRanges?: WorkbookRangeProtectionSnapshot[]
  commentThreads?: WorkbookCommentThreadSnapshot[]
  notes?: WorkbookNoteSnapshot[]
  hyperlinks?: WorkbookHyperlinkSnapshot[]
  drawingArtifacts?: WorkbookSheetDrawingArtifactsSnapshot
  controlArtifacts?: WorkbookSheetControlArtifactsSnapshot
  arrayFormulas?: WorkbookSheetArrayFormulasSnapshot
  dataTableFormulas?: WorkbookSheetDataTableFormulasSnapshot
  legacyCommentVml?: WorkbookLegacyCommentVmlSnapshot
  printerSettings?: WorkbookPrinterSettingsSnapshot[]
  sheetPr?: WorkbookSheetPrSnapshot
  ignoredErrors?: WorkbookIgnoredErrorsSnapshot
  sparklines?: WorkbookSparklinesSnapshot
  styleArtifacts?: WorkbookSheetStyleArtifactsSnapshot
  pivotArtifacts?: WorkbookSheetPivotArtifactsSnapshot
  visibility?: WorkbookSheetVisibilitySnapshot
  cellMetadataRefs?: WorkbookCellMetadataReferenceSnapshot[]
  richTextArtifacts?: WorkbookSheetRichTextArtifactsSnapshot
  threadedCommentArtifacts?: WorkbookSheetThreadedCommentArtifactsSnapshot
  viewState?: WorkbookSheetViewStateSnapshot
  printPageSetup?: WorkbookSheetPrintPageSetupSnapshot
}

export interface WorkbookSnapshot {
  version: 1
  workbook: {
    name: string
    metadata?: WorkbookMetadataSnapshot
  }
  sheets: Array<{
    id?: number
    name: string
    order: number
    metadata?: SheetMetadataSnapshot
    cells: Array<{
      address: string
      row?: number
      col?: number
      value?: LiteralInput
      formula?: string
      format?: string
    }>
  }>
}

export interface FormulaRecord {
  id: FormulaId
  source: string
  mode: FormulaMode
  depsPtr: number
  depsLen: number
  programOffset: number
  programLength: number
  constNumberOffset: number
  constNumberLength: number
  rangeListOffset: number
  rangeListLength: number
  maxStackDepth: number
}
