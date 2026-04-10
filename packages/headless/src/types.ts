import type { CellValue, LiteralInput, RecalcMetrics } from "@bilig/protocol";
import type { EvaluationResult } from "@bilig/formula";

export type RawCellContent = LiteralInput | string;

export type HeadlessSheet = readonly (readonly RawCellContent[])[];
export type HeadlessSheets = Record<string, HeadlessSheet>;

export interface HeadlessCellAddress {
  sheet: number;
  col: number;
  row: number;
}

export interface HeadlessCellRange {
  start: HeadlessCellAddress;
  end: HeadlessCellAddress;
}

export interface HeadlessAddressFormatOptions {
  includeSheetName?: boolean;
}

export type HeadlessAddressLike = HeadlessCellAddress | HeadlessCellRange;
export type HeadlessAxisInterval = readonly [start: number, count?: number];
export type HeadlessAxisSwapMapping = readonly [from: number, to: number];

export interface HeadlessSheetDimensions {
  width: number;
  height: number;
}

export type HeadlessChange = HeadlessCellChange | HeadlessNamedExpressionChange;

export interface HeadlessCellChange {
  kind: "cell";
  address: HeadlessCellAddress;
  sheetName: string;
  a1: string;
  newValue: CellValue;
}

export interface HeadlessNamedExpressionChange {
  kind: "named-expression";
  name: string;
  scope?: number;
  newValue: CellValue | CellValue[][];
}

export interface HeadlessNamedExpression {
  name: string;
  expression: RawCellContent;
  scope?: number;
  options?: Record<string, string | number | boolean>;
}

export interface SerializedHeadlessNamedExpression extends HeadlessNamedExpression {}

export type HeadlessFunctionArgumentType =
  | "STRING"
  | "NUMBER"
  | "BOOLEAN"
  | "SCALAR"
  | "NOERROR"
  | "RANGE"
  | "INTEGER"
  | "COMPLEX"
  | "ANY";

export interface HeadlessFunctionArgument {
  argumentType: HeadlessFunctionArgumentType;
  passSubtype?: boolean;
  defaultValue?: unknown;
  optionalArg?: boolean;
  minValue?: number;
  maxValue?: number;
  lessThan?: number;
  greaterThan?: number;
}

export interface HeadlessFunctionMetadata {
  method: string;
  parameters?: HeadlessFunctionArgument[];
  repeatLastArgs?: number;
  expandRanges?: boolean;
  returnNumberType?: string;
  sizeOfResultArrayMethod?: string;
  isVolatile?: boolean;
  isDependentOnSheetStructureChange?: boolean;
  doesNotNeedArgumentsToBeComputed?: boolean;
  enableArrayArithmeticForArguments?: boolean;
  vectorizationForbidden?: boolean;
}

export interface HeadlessFunctionPlugin {
  implementedFunctions: Record<string, HeadlessFunctionMetadata>;
  aliases?: Record<string, string>;
}

export interface HeadlessFunctionPluginDefinition extends HeadlessFunctionPlugin {
  id: string;
  functions?: Record<string, (...args: CellValue[]) => EvaluationResult | CellValue>;
}

export type HeadlessFunctionTranslationsPackage = Record<string, Record<string, string>>;

export interface HeadlessLanguagePackage {
  readonly functions?: Record<string, string>;
  readonly errors?: Record<string, string>;
  readonly ui?: Record<string, string>;
  readonly [key: string]: unknown;
}

export type HeadlessLicenseKeyValidityState = "valid" | "invalid" | "expired" | "missing";

export interface HeadlessConfig {
  accentSensitive?: boolean;
  caseSensitive?: boolean;
  caseFirst?: "upper" | "lower" | "false";
  chooseAddressMappingPolicy?: unknown;
  context?: unknown;
  currencySymbol?: string[];
  dateFormats?: string[];
  functionArgSeparator?: string;
  decimalSeparator?: "." | ",";
  evaluateNullToZero?: boolean;
  functionPlugins?: HeadlessFunctionPluginDefinition[];
  ignorePunctuation?: boolean;
  language?: string;
  ignoreWhiteSpace?: "standard" | "any";
  leapYear1900?: boolean;
  licenseKey?: string;
  localeLang?: string;
  matchWholeCell?: boolean;
  arrayColumnSeparator?: "," | ";";
  arrayRowSeparator?: ";" | "|";
  maxRows?: number;
  maxColumns?: number;
  nullDate?: { year: number; month: number; day: number };
  nullYear?: number;
  parseDateTime?: (input: string) => unknown;
  precisionEpsilon?: number;
  precisionRounding?: number;
  stringifyDateTime?: (value: unknown) => string | undefined;
  stringifyDuration?: (value: unknown) => string | undefined;
  smartRounding?: boolean;
  thousandSeparator?: "" | "," | ".";
  timeFormats?: string[];
  useArrayArithmetic?: boolean;
  useColumnIndex?: boolean;
  useStats?: boolean;
  undoLimit?: number;
  useRegularExpressions?: boolean;
  useWildcards?: boolean;
}

export interface HeadlessWorkbookDetailedEventMap {
  sheetAdded: { sheetId: number; sheetName: string };
  sheetRemoved: { sheetId: number; sheetName: string; changes: HeadlessChange[] };
  sheetRenamed: { sheetId: number; oldName: string; newName: string };
  namedExpressionAdded: { name: string; scope?: number; changes: HeadlessChange[] };
  namedExpressionRemoved: { name: string; scope?: number; changes: HeadlessChange[] };
  valuesUpdated: { changes: HeadlessChange[] };
  evaluationSuspended: {};
  evaluationResumed: { changes: HeadlessChange[] };
}

export interface HeadlessWorkbookEventMap {
  sheetAdded: [sheetName: string];
  sheetRemoved: [sheetName: string, changes: HeadlessChange[]];
  sheetRenamed: [oldName: string, newName: string];
  namedExpressionAdded: [name: string, changes: HeadlessChange[]];
  namedExpressionRemoved: [name: string, changes: HeadlessChange[]];
  valuesUpdated: [changes: HeadlessChange[]];
  evaluationSuspended: [];
  evaluationResumed: [changes: HeadlessChange[]];
}

export type HeadlessWorkbookEventName = keyof HeadlessWorkbookEventMap;

export type HeadlessWorkbookListener<EventName extends HeadlessWorkbookEventName> = (
  ...args: HeadlessWorkbookEventMap[EventName]
) => void;

export type HeadlessWorkbookDetailedListener<EventName extends HeadlessWorkbookEventName> = (
  payload: HeadlessWorkbookDetailedEventMap[EventName],
) => void;

export type HeadlessCellType = "EMPTY" | "VALUE" | "FORMULA" | "ARRAY";
export type HeadlessCellValueType = "EMPTY" | "NUMBER" | "STRING" | "BOOLEAN" | "ERROR";
export type HeadlessCellValueDetailedType = HeadlessCellValueType | "DATE" | "TIME" | "DATETIME";

export type HeadlessDependencyRef =
  | { kind: "cell"; address: HeadlessCellAddress }
  | { kind: "range"; range: HeadlessCellRange }
  | { kind: "name"; name: string };

export interface HeadlessDateTime {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
  seconds: number;
}

export interface HeadlessStats {
  batchDepth: number;
  evaluationSuspended: boolean;
  lastMetrics: RecalcMetrics;
}

export interface HeadlessGraphAdapter {
  getDependents(reference: HeadlessAddressLike): HeadlessDependencyRef[];
  getPrecedents(reference: HeadlessAddressLike): HeadlessDependencyRef[];
}

export interface HeadlessRangeMappingAdapter {
  getValues(range: HeadlessCellRange): CellValue[][];
  getSerialized(range: HeadlessCellRange): RawCellContent[][];
}

export interface HeadlessArrayMappingAdapter {
  isPartOfArray(address: HeadlessCellAddress): boolean;
  getFormula(address: HeadlessCellAddress): string | undefined;
}

export interface HeadlessSheetMappingAdapter {
  getSheetName(sheetId: number): string | undefined;
  getSheetId(name: string): number | undefined;
  getSheetNames(): string[];
  countSheets(): number;
}

export interface HeadlessAddressMappingAdapter {
  has(address: HeadlessCellAddress): boolean;
  getValue(address: HeadlessCellAddress): CellValue;
  getFormula(address: HeadlessCellAddress): string | undefined;
}

export interface HeadlessDependencyGraphAdapter {
  getCellDependents(reference: HeadlessAddressLike): HeadlessDependencyRef[];
  getCellPrecedents(reference: HeadlessAddressLike): HeadlessDependencyRef[];
}

export interface HeadlessEvaluatorAdapter {
  recalculate(): HeadlessChange[];
  calculateFormula(formula: string, scope?: number): CellValue | CellValue[][];
}

export interface HeadlessColumnSearchAdapter {
  find(
    sheetId: number,
    column: number,
    matcher: string | ((value: CellValue) => boolean),
  ): HeadlessCellAddress[];
}

export interface HeadlessLazilyTransformingAstServiceAdapter {
  normalizeFormula(formula: string): string;
  validateFormula(formula: string): boolean;
  getNamedExpressionsFromFormula(formula: string): string[];
}

export interface HeadlessWorkbookInternals {
  graph: HeadlessGraphAdapter;
  rangeMapping: HeadlessRangeMappingAdapter;
  arrayMapping: HeadlessArrayMappingAdapter;
  sheetMapping: HeadlessSheetMappingAdapter;
  addressMapping: HeadlessAddressMappingAdapter;
  dependencyGraph: HeadlessDependencyGraphAdapter;
  evaluator: HeadlessEvaluatorAdapter;
  columnSearch: HeadlessColumnSearchAdapter;
  lazilyTransformingAstService: HeadlessLazilyTransformingAstServiceAdapter;
}
