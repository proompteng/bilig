import { SpreadsheetEngine } from "@bilig/core";
import type { EngineOp } from "@bilig/workbook-domain";
import {
  ErrorCode,
  type EngineEvent,
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  type CellRangeRef,
  type CellSnapshot,
  type CellValue,
  type LiteralInput,
  type WorkbookDefinedNameValueSnapshot,
} from "@bilig/protocol";
import {
  excelSerialToDateParts,
  formatAddress,
  formatRangeAddress,
  installExternalFunctionAdapter,
  isArrayValue,
  isCellReferenceText,
  parseCellAddress,
  parseFormula,
  parseRangeAddress,
  serializeFormula,
  translateFormulaReferences,
  type EvaluationResult,
  type FormulaNode,
  type NameRefNode,
  type CallExprNode,
} from "@bilig/formula";
import {
  ConfigValueTooBigError,
  ConfigValueTooSmallError,
  HeadlessArgumentError,
  HeadlessEvaluationSuspendedError,
  HeadlessOperationError,
  HeadlessParseError,
  HeadlessSheetError,
  ExpectedOneOfValuesError,
  FunctionPluginValidationError,
  InvalidArgumentsError,
  LanguageAlreadyRegisteredError,
  LanguageNotRegisteredError,
  NamedExpressionDoesNotExistError,
  NamedExpressionNameIsAlreadyTakenError,
  NamedExpressionNameIsInvalidError,
  NoOperationToRedoError,
  NoOperationToUndoError,
  NoRelativeAddressesAllowedError,
  NoSheetWithIdError,
  NoSheetWithNameError,
  NotAFormulaError,
  NothingToPasteError,
  SheetNameAlreadyTakenError,
  SheetSizeLimitExceededError,
  UnableToParseError,
} from "./errors.js";
import { buildMatrixMutationPlan } from "./matrix-mutation-plan.js";
import type {
  HeadlessAddressMappingAdapter,
  HeadlessAddressFormatOptions,
  HeadlessAddressLike,
  HeadlessArrayMappingAdapter,
  HeadlessAxisInterval,
  HeadlessAxisSwapMapping,
  HeadlessCellAddress,
  HeadlessCellRange,
  HeadlessCellType,
  HeadlessCellValueDetailedType,
  HeadlessCellValueType,
  HeadlessChange,
  HeadlessColumnSearchAdapter,
  HeadlessConfig,
  HeadlessDateTime,
  HeadlessDependencyGraphAdapter,
  HeadlessDependencyRef,
  HeadlessEvaluatorAdapter,
  HeadlessFunctionPluginDefinition,
  HeadlessFunctionTranslationsPackage,
  HeadlessGraphAdapter,
  HeadlessLanguagePackage,
  HeadlessLazilyTransformingAstServiceAdapter,
  HeadlessLicenseKeyValidityState,
  HeadlessNamedExpression,
  HeadlessRangeMappingAdapter,
  HeadlessSheet,
  HeadlessSheetDimensions,
  HeadlessSheetMappingAdapter,
  HeadlessSheets,
  HeadlessStats,
  HeadlessWorkbookDetailedEventMap,
  HeadlessWorkbookDetailedListener,
  HeadlessWorkbookEventName,
  HeadlessWorkbookInternals,
  HeadlessWorkbookListener,
  RawCellContent,
  SerializedHeadlessNamedExpression,
} from "./types.js";

type ListenerMap = {
  [EventName in HeadlessWorkbookEventName]: Set<HeadlessWorkbookListener<EventName>>;
};

type DetailedListenerMap = {
  [EventName in HeadlessWorkbookEventName]: Set<HeadlessWorkbookDetailedListener<EventName>>;
};

type DetailedEvent = {
  [EventName in HeadlessWorkbookEventName]: {
    eventName: EventName;
    payload: HeadlessWorkbookDetailedEventMap[EventName];
  };
}[HeadlessWorkbookEventName];

interface InternalNamedExpression {
  publicName: string;
  normalizedName: string;
  internalName: string;
  scope?: number;
  expression: RawCellContent;
  options?: Record<string, string | number | boolean>;
}

interface InternalFunctionBinding {
  pluginId: string;
  publicName: string;
  internalName: string;
  implementation?: (...args: CellValue[]) => EvaluationResult | CellValue;
}

interface SheetStateSnapshot {
  sheetId: number;
  sheetName: string;
  order: number;
  cells: Map<string, CellValue>;
}

type VisibilitySnapshot = Map<number, SheetStateSnapshot>;
type NamedExpressionValueSnapshot = Map<string, CellValue | CellValue[][]>;

interface TrackedCellRef {
  sheetId: number;
  sheetName: string;
  address: string;
  row: number;
  col: number;
}

interface ClipboardPayload {
  sourceAnchor: HeadlessCellAddress;
  serialized: RawCellContent[][];
  values: CellValue[][];
}

type QueuedEvent = Extract<
  DetailedEvent,
  {
    eventName:
      | "sheetAdded"
      | "sheetRemoved"
      | "sheetRenamed"
      | "namedExpressionAdded"
      | "namedExpressionRemoved";
  }
>;

function cloneTrackedEngineEvent(event: EngineEvent): EngineEvent {
  return {
    ...event,
    changedCellIndices: Array.from(event.changedCellIndices),
    invalidatedRanges: event.invalidatedRanges.map((range) => ({ ...range })),
    invalidatedRows: event.invalidatedRows.map((range) => ({ ...range })),
    invalidatedColumns: event.invalidatedColumns.map((range) => ({ ...range })),
    metrics: { ...event.metrics },
  };
}

interface HistoryRecord {
  forward: { ops: unknown[]; potentialNewCells?: number };
  inverse: { ops: unknown[]; potentialNewCells?: number };
}

const DEFAULT_CONFIG: Readonly<HeadlessConfig> = Object.freeze({
  accentSensitive: false,
  caseSensitive: false,
  caseFirst: "false",
  chooseAddressMappingPolicy: undefined,
  context: undefined,
  currencySymbol: ["$"],
  dateFormats: [],
  functionArgSeparator: ",",
  decimalSeparator: ".",
  evaluateNullToZero: true,
  functionPlugins: [],
  ignorePunctuation: false,
  language: "enGB",
  ignoreWhiteSpace: "standard",
  leapYear1900: true,
  licenseKey: "internal",
  localeLang: "en-US",
  matchWholeCell: true,
  arrayColumnSeparator: ",",
  arrayRowSeparator: ";",
  maxRows: MAX_ROWS,
  maxColumns: MAX_COLS,
  nullDate: { year: 1899, month: 12, day: 30 },
  nullYear: 30,
  parseDateTime: undefined,
  precisionEpsilon: 1e-13,
  precisionRounding: 14,
  stringifyDateTime: undefined,
  stringifyDuration: undefined,
  smartRounding: true,
  thousandSeparator: ",",
  timeFormats: [],
  useArrayArithmetic: true,
  useColumnIndex: false,
  useStats: true,
  undoLimit: 100,
  useRegularExpressions: true,
  useWildcards: true,
});

const HEADLESS_CONFIG_KEYS = [
  "accentSensitive",
  "caseSensitive",
  "caseFirst",
  "chooseAddressMappingPolicy",
  "context",
  "currencySymbol",
  "dateFormats",
  "functionArgSeparator",
  "decimalSeparator",
  "evaluateNullToZero",
  "functionPlugins",
  "ignorePunctuation",
  "language",
  "ignoreWhiteSpace",
  "leapYear1900",
  "licenseKey",
  "localeLang",
  "matchWholeCell",
  "arrayColumnSeparator",
  "arrayRowSeparator",
  "maxRows",
  "maxColumns",
  "nullDate",
  "nullYear",
  "parseDateTime",
  "precisionEpsilon",
  "precisionRounding",
  "stringifyDateTime",
  "stringifyDuration",
  "smartRounding",
  "thousandSeparator",
  "timeFormats",
  "useArrayArithmetic",
  "useColumnIndex",
  "useStats",
  "undoLimit",
  "useRegularExpressions",
  "useWildcards",
] as const satisfies readonly (keyof HeadlessConfig)[];

const HEADLESS_PUBLIC_ERROR_NAMES = new Set([
  "ConfigValueTooBigError",
  "ConfigValueTooSmallError",
  "EvaluationSuspendedError",
  "ExpectedOneOfValuesError",
  "ExpectedValueOfTypeError",
  "FunctionPluginValidationError",
  "InvalidAddressError",
  "InvalidArgumentsError",
  "LanguageAlreadyRegisteredError",
  "LanguageNotRegisteredError",
  "MissingTranslationError",
  "NamedExpressionDoesNotExistError",
  "NamedExpressionNameIsAlreadyTakenError",
  "NamedExpressionNameIsInvalidError",
  "NoOperationToRedoError",
  "NoOperationToUndoError",
  "NoRelativeAddressesAllowedError",
  "NoSheetWithIdError",
  "NoSheetWithNameError",
  "NotAFormulaError",
  "NothingToPasteError",
  "ProtectedFunctionTranslationError",
  "SheetNameAlreadyTakenError",
  "SheetSizeLimitExceededError",
  "SourceLocationHasArrayError",
  "TargetLocationHasArrayError",
  "UnableToParseError",
  "HeadlessArgumentError",
  "HeadlessConfigError",
  "HeadlessSheetError",
  "HeadlessNamedExpressionError",
  "HeadlessClipboardError",
  "HeadlessEvaluationSuspendedError",
  "HeadlessParseError",
  "HeadlessOperationError",
]);

const HEADLESS_VERSION = "0.1.2";
const HEADLESS_BUILD_DATE = "2026-04-10";
const HEADLESS_RELEASE_DATE = "2026-04-10";

const globalCustomFunctions = new Map<
  string,
  (...args: CellValue[]) => EvaluationResult | CellValue | undefined
>();

let customAdapterInstalled = false;
let nextWorkbookId = 1;

function ensureCustomAdapterInstalled(): void {
  if (customAdapterInstalled) {
    return;
  }
  installExternalFunctionAdapter({
    surface: "host",
    resolveFunction(name) {
      const implementation = globalCustomFunctions.get(name.trim().toUpperCase());
      if (!implementation) {
        return undefined;
      }
      return {
        kind: "scalar",
        implementation: (...args: CellValue[]) => {
          const result = implementation(...args);
          if (!result) {
            return errorValue(ErrorCode.Value);
          }
          return scalarFromResult(result);
        },
      };
    },
  });
  customAdapterInstalled = true;
}

function clonePluginDefinition(
  plugin: HeadlessFunctionPluginDefinition,
): HeadlessFunctionPluginDefinition {
  return {
    ...plugin,
    implementedFunctions: Object.fromEntries(
      Object.entries(plugin.implementedFunctions).map(([name, metadata]) => [
        name,
        { ...metadata },
      ]),
    ),
    aliases: plugin.aliases ? { ...plugin.aliases } : undefined,
    functions: plugin.functions ? { ...plugin.functions } : undefined,
  };
}

function cloneConfig(config: HeadlessConfig): HeadlessConfig {
  return {
    ...config,
    currencySymbol: config.currencySymbol ? [...config.currencySymbol] : undefined,
    dateFormats: config.dateFormats ? [...config.dateFormats] : undefined,
    functionPlugins: config.functionPlugins
      ? config.functionPlugins.map((plugin) => clonePluginDefinition(plugin))
      : undefined,
    nullDate: config.nullDate ? { ...config.nullDate } : undefined,
    timeFormats: config.timeFormats ? [...config.timeFormats] : undefined,
  };
}

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function scalarValueFromLiteral(value: LiteralInput): CellValue {
  if (value === null) {
    return emptyValue();
  }
  if (typeof value === "number") {
    return { tag: ValueTag.Number, value };
  }
  if (typeof value === "boolean") {
    return { tag: ValueTag.Boolean, value };
  }
  return { tag: ValueTag.String, value, stringId: 0 };
}

function scalarFromResult(result: EvaluationResult | CellValue): CellValue {
  if (!isArrayValue(result)) {
    return result;
  }
  return result.values[0] ?? emptyValue();
}

function valuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false;
  }
  switch (left.tag) {
    case ValueTag.Number:
      return right.tag === ValueTag.Number && left.value === right.value;
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value;
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value;
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code;
    case ValueTag.Empty:
      return true;
    default:
      return false;
  }
}

function matrixValuesEqual(
  left: CellValue | CellValue[][] | undefined,
  right: CellValue | CellValue[][] | undefined,
): boolean {
  if (!left && !right) {
    return true;
  }
  if (!left || !right) {
    return false;
  }
  if (isCellValueMatrix(left) !== isCellValueMatrix(right)) {
    return false;
  }
  if (!isCellValueMatrix(left) && !isCellValueMatrix(right)) {
    return valuesEqual(left, right);
  }
  if (!isCellValueMatrix(left) || !isCellValueMatrix(right)) {
    return false;
  }
  const leftMatrix = left;
  const rightMatrix = right;
  if (leftMatrix.length !== rightMatrix.length) {
    return false;
  }
  return leftMatrix.every((row: CellValue[], rowIndex: number) => {
    const otherRow = rightMatrix[rowIndex];
    if (!otherRow || row.length !== otherRow.length) {
      return false;
    }
    return row.every((value: CellValue, columnIndex: number) => {
      const otherValue = otherRow[columnIndex];
      if (!otherValue) {
        return false;
      }
      return valuesEqual(value, otherValue);
    });
  });
}

function normalizeName(name: string): string {
  return name.trim().toUpperCase();
}

function makeNamedExpressionKey(name: string, scope?: number): string {
  return `${scope ?? "workbook"}:${normalizeName(name)}`;
}

function makeInternalScopedName(scope: number, name: string): string {
  return `__BILIG_HEADLESS_SCOPE_${scope}_${normalizeName(name)}`;
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === "string" && content.trim().startsWith("=");
}

function isCellValueMatrix(value: CellValue | CellValue[][]): value is CellValue[][] {
  return Array.isArray(value);
}

function isHeadlessSheetMatrix(value: RawCellContent | HeadlessSheet): value is HeadlessSheet {
  return Array.isArray(value);
}

function matrixContainsFormulaContent(content: HeadlessSheet): boolean {
  return content.some((row) => row.some((cell) => isFormulaContent(cell)));
}

function isDeferredBatchLiteralContent(content: RawCellContent): boolean {
  return (
    content === null ||
    typeof content === "boolean" ||
    typeof content === "number" ||
    typeof content === "string"
  );
}

function stripLeadingEquals(formula: string): string {
  return formula.trim().startsWith("=") ? formula.trim().slice(1) : formula.trim();
}

function assertRowAndColumn(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new InvalidArgumentsError(`${label} to be a non-negative integer`);
  }
}

function assertRange(range: HeadlessCellRange): void {
  assertRowAndColumn(range.start.sheet, "start.sheet");
  assertRowAndColumn(range.start.row, "start.row");
  assertRowAndColumn(range.start.col, "start.col");
  assertRowAndColumn(range.end.sheet, "end.sheet");
  assertRowAndColumn(range.end.row, "end.row");
  assertRowAndColumn(range.end.col, "end.col");
  if (range.start.sheet !== range.end.sheet) {
    throw new HeadlessArgumentError("Ranges must stay on a single sheet");
  }
}

function isCellRange(value: HeadlessAddressLike): value is HeadlessCellRange {
  return "start" in value && "end" in value;
}

function cloneCellValue(value: CellValue): CellValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return emptyValue();
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: value.value };
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: value.value };
    case ValueTag.String:
      return { tag: ValueTag.String, value: value.value, stringId: value.stringId };
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: value.code };
    default:
      return emptyValue();
  }
}

function transformFormulaNode(
  node: FormulaNode,
  transform: (current: FormulaNode) => FormulaNode,
): FormulaNode {
  const current = transform(node);
  switch (current.kind) {
    case "BooleanLiteral":
    case "CellRef":
    case "ColumnRef":
    case "ErrorLiteral":
    case "NameRef":
    case "NumberLiteral":
    case "RangeRef":
    case "RowRef":
    case "SpillRef":
    case "StringLiteral":
    case "StructuredRef":
      return current;
    case "UnaryExpr":
      return {
        ...current,
        argument: transformFormulaNode(current.argument, transform),
      };
    case "BinaryExpr":
      return {
        ...current,
        left: transformFormulaNode(current.left, transform),
        right: transformFormulaNode(current.right, transform),
      };
    case "CallExpr":
      return {
        ...current,
        args: current.args.map((argument) => transformFormulaNode(argument, transform)),
      };
    case "InvokeExpr":
      return {
        ...current,
        callee: transformFormulaNode(current.callee, transform),
        args: current.args.map((argument) => transformFormulaNode(argument, transform)),
      };
    default:
      return current;
  }
}

function collectFormulaNameRefs(node: FormulaNode, output: Set<string>): void {
  switch (node.kind) {
    case "BooleanLiteral":
    case "CellRef":
    case "ColumnRef":
    case "ErrorLiteral":
    case "NameRef":
      if (node.kind === "NameRef") {
        output.add(node.name);
      }
      return;
    case "NumberLiteral":
    case "RangeRef":
    case "RowRef":
    case "SpillRef":
    case "StringLiteral":
    case "StructuredRef":
      return;
    case "UnaryExpr":
      collectFormulaNameRefs(node.argument, output);
      return;
    case "BinaryExpr":
      collectFormulaNameRefs(node.left, output);
      collectFormulaNameRefs(node.right, output);
      return;
    case "CallExpr":
      node.args.forEach((argument) => collectFormulaNameRefs(argument, output));
      return;
    case "InvokeExpr":
      collectFormulaNameRefs(node.callee, output);
      node.args.forEach((argument) => collectFormulaNameRefs(argument, output));
      return;
    default:
      return;
  }
}

function isAbsoluteCellReference(value: string): boolean {
  return /^\$[A-Z]+\$[1-9][0-9]*$/.test(value.toUpperCase());
}

function isAbsoluteRowReference(value: string): boolean {
  return /^\$[1-9][0-9]*$/.test(value);
}

function isAbsoluteColumnReference(value: string): boolean {
  return /^\$[A-Z]+$/.test(value.toUpperCase());
}

function formulaHasRelativeReferences(node: FormulaNode): boolean {
  switch (node.kind) {
    case "BooleanLiteral":
    case "ErrorLiteral":
    case "NameRef":
    case "NumberLiteral":
    case "StringLiteral":
    case "StructuredRef":
      return false;
    case "CellRef":
    case "SpillRef":
      return !isAbsoluteCellReference(node.ref);
    case "RowRef":
      return !isAbsoluteRowReference(node.ref);
    case "ColumnRef":
      return !isAbsoluteColumnReference(node.ref);
    case "RangeRef":
      if (node.refKind === "cells") {
        return !isAbsoluteCellReference(node.start) || !isAbsoluteCellReference(node.end);
      }
      if (node.refKind === "rows") {
        return !isAbsoluteRowReference(node.start) || !isAbsoluteRowReference(node.end);
      }
      return !isAbsoluteColumnReference(node.start) || !isAbsoluteColumnReference(node.end);
    case "UnaryExpr":
      return formulaHasRelativeReferences(node.argument);
    case "BinaryExpr":
      return formulaHasRelativeReferences(node.left) || formulaHasRelativeReferences(node.right);
    case "CallExpr":
      return node.args.some((argument) => formulaHasRelativeReferences(argument));
    case "InvokeExpr":
      return (
        formulaHasRelativeReferences(node.callee) ||
        node.args.some((argument) => formulaHasRelativeReferences(argument))
      );
    default:
      return false;
  }
}

function compareSheetNames(left: string, right: string): number {
  return left.localeCompare(right);
}

function checkHeadlessLicenseKeyValidity(
  licenseKey: string | undefined,
): HeadlessLicenseKeyValidityState {
  if (!licenseKey || licenseKey.trim().length === 0) {
    return "missing";
  }
  if (
    licenseKey === "internal" ||
    licenseKey === "gpl-v3" ||
    licenseKey === "internal-use-in-handsontable"
  ) {
    return "valid";
  }
  return "invalid";
}

function validateHeadlessConfig(config: HeadlessConfig): void {
  if (config.maxRows !== undefined && (!Number.isInteger(config.maxRows) || config.maxRows < 1)) {
    throw new ConfigValueTooSmallError("maxRows", 1);
  }
  if (
    config.maxColumns !== undefined &&
    (!Number.isInteger(config.maxColumns) || config.maxColumns < 1)
  ) {
    throw new ConfigValueTooSmallError("maxColumns", 1);
  }
  if ((config.maxRows ?? MAX_ROWS) > MAX_ROWS) {
    throw new ConfigValueTooBigError("maxRows", MAX_ROWS);
  }
  if ((config.maxColumns ?? MAX_COLS) > MAX_COLS) {
    throw new ConfigValueTooBigError("maxColumns", MAX_COLS);
  }
  if (
    config.decimalSeparator !== undefined &&
    config.decimalSeparator !== "." &&
    config.decimalSeparator !== ","
  ) {
    throw new ExpectedOneOfValuesError('".", ","', "decimalSeparator");
  }
  if (
    config.arrayColumnSeparator !== undefined &&
    config.arrayColumnSeparator !== "," &&
    config.arrayColumnSeparator !== ";"
  ) {
    throw new ExpectedOneOfValuesError('",", ";"', "arrayColumnSeparator");
  }
  if (
    config.arrayRowSeparator !== undefined &&
    config.arrayRowSeparator !== ";" &&
    config.arrayRowSeparator !== "|"
  ) {
    throw new ExpectedOneOfValuesError('";", "|"', "arrayRowSeparator");
  }
  if (
    config.ignoreWhiteSpace !== undefined &&
    config.ignoreWhiteSpace !== "standard" &&
    config.ignoreWhiteSpace !== "any"
  ) {
    throw new ExpectedOneOfValuesError('"standard", "any"', "ignoreWhiteSpace");
  }
  if (
    config.caseFirst !== undefined &&
    config.caseFirst !== "upper" &&
    config.caseFirst !== "lower" &&
    config.caseFirst !== "false"
  ) {
    throw new ExpectedOneOfValuesError('"upper", "lower", "false"', "caseFirst");
  }
}

function validateSheetWithinLimits(
  sheetName: string,
  sheet: HeadlessSheet,
  config: HeadlessConfig,
): void {
  const height = sheet.length;
  const width = Math.max(0, ...sheet.map((row) => row.length));
  if (height > (config.maxRows ?? MAX_ROWS) || width > (config.maxColumns ?? MAX_COLS)) {
    throw new SheetSizeLimitExceededError();
  }
  sheet.forEach((row) => {
    if (!Array.isArray(row)) {
      throw new UnableToParseError({ sheetName, reason: "Rows must be arrays" });
    }
  });
}

function isHistoryRecordArray(value: unknown): value is HistoryRecord[] {
  return Array.isArray(value);
}

function withEventChanges(event: QueuedEvent, changes: HeadlessChange[]): QueuedEvent {
  switch (event.eventName) {
    case "sheetAdded":
      return event;
    case "sheetRemoved":
      return {
        eventName: "sheetRemoved",
        payload: {
          ...event.payload,
          changes,
        },
      };
    case "sheetRenamed":
      return event;
    case "namedExpressionAdded":
      return {
        eventName: "namedExpressionAdded",
        payload: {
          ...event.payload,
          changes,
        },
      };
    case "namedExpressionRemoved":
      return {
        eventName: "namedExpressionRemoved",
        payload: {
          ...event.payload,
          changes,
        },
      };
  }
}

function quoteSheetNameIfNeeded(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`;
}

function formatQualifiedCellAddress(
  sheetName: string | undefined,
  row: number,
  col: number,
): string {
  const base = formatAddress(row, col);
  return sheetName ? `${quoteSheetNameIfNeeded(sheetName)}!${base}` : base;
}

class HeadlessEmitter {
  private readonly listeners: ListenerMap = {
    sheetAdded: new Set(),
    sheetRemoved: new Set(),
    sheetRenamed: new Set(),
    namedExpressionAdded: new Set(),
    namedExpressionRemoved: new Set(),
    valuesUpdated: new Set(),
    evaluationSuspended: new Set(),
    evaluationResumed: new Set(),
  };

  private readonly detailedListeners: DetailedListenerMap = {
    sheetAdded: new Set(),
    sheetRemoved: new Set(),
    sheetRenamed: new Set(),
    namedExpressionAdded: new Set(),
    namedExpressionRemoved: new Set(),
    valuesUpdated: new Set(),
    evaluationSuspended: new Set(),
    evaluationResumed: new Set(),
  };

  on<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookListener<EventName>,
  ): void {
    this.listeners[eventName].add(listener);
  }

  off<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookListener<EventName>,
  ): void {
    this.listeners[eventName].delete(listener);
  }

  once<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookListener<EventName>,
  ): void {
    const wrapper: HeadlessWorkbookListener<EventName> = (...args) => {
      this.off(eventName, wrapper);
      listener(...args);
    };
    this.on(eventName, wrapper);
  }

  onDetailed<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookDetailedListener<EventName>,
  ): void {
    this.detailedListeners[eventName].add(listener);
  }

  offDetailed<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookDetailedListener<EventName>,
  ): void {
    this.detailedListeners[eventName].delete(listener);
  }

  onceDetailed<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookDetailedListener<EventName>,
  ): void {
    const wrapper: HeadlessWorkbookDetailedListener<EventName> = (payload) => {
      this.offDetailed(eventName, wrapper);
      listener(payload);
    };
    this.onDetailed(eventName, wrapper);
  }

  emitDetailed(event: DetailedEvent): void {
    this.dispatchDetailed(event);
  }

  private dispatchDetailed(event: DetailedEvent): void {
    switch (event.eventName) {
      case "sheetAdded":
        this.listeners.sheetAdded.forEach((listener) => {
          listener(event.payload.sheetName);
        });
        this.detailedListeners.sheetAdded.forEach((listener) => {
          listener(event.payload);
        });
        return;
      case "sheetRemoved":
        this.listeners.sheetRemoved.forEach((listener) => {
          listener(event.payload.sheetName, event.payload.changes);
        });
        this.detailedListeners.sheetRemoved.forEach((listener) => {
          listener(event.payload);
        });
        return;
      case "sheetRenamed":
        this.listeners.sheetRenamed.forEach((listener) => {
          listener(event.payload.oldName, event.payload.newName);
        });
        this.detailedListeners.sheetRenamed.forEach((listener) => {
          listener(event.payload);
        });
        return;
      case "namedExpressionAdded":
        this.listeners.namedExpressionAdded.forEach((listener) => {
          listener(event.payload.name, event.payload.changes);
        });
        this.detailedListeners.namedExpressionAdded.forEach((listener) => {
          listener(event.payload);
        });
        return;
      case "namedExpressionRemoved":
        this.listeners.namedExpressionRemoved.forEach((listener) => {
          listener(event.payload.name, event.payload.changes);
        });
        this.detailedListeners.namedExpressionRemoved.forEach((listener) => {
          listener(event.payload);
        });
        return;
      case "valuesUpdated":
        this.listeners.valuesUpdated.forEach((listener) => {
          listener(event.payload.changes);
        });
        this.detailedListeners.valuesUpdated.forEach((listener) => {
          listener(event.payload);
        });
        return;
      case "evaluationSuspended":
        this.listeners.evaluationSuspended.forEach((listener) => {
          listener();
        });
        this.detailedListeners.evaluationSuspended.forEach((listener) => {
          listener(event.payload);
        });
        return;
      case "evaluationResumed":
        this.listeners.evaluationResumed.forEach((listener) => {
          listener(event.payload.changes);
        });
        this.detailedListeners.evaluationResumed.forEach((listener) => {
          listener(event.payload);
        });
    }
  }

  clear(): void {
    Object.values(this.listeners).forEach((listeners) => listeners.clear());
    Object.values(this.detailedListeners).forEach((listeners) => listeners.clear());
  }
}

export class HeadlessWorkbook {
  static version = HEADLESS_VERSION;
  static buildDate = HEADLESS_BUILD_DATE;
  static releaseDate = HEADLESS_RELEASE_DATE;
  static readonly languages: Record<string, HeadlessLanguagePackage> = {};
  static readonly defaultConfig: HeadlessConfig = cloneConfig(DEFAULT_CONFIG);

  private static readonly languageRegistry = new Map<string, HeadlessLanguagePackage>();
  private static readonly functionPluginRegistry = new Map<
    string,
    HeadlessFunctionPluginDefinition
  >();

  readonly workbookId = nextWorkbookId++;
  private engine: SpreadsheetEngine;
  private readonly emitter = new HeadlessEmitter();
  private readonly namedExpressions = new Map<string, InternalNamedExpression>();
  private readonly functionSnapshot = new Map<string, InternalFunctionBinding>();
  private readonly functionAliasLookup = new Map<string, InternalFunctionBinding>();
  private readonly internalFunctionLookup = new Map<string, InternalFunctionBinding>();
  readonly internals: HeadlessWorkbookInternals;
  private config: HeadlessConfig;
  private clipboard: ClipboardPayload | null = null;
  private visibilityCache: VisibilitySnapshot | null = null;
  private namedExpressionValueCache: NamedExpressionValueSnapshot | null = null;
  private batchDepth = 0;
  private batchStartVisibility: VisibilitySnapshot | null = null;
  private batchStartNamedValues: NamedExpressionValueSnapshot | null = null;
  private batchUndoStackLength = 0;
  private pendingBatchOps: EngineOp[] = [];
  private pendingBatchPotentialNewCells = 0;
  private evaluationSuspended = false;
  private suspendedVisibility: VisibilitySnapshot | null = null;
  private suspendedNamedValues: NamedExpressionValueSnapshot | null = null;
  private queuedEvents: QueuedEvent[] = [];
  private trackedEngineEvents: EngineEvent[] = [];
  private unsubscribeEngineEvents: (() => void) | null = null;
  private disposed = false;

  private constructor(configInput: HeadlessConfig = {}) {
    ensureCustomAdapterInstalled();
    validateHeadlessConfig(configInput);
    this.config = {
      ...cloneConfig(DEFAULT_CONFIG),
      ...cloneConfig(configInput),
    };
    this.engine = new SpreadsheetEngine({
      workbookName: "Workbook",
      useColumnIndex: this.config.useColumnIndex,
    });
    this.attachEngineEventTracking();
    this.captureFunctionRegistry();
    this.internals = Object.freeze({
      graph: Object.freeze<HeadlessGraphAdapter>({
        getDependents: (reference) => this.getCellDependents(reference),
        getPrecedents: (reference) => this.getCellPrecedents(reference),
      }),
      rangeMapping: Object.freeze<HeadlessRangeMappingAdapter>({
        getValues: (range) => this.getRangeValues(range),
        getSerialized: (range) => this.getRangeSerialized(range),
      }),
      arrayMapping: Object.freeze<HeadlessArrayMappingAdapter>({
        isPartOfArray: (address) => this.isCellPartOfArray(address),
        getFormula: (address) => this.getCellFormula(address),
      }),
      sheetMapping: Object.freeze<HeadlessSheetMappingAdapter>({
        getSheetName: (sheetId) => this.getSheetName(sheetId),
        getSheetId: (name) => this.getSheetId(name),
        getSheetNames: () => this.getSheetNames(),
        countSheets: () => this.countSheets(),
      }),
      addressMapping: Object.freeze<HeadlessAddressMappingAdapter>({
        has: (address) => !this.isCellEmpty(address) || this.doesCellHaveFormula(address),
        getValue: (address) => this.getCellValue(address),
        getFormula: (address) => this.getCellFormula(address),
      }),
      dependencyGraph: Object.freeze<HeadlessDependencyGraphAdapter>({
        getCellDependents: (reference) => this.getCellDependents(reference),
        getCellPrecedents: (reference) => this.getCellPrecedents(reference),
      }),
      evaluator: Object.freeze<HeadlessEvaluatorAdapter>({
        recalculate: () => this.rebuildAndRecalculate(),
        calculateFormula: (formula, scope) => this.calculateFormula(formula, scope),
      }),
      columnSearch: Object.freeze<HeadlessColumnSearchAdapter>({
        find: (sheetId, column, matcher) => {
          const dimensions = this.getSheetDimensions(sheetId);
          const matches: HeadlessCellAddress[] = [];
          for (let row = 0; row < dimensions.height; row += 1) {
            const address = { sheet: sheetId, row, col: column };
            const value = this.getCellValue(address);
            const isMatch =
              typeof matcher === "string"
                ? value.tag === ValueTag.String && value.value === matcher
                : matcher(value);
            if (isMatch) {
              matches.push(address);
            }
          }
          return matches;
        },
      }),
      lazilyTransformingAstService: Object.freeze<HeadlessLazilyTransformingAstServiceAdapter>({
        normalizeFormula: (formula) => this.normalizeFormula(formula),
        validateFormula: (formula) => this.validateFormula(formula),
        getNamedExpressionsFromFormula: (formula) => this.getNamedExpressionsFromFormula(formula),
      }),
    });
  }

  static buildEmpty(
    configInput: HeadlessConfig = {},
    namedExpressions: readonly SerializedHeadlessNamedExpression[] = [],
  ): HeadlessWorkbook {
    const workbook = new HeadlessWorkbook(configInput);
    namedExpressions.forEach((expression) => {
      workbook.upsertNamedExpressionInternal(expression, { duringInitialization: true });
    });
    workbook.clearHistoryStacks();
    workbook.primeChangeTrackingCaches();
    return workbook;
  }

  static buildFromArray(
    sheet: HeadlessSheet,
    configInput: HeadlessConfig = {},
    namedExpressions: readonly SerializedHeadlessNamedExpression[] = [],
  ): HeadlessWorkbook {
    return this.buildFromSheets({ Sheet1: sheet }, configInput, namedExpressions);
  }

  static buildFromSheets(
    sheets: HeadlessSheets,
    configInput: HeadlessConfig = {},
    namedExpressions: readonly SerializedHeadlessNamedExpression[] = [],
  ): HeadlessWorkbook {
    const workbook = new HeadlessWorkbook(configInput);
    Object.entries(sheets).forEach(([sheetName, sheet]) => {
      validateSheetWithinLimits(sheetName, sheet, workbook.config);
    });
    Object.keys(sheets).forEach((sheetName) => {
      workbook.engine.createSheet(sheetName);
    });
    namedExpressions.forEach((expression) => {
      workbook.upsertNamedExpressionInternal(expression, { duringInitialization: true });
    });
    Object.entries(sheets).forEach(([sheetName, sheet]) => {
      const sheetId = workbook.requireSheetId(sheetName);
      workbook.replaceSheetContentInternal(sheetId, sheet, { duringInitialization: true });
    });
    workbook.clearHistoryStacks();
    workbook.primeChangeTrackingCaches();
    return workbook;
  }

  static getLanguage(languageCode: string): HeadlessLanguagePackage {
    const language = this.languageRegistry.get(languageCode);
    if (!language) {
      throw new LanguageNotRegisteredError(languageCode);
    }
    return structuredClone(language);
  }

  static registerLanguage(languageCode: string, languagePackage: HeadlessLanguagePackage): void {
    if (this.languageRegistry.has(languageCode)) {
      throw new LanguageAlreadyRegisteredError(languageCode);
    }
    this.languageRegistry.set(languageCode, structuredClone(languagePackage));
    this.languages[languageCode] = structuredClone(languagePackage);
  }

  static unregisterLanguage(languageCode: string): void {
    if (!this.languageRegistry.delete(languageCode)) {
      throw new LanguageNotRegisteredError(languageCode);
    }
    delete this.languages[languageCode];
  }

  static getRegisteredLanguagesCodes(): string[] {
    return [...this.languageRegistry.keys()].toSorted(compareSheetNames);
  }

  static registerFunctionPlugin(
    plugin: HeadlessFunctionPluginDefinition,
    translations?: HeadlessFunctionTranslationsPackage,
  ): void {
    this.functionPluginRegistry.set(plugin.id, clonePluginDefinition(plugin));
    if (translations) {
      this.loadFunctionTranslations(translations);
    }
  }

  static unregisterFunctionPlugin(plugin: HeadlessFunctionPluginDefinition | string): void {
    const pluginId = typeof plugin === "string" ? plugin : plugin.id;
    this.functionPluginRegistry.delete(pluginId);
  }

  static registerFunction(
    functionId: string,
    plugin: HeadlessFunctionPluginDefinition,
    translations?: HeadlessFunctionTranslationsPackage,
  ): void {
    const existing = this.functionPluginRegistry.get(plugin.id);
    const nextPlugin = clonePluginDefinition(existing ?? plugin);
    if (!nextPlugin.implementedFunctions[functionId]) {
      throw FunctionPluginValidationError.functionNotDeclaredInPlugin(functionId, plugin.id);
    }
    this.functionPluginRegistry.set(nextPlugin.id, nextPlugin);
    if (translations) {
      this.loadFunctionTranslations(translations);
    }
  }

  static unregisterFunction(functionId: string): void {
    const normalized = functionId.trim().toUpperCase();
    this.functionPluginRegistry.forEach((plugin, pluginId) => {
      if (!plugin.implementedFunctions[normalized]) {
        return;
      }
      const nextPlugin = clonePluginDefinition(plugin);
      delete nextPlugin.implementedFunctions[normalized];
      if (nextPlugin.functions) {
        delete nextPlugin.functions[normalized];
      }
      if (nextPlugin.aliases) {
        Object.entries(nextPlugin.aliases).forEach(([alias, target]) => {
          if (
            target.trim().toUpperCase() === normalized ||
            alias.trim().toUpperCase() === normalized
          ) {
            delete nextPlugin.aliases![alias];
          }
        });
      }
      this.functionPluginRegistry.set(pluginId, nextPlugin);
    });
  }

  static unregisterAllFunctions(): void {
    this.functionPluginRegistry.clear();
  }

  static getRegisteredFunctionNames(languageCode?: string): string[] {
    const normalized = languageCode ?? "enGB";
    const language = this.languageRegistry.get(normalized);
    const functions = [...this.functionPluginRegistry.values()].flatMap((plugin) =>
      Object.keys(plugin.implementedFunctions),
    );
    if (!language?.functions) {
      return functions.toSorted(compareSheetNames);
    }
    return functions.map((name) => language.functions?.[name] ?? name).toSorted(compareSheetNames);
  }

  static getFunctionPlugin(functionId: string): HeadlessFunctionPluginDefinition | undefined {
    const normalized = functionId.trim().toUpperCase();
    const plugin = [...this.functionPluginRegistry.values()].find(
      (candidate) =>
        candidate.implementedFunctions[normalized] !== undefined ||
        candidate.aliases?.[normalized] !== undefined,
    );
    return plugin ? clonePluginDefinition(plugin) : undefined;
  }

  static getAllFunctionPlugins(): HeadlessFunctionPluginDefinition[] {
    return [...this.functionPluginRegistry.values()].map((plugin) => clonePluginDefinition(plugin));
  }

  private static loadFunctionTranslations(translations: HeadlessFunctionTranslationsPackage): void {
    Object.entries(translations).forEach(([languageCode, functionTranslations]) => {
      const existing = this.languageRegistry.get(languageCode);
      if (!existing) {
        throw new LanguageNotRegisteredError(languageCode);
      }
      const nextLanguage: HeadlessLanguagePackage = {
        ...structuredClone(existing),
        functions: {
          ...existing.functions,
          ...functionTranslations,
        },
      };
      this.languageRegistry.set(languageCode, nextLanguage);
      this.languages[languageCode] = structuredClone(nextLanguage);
    });
  }

  on<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookListener<EventName>,
  ): void {
    this.assertNotDisposed();
    this.emitter.on(eventName, listener);
  }

  once<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookListener<EventName>,
  ): void {
    this.assertNotDisposed();
    this.emitter.once(eventName, listener);
  }

  off<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookListener<EventName>,
  ): void {
    this.emitter.off(eventName, listener);
  }

  onDetailed<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookDetailedListener<EventName>,
  ): void {
    this.assertNotDisposed();
    this.emitter.onDetailed(eventName, listener);
  }

  onceDetailed<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookDetailedListener<EventName>,
  ): void {
    this.assertNotDisposed();
    this.emitter.onceDetailed(eventName, listener);
  }

  offDetailed<EventName extends HeadlessWorkbookEventName>(
    eventName: EventName,
    listener: HeadlessWorkbookDetailedListener<EventName>,
  ): void {
    this.emitter.offDetailed(eventName, listener);
  }

  getConfig(): HeadlessConfig {
    return cloneConfig(this.config);
  }

  get graph(): HeadlessGraphAdapter {
    return this.internals.graph;
  }

  get rangeMapping(): HeadlessRangeMappingAdapter {
    return this.internals.rangeMapping;
  }

  get arrayMapping(): HeadlessArrayMappingAdapter {
    return this.internals.arrayMapping;
  }

  get sheetMapping(): HeadlessSheetMappingAdapter {
    return this.internals.sheetMapping;
  }

  get addressMapping(): HeadlessAddressMappingAdapter {
    return this.internals.addressMapping;
  }

  get dependencyGraph(): HeadlessDependencyGraphAdapter {
    return this.internals.dependencyGraph;
  }

  get evaluator(): HeadlessEvaluatorAdapter {
    return this.internals.evaluator;
  }

  get columnSearch(): HeadlessColumnSearchAdapter {
    return this.internals.columnSearch;
  }

  get lazilyTransformingAstService(): HeadlessLazilyTransformingAstServiceAdapter {
    return this.internals.lazilyTransformingAstService;
  }

  get licenseKeyValidityState(): HeadlessLicenseKeyValidityState {
    return checkHeadlessLicenseKeyValidity(this.config.licenseKey);
  }

  updateConfig(next: HeadlessConfig): void {
    this.assertNotDisposed();
    const merged = {
      ...this.config,
      ...cloneConfig(next),
    };
    const hasChanges = HEADLESS_CONFIG_KEYS.some(
      (key) => Object.hasOwn(next, key) && this.config[key] !== next[key],
    );
    if (!hasChanges) {
      return;
    }
    this.rebuildWithConfig(merged);
  }

  getStats(): HeadlessStats {
    this.assertNotDisposed();
    return {
      batchDepth: this.batchDepth,
      evaluationSuspended: this.evaluationSuspended,
      lastMetrics: structuredClone(this.engine.getLastMetrics()),
    };
  }

  rebuildAndRecalculate(): HeadlessChange[] {
    return this.captureChanges(undefined, () => {
      this.rebuildWithConfig(this.config);
    });
  }

  batch(batchOperations: () => void): HeadlessChange[] {
    this.assertNotDisposed();
    const isOutermost = this.batchDepth === 0;
    if (isOutermost) {
      this.batchStartVisibility = this.ensureVisibilityCache();
      this.batchStartNamedValues = this.ensureNamedExpressionValueCache();
      this.batchUndoStackLength = this.getUndoStack().length;
      this.drainTrackedEngineEvents();
    }
    this.batchDepth += 1;
    try {
      batchOperations();
    } finally {
      this.batchDepth -= 1;
      if (isOutermost) {
        this.flushPendingBatchOps();
        this.mergeUndoHistory(this.batchUndoStackLength);
      }
    }
    if (!isOutermost) {
      return [];
    }
    const changes = this.computeChangesAfterMutation(
      this.batchStartVisibility ?? new Map(),
      this.batchStartNamedValues ?? new Map(),
    );
    this.batchStartVisibility = null;
    this.batchStartNamedValues = null;
    if (!this.evaluationSuspended) {
      this.flushQueuedEvents();
      if (changes.length > 0) {
        this.emitter.emitDetailed({ eventName: "valuesUpdated", payload: { changes } });
      }
    }
    return changes;
  }

  suspendEvaluation(): void {
    this.assertNotDisposed();
    if (this.evaluationSuspended) {
      return;
    }
    this.evaluationSuspended = true;
    this.suspendedVisibility = this.ensureVisibilityCache();
    this.suspendedNamedValues = this.ensureNamedExpressionValueCache();
    this.drainTrackedEngineEvents();
    this.emitter.emitDetailed({ eventName: "evaluationSuspended", payload: {} });
  }

  resumeEvaluation(): HeadlessChange[] {
    this.assertNotDisposed();
    if (!this.evaluationSuspended) {
      return [];
    }
    const changes = this.computeChangesAfterMutation(
      this.suspendedVisibility ?? new Map(),
      this.suspendedNamedValues ?? new Map(),
    );
    this.evaluationSuspended = false;
    this.suspendedVisibility = null;
    this.suspendedNamedValues = null;
    this.flushQueuedEvents();
    this.emitter.emitDetailed({ eventName: "evaluationResumed", payload: { changes } });
    if (changes.length > 0) {
      this.emitter.emitDetailed({ eventName: "valuesUpdated", payload: { changes } });
    }
    return changes;
  }

  isEvaluationSuspended(): boolean {
    return this.evaluationSuspended;
  }

  undo(): HeadlessChange[] {
    this.assertNotDisposed();
    return this.captureChanges(undefined, () => {
      if (!this.engine.undo()) {
        throw new NoOperationToUndoError();
      }
    });
  }

  redo(): HeadlessChange[] {
    this.assertNotDisposed();
    return this.captureChanges(undefined, () => {
      if (!this.engine.redo()) {
        throw new NoOperationToRedoError();
      }
    });
  }

  isThereSomethingToUndo(): boolean {
    return this.getUndoStack().length > 0;
  }

  isThereSomethingToRedo(): boolean {
    return this.getRedoStack().length > 0;
  }

  clearUndoStack(): void {
    this.getUndoStack().length = 0;
  }

  clearRedoStack(): void {
    this.getRedoStack().length = 0;
  }

  copy(range: HeadlessCellRange): CellValue[][] {
    this.assertReadable();
    assertRange(range);
    const serialized = this.getRangeSerialized(range);
    const values = this.getRangeValues(range);
    this.clipboard = {
      sourceAnchor: { ...range.start },
      serialized,
      values,
    };
    return values;
  }

  cut(range: HeadlessCellRange): CellValue[][] {
    this.assertReadable();
    const values = this.copy(range);
    this.batch(() => {
      this.setCellContents(range.start, this.buildNullMatrixForRange(range));
    });
    return values;
  }

  paste(targetLeftCorner: HeadlessCellAddress): HeadlessChange[] {
    this.assertNotDisposed();
    if (!this.clipboard) {
      throw new NothingToPasteError();
    }
    return this.captureChanges(undefined, () => {
      this.applySerializedMatrix(
        targetLeftCorner,
        this.clipboard!.serialized,
        this.clipboard!.sourceAnchor,
      );
    });
  }

  isClipboardEmpty(): boolean {
    return this.clipboard === null;
  }

  clearClipboard(): void {
    this.clipboard = null;
  }

  getFillRangeData(
    source: HeadlessCellRange,
    target: HeadlessCellRange,
    offsetsFromTarget = false,
  ): RawCellContent[][] {
    assertRange(source);
    assertRange(target);
    const sourceSerialized = this.getRangeSerialized(source);
    const targetHeight = target.end.row - target.start.row + 1;
    const targetWidth = target.end.col - target.start.col + 1;
    const sourceHeight = Math.max(sourceSerialized.length, 1);
    const sourceWidth = Math.max(sourceSerialized[0]?.length ?? 0, 1);
    const output: RawCellContent[][] = [];
    for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
      const row: RawCellContent[] = [];
      for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
        const targetRow = target.start.row + rowOffset;
        const targetCol = target.start.col + colOffset;
        const sourceRow =
          (((targetRow - (offsetsFromTarget ? target.start.row : source.start.row)) %
            sourceHeight) +
            sourceHeight) %
          sourceHeight;
        const sourceCol =
          (((targetCol - (offsetsFromTarget ? target.start.col : source.start.col)) % sourceWidth) +
            sourceWidth) %
          sourceWidth;
        const raw = sourceSerialized[sourceRow]?.[sourceCol] ?? null;
        if (typeof raw === "string" && raw.startsWith("=")) {
          row.push(
            `=${translateFormulaReferences(
              raw.slice(1),
              targetRow - (source.start.row + sourceRow),
              targetCol - (source.start.col + sourceCol),
            )}`,
          );
        } else {
          row.push(raw);
        }
      }
      output.push(row);
    }
    return output;
  }

  getCellValue(address: HeadlessCellAddress): CellValue {
    this.assertReadable();
    return cloneCellValue(
      this.engine.getCellValue(this.sheetName(address.sheet), this.a1(address)),
    );
  }

  getCellFormula(address: HeadlessCellAddress): string | undefined {
    this.prepareReadableState();
    const cell = this.engine.getCell(this.sheetName(address.sheet), this.a1(address));
    if (!cell.formula) {
      return undefined;
    }
    return `=${this.restorePublicFormula(cell.formula, address.sheet)}`;
  }

  getCellHyperlink(address: HeadlessCellAddress): string | undefined {
    const formula = this.getCellFormula(address);
    if (!formula) {
      return undefined;
    }
    const parsed = parseFormula(stripLeadingEquals(formula));
    if (parsed.kind !== "CallExpr" || parsed.callee.trim().toUpperCase() !== "HYPERLINK") {
      return undefined;
    }
    const firstArgument = parsed.args[0];
    return firstArgument?.kind === "StringLiteral" ? firstArgument.value : undefined;
  }

  getCellSerialized(address: HeadlessCellAddress): RawCellContent {
    this.prepareReadableState();
    return this.cellSnapshotToRawContent(
      this.engine.getCell(this.sheetName(address.sheet), this.a1(address)),
      address.sheet,
    );
  }

  getRangeValues(range: HeadlessCellRange): CellValue[][] {
    this.assertReadable();
    const ref = this.rangeRef(range);
    return this.engine
      .getRangeValues(ref)
      .map((row: readonly CellValue[]) => row.map((value: CellValue) => cloneCellValue(value)));
  }

  getRangeFormulas(range: HeadlessCellRange): Array<Array<string | undefined>> {
    return this.getDenseRange(range, (address) => this.getCellFormula(address));
  }

  getRangeSerialized(range: HeadlessCellRange): RawCellContent[][] {
    return this.getDenseRange(range, (address) => this.getCellSerialized(address));
  }

  getSheetValues(sheetId: number): CellValue[][] {
    this.assertReadable();
    const dimensions = this.getSheetDimensions(sheetId);
    if (dimensions.width === 0 || dimensions.height === 0) {
      return [];
    }
    return this.getRangeValues({
      start: { sheet: sheetId, row: 0, col: 0 },
      end: { sheet: sheetId, row: dimensions.height - 1, col: dimensions.width - 1 },
    });
  }

  getSheetFormulas(sheetId: number): Array<Array<string | undefined>> {
    const dimensions = this.getSheetDimensions(sheetId);
    if (dimensions.width === 0 || dimensions.height === 0) {
      return [];
    }
    return this.getRangeFormulas({
      start: { sheet: sheetId, row: 0, col: 0 },
      end: { sheet: sheetId, row: dimensions.height - 1, col: dimensions.width - 1 },
    });
  }

  getSheetSerialized(sheetId: number): RawCellContent[][] {
    const dimensions = this.getSheetDimensions(sheetId);
    if (dimensions.width === 0 || dimensions.height === 0) {
      return [];
    }
    return this.getRangeSerialized({
      start: { sheet: sheetId, row: 0, col: 0 },
      end: { sheet: sheetId, row: dimensions.height - 1, col: dimensions.width - 1 },
    });
  }

  getAllSheetsValues(): Record<string, CellValue[][]> {
    this.assertReadable();
    return Object.fromEntries(
      this.listSheetRecords().map((sheet) => [sheet.name, this.getSheetValues(sheet.id)]),
    );
  }

  getAllSheetsFormulas(): Record<string, Array<Array<string | undefined>>> {
    return Object.fromEntries(
      this.listSheetRecords().map((sheet) => [sheet.name, this.getSheetFormulas(sheet.id)]),
    );
  }

  getAllSheetsSerialized(): Record<string, RawCellContent[][]> {
    return Object.fromEntries(
      this.listSheetRecords().map((sheet) => [sheet.name, this.getSheetSerialized(sheet.id)]),
    );
  }

  getAllSheetsDimensions(): Record<string, HeadlessSheetDimensions> {
    return Object.fromEntries(
      this.listSheetRecords().map((sheet) => [sheet.name, this.getSheetDimensions(sheet.id)]),
    );
  }

  getSheetDimensions(sheetId: number): HeadlessSheetDimensions {
    this.prepareReadableState();
    const sheet = this.sheetRecord(sheetId);
    let width = 0;
    let height = 0;
    sheet.grid.forEachCellEntry((_cellIndex: number, row: number, col: number) => {
      height = Math.max(height, row + 1);
      width = Math.max(width, col + 1);
    });
    return { width, height };
  }

  simpleCellAddressFromString(
    value: string,
    defaultSheetId?: number,
  ): HeadlessCellAddress | undefined {
    this.assertNotDisposed();
    const defaultSheetName =
      defaultSheetId !== undefined
        ? this.sheetName(defaultSheetId)
        : this.listSheetRecords().length === 1
          ? this.listSheetRecords()[0]!.name
          : undefined;
    try {
      const parsed = parseCellAddress(value, defaultSheetName);
      const sheetName = parsed.sheetName ?? defaultSheetName;
      if (!sheetName) {
        return undefined;
      }
      return {
        sheet: this.requireSheetId(sheetName),
        row: parsed.row,
        col: parsed.col,
      };
    } catch {
      return undefined;
    }
  }

  simpleCellRangeFromString(value: string, defaultSheetId?: number): HeadlessCellRange | undefined {
    this.assertNotDisposed();
    const defaultSheetName =
      defaultSheetId !== undefined
        ? this.sheetName(defaultSheetId)
        : this.listSheetRecords().length === 1
          ? this.listSheetRecords()[0]!.name
          : undefined;
    try {
      const parsed = parseRangeAddress(value, defaultSheetName);
      if (parsed.kind !== "cells") {
        return undefined;
      }
      const sheetName = parsed.sheetName ?? defaultSheetName;
      if (!sheetName) {
        return undefined;
      }
      const sheetId = this.requireSheetId(sheetName);
      return {
        start: { sheet: sheetId, row: parsed.start.row, col: parsed.start.col },
        end: { sheet: sheetId, row: parsed.end.row, col: parsed.end.col },
      };
    } catch {
      return undefined;
    }
  }

  simpleCellAddressToString(
    address: HeadlessCellAddress,
    optionsOrContextSheetId: HeadlessAddressFormatOptions | number = {},
  ): string {
    this.assertNotDisposed();
    const includeSheetName =
      typeof optionsOrContextSheetId === "number"
        ? optionsOrContextSheetId !== address.sheet
        : optionsOrContextSheetId.includeSheetName === true;
    return formatQualifiedCellAddress(
      includeSheetName ? this.sheetName(address.sheet) : undefined,
      address.row,
      address.col,
    );
  }

  simpleCellRangeToString(
    range: HeadlessCellRange,
    optionsOrContextSheetId: HeadlessAddressFormatOptions | number = {},
  ): string {
    const includeSheetName =
      typeof optionsOrContextSheetId === "number"
        ? optionsOrContextSheetId !== range.start.sheet
        : optionsOrContextSheetId.includeSheetName === true;
    const sheetName = includeSheetName ? this.sheetName(range.start.sheet) : undefined;
    return formatRangeAddress({
      kind: "cells",
      sheetName,
      start: {
        row: range.start.row,
        col: range.start.col,
        text: formatAddress(range.start.row, range.start.col),
      },
      end: {
        row: range.end.row,
        col: range.end.col,
        text: formatAddress(range.end.row, range.end.col),
      },
    });
  }

  getCellDependents(address: HeadlessAddressLike): HeadlessDependencyRef[] {
    this.flushPendingBatchOps();
    if (!isCellRange(address)) {
      return this.toDependencyRefs(
        this.engine.getDependents(this.sheetName(address.sheet), this.a1(address)).directDependents,
      );
    }
    return this.collectRangeDependencies(
      address,
      (cellAddress) =>
        this.engine.getDependents(this.sheetName(cellAddress.sheet), this.a1(cellAddress))
          .directDependents,
    );
  }

  getCellPrecedents(address: HeadlessAddressLike): HeadlessDependencyRef[] {
    this.flushPendingBatchOps();
    if (!isCellRange(address)) {
      return this.getDirectPrecedentRefs(address);
    }
    return this.collectRangeDependencies(address, (cellAddress) =>
      this.getDirectPrecedentStrings(cellAddress),
    );
  }

  getSheetName(sheetId: number): string | undefined {
    return this.engine.workbook.getSheetById(sheetId)?.name;
  }

  getSheetNames(): string[] {
    return this.listSheetRecords().map((sheet) => sheet.name);
  }

  getSheetId(name: string): number | undefined {
    return this.engine.workbook.getSheet(name)?.id;
  }

  doesSheetExist(name: string): boolean {
    return this.engine.workbook.getSheet(name) !== undefined;
  }

  countSheets(): number {
    return this.listSheetRecords().length;
  }

  getCellType(address: HeadlessCellAddress): HeadlessCellType {
    this.flushPendingBatchOps();
    const cell = this.engine.getCell(this.sheetName(address.sheet), this.a1(address));
    if (this.isCellEmpty(address)) {
      return "EMPTY";
    }
    if (this.isCellPartOfArray(address)) {
      return "ARRAY";
    }
    return cell.formula ? "FORMULA" : "VALUE";
  }

  doesCellHaveSimpleValue(address: HeadlessCellAddress): boolean {
    this.flushPendingBatchOps();
    const cell = this.engine.getCell(this.sheetName(address.sheet), this.a1(address));
    return !cell.formula && !this.isCellEmpty(address);
  }

  doesCellHaveFormula(address: HeadlessCellAddress): boolean {
    this.flushPendingBatchOps();
    return (
      this.engine.getCell(this.sheetName(address.sheet), this.a1(address)).formula !== undefined
    );
  }

  isCellEmpty(address: HeadlessCellAddress): boolean {
    this.flushPendingBatchOps();
    return (
      this.engine.getCellValue(this.sheetName(address.sheet), this.a1(address)).tag ===
      ValueTag.Empty
    );
  }

  isCellPartOfArray(address: HeadlessCellAddress): boolean {
    this.flushPendingBatchOps();
    return this.engine
      .getSpillRanges()
      .some((spill: { sheetName: string; address: string; rows: number; cols: number }) => {
        if (this.requireSheetId(spill.sheetName) !== address.sheet) {
          return false;
        }
        const owner = parseCellAddress(spill.address, spill.sheetName);
        return (
          address.row >= owner.row &&
          address.row < owner.row + spill.rows &&
          address.col >= owner.col &&
          address.col < owner.col + spill.cols
        );
      });
  }

  getCellValueType(address: HeadlessCellAddress): HeadlessCellValueType {
    const value = this.getCellValue(address);
    switch (value.tag) {
      case ValueTag.Number:
        return "NUMBER";
      case ValueTag.String:
        return "STRING";
      case ValueTag.Boolean:
        return "BOOLEAN";
      case ValueTag.Error:
        return "ERROR";
      case ValueTag.Empty:
      default:
        return "EMPTY";
    }
  }

  getCellValueDetailedType(address: HeadlessCellAddress): HeadlessCellValueDetailedType {
    const type = this.getCellValueType(address);
    if (type !== "NUMBER") {
      return type;
    }
    const format = this.getCellValueFormat(address)?.toLowerCase() ?? "";
    if (format.includes("yy") || format.includes("dd")) {
      if (format.includes("h") || format.includes("s")) {
        return "DATETIME";
      }
      return "DATE";
    }
    if (format.includes("h") || format.includes("s")) {
      return "TIME";
    }
    return type;
  }

  getCellValueFormat(address: HeadlessCellAddress): string | undefined {
    this.flushPendingBatchOps();
    const cell = this.engine.getCell(this.sheetName(address.sheet), this.a1(address));
    return cell.format;
  }

  getNamedExpressionValue(name: string, scope?: number): CellValue | CellValue[][] | undefined {
    this.assertReadable();
    const expression = this.namedExpressions.get(makeNamedExpressionKey(name, scope));
    return expression ? this.evaluateNamedExpression(expression) : undefined;
  }

  getNamedExpressionFormula(name: string, scope?: number): string | undefined {
    const expression = this.namedExpressions.get(makeNamedExpressionKey(name, scope));
    if (!expression) {
      return undefined;
    }
    return isFormulaContent(expression.expression) ? expression.expression : undefined;
  }

  getNamedExpression(name: string, scope?: number): HeadlessNamedExpression | undefined {
    const expression = this.namedExpressions.get(makeNamedExpressionKey(name, scope));
    if (!expression) {
      return undefined;
    }
    return {
      name: expression.publicName,
      expression: expression.expression,
      scope: expression.scope,
      options: expression.options ? structuredClone(expression.options) : undefined,
    };
  }

  addNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ): HeadlessChange[] {
    if (!this.isItPossibleToAddNamedExpression(expressionName, expression, scope)) {
      throw new NamedExpressionNameIsAlreadyTakenError(expressionName);
    }
    return this.captureChanges(
      {
        eventName: "namedExpressionAdded",
        payload: {
          name: expressionName.trim(),
          scope,
          changes: [],
        },
      },
      () => {
        this.upsertNamedExpressionInternal(
          { name: expressionName, expression, scope, options },
          { duringInitialization: false },
        );
      },
    );
  }

  changeNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ): HeadlessChange[] {
    if (!this.isItPossibleToChangeNamedExpression(expressionName, expression, scope)) {
      throw new NamedExpressionDoesNotExistError(expressionName);
    }
    return this.captureChanges(undefined, () => {
      this.upsertNamedExpressionInternal(
        { name: expressionName, expression, scope, options },
        { duringInitialization: false },
      );
    });
  }

  removeNamedExpression(expressionName: string, scope?: number): HeadlessChange[] {
    if (!this.isItPossibleToRemoveNamedExpression(expressionName, scope)) {
      throw new NamedExpressionDoesNotExistError(expressionName);
    }
    const existing = this.namedExpressionRecord(expressionName, scope);
    return this.captureChanges(
      {
        eventName: "namedExpressionRemoved",
        payload: {
          name: existing.publicName,
          scope: existing.scope,
          changes: [],
        },
      },
      () => {
        this.namedExpressions.delete(makeNamedExpressionKey(expressionName, scope));
        this.engine.deleteDefinedName(existing.internalName);
      },
    );
  }

  listNamedExpressions(scope?: number): string[] {
    return [...this.namedExpressions.values()]
      .filter((expression) => expression.scope === scope)
      .map((expression) => expression.publicName)
      .toSorted(compareSheetNames);
  }

  getAllNamedExpressionsSerialized(): SerializedHeadlessNamedExpression[] {
    return [...this.namedExpressions.values()]
      .map((expression) => ({
        name: expression.publicName,
        expression: expression.expression,
        scope: expression.scope,
        options: expression.options ? structuredClone(expression.options) : undefined,
      }))
      .toSorted(
        (left, right) =>
          (left.scope ?? -1) - (right.scope ?? -1) || left.name.localeCompare(right.name),
      );
  }

  normalizeFormula(formula: string): string {
    if (!formula.trim().startsWith("=")) {
      throw new NotAFormulaError();
    }
    try {
      return `=${serializeFormula(parseFormula(stripLeadingEquals(formula)))}`;
    } catch (error) {
      throw new HeadlessParseError(this.messageOf(error, `Unable to normalize formula`));
    }
  }

  calculateFormula(formula: string, scope?: number): CellValue | CellValue[][] {
    if (!formula.trim().startsWith("=")) {
      throw new NotAFormulaError();
    }
    try {
      const temporaryWorkbook = new HeadlessWorkbook(this.getConfig());
      const serializedSheets = this.getAllSheetsSerialized();
      Object.keys(serializedSheets).forEach((sheetName) => {
        temporaryWorkbook.engine.createSheet(sheetName);
      });
      this.getAllNamedExpressionsSerialized().forEach((expression) => {
        temporaryWorkbook.upsertNamedExpressionInternal(expression, { duringInitialization: true });
      });
      Object.entries(serializedSheets).forEach(([sheetName, sheet]) => {
        const sheetId = temporaryWorkbook.requireSheetId(sheetName);
        temporaryWorkbook.replaceSheetContentInternal(sheetId, sheet, {
          duringInitialization: true,
        });
      });
      temporaryWorkbook.clearHistoryStacks();
      const scratchSheetName =
        scope !== undefined ? `__HEADLESS_CALC_${scope}__` : "__HEADLESS_CALC__";
      temporaryWorkbook.engine.createSheet(scratchSheetName);
      const scratchSheetId = temporaryWorkbook.requireSheetId(scratchSheetName);
      temporaryWorkbook.applyRawContent(
        scratchSheetName,
        "A1",
        formula.trim().startsWith("=") ? formula : `=${formula}`,
        scratchSheetId,
      );
      const spill = temporaryWorkbook.engine
        .getSpillRanges()
        .find(
          (candidate: { sheetName: string; address: string; rows: number; cols: number }) =>
            candidate.sheetName === scratchSheetName && candidate.address === "A1",
        );
      const value = spill
        ? temporaryWorkbook.getRangeValues({
            start: { sheet: scratchSheetId, row: 0, col: 0 },
            end: { sheet: scratchSheetId, row: spill.rows - 1, col: spill.cols - 1 },
          })
        : temporaryWorkbook.getCellValue({ sheet: scratchSheetId, row: 0, col: 0 });
      temporaryWorkbook.dispose();
      return value;
    } catch (error) {
      throw new HeadlessParseError(this.messageOf(error, "Unable to calculate formula"));
    }
  }

  getNamedExpressionsFromFormula(formula: string): string[] {
    if (!formula.trim().startsWith("=")) {
      throw new NotAFormulaError();
    }
    try {
      const parsed = parseFormula(stripLeadingEquals(formula));
      const names = new Set<string>();
      collectFormulaNameRefs(parsed, names);
      return [...names].toSorted(compareSheetNames);
    } catch (error) {
      throw new HeadlessParseError(this.messageOf(error, "Unable to inspect formula"));
    }
  }

  validateFormula(formula: string): boolean {
    if (!formula.trim().startsWith("=")) {
      return false;
    }
    try {
      parseFormula(stripLeadingEquals(formula));
      return true;
    } catch {
      return false;
    }
  }

  getRegisteredFunctionNames(languageCode?: string): string[] {
    const code = languageCode ?? this.config.language ?? "enGB";
    const language = HeadlessWorkbook.languageRegistry.get(code);
    const functions = [...this.functionSnapshot.values()]
      .filter((binding) => binding.publicName === binding.publicName.toUpperCase())
      .map((binding) => binding.publicName)
      .toSorted(compareSheetNames);
    if (!language?.functions) {
      return functions;
    }
    return functions.map((name) => language.functions?.[name] ?? name);
  }

  getFunctionPlugin(functionId: string): HeadlessFunctionPluginDefinition | undefined {
    const binding = this.functionAliasLookup.get(functionId.trim().toUpperCase());
    if (!binding) {
      return undefined;
    }
    const plugin = HeadlessWorkbook.functionPluginRegistry.get(binding.pluginId);
    return plugin ? clonePluginDefinition(plugin) : undefined;
  }

  getAllFunctionPlugins(): HeadlessFunctionPluginDefinition[] {
    const pluginIds = new Set(
      [...this.functionSnapshot.values()].map((binding) => binding.pluginId),
    );
    return [...pluginIds]
      .map((pluginId) => HeadlessWorkbook.functionPluginRegistry.get(pluginId))
      .filter((plugin): plugin is HeadlessFunctionPluginDefinition => plugin !== undefined)
      .map((plugin) => clonePluginDefinition(plugin));
  }

  numberToDateTime(value: number): HeadlessDateTime | undefined {
    const dateParts = excelSerialToDateParts(value);
    if (!dateParts) {
      return undefined;
    }
    const whole = Math.floor(value);
    const fraction = value - whole;
    const totalSeconds = Math.round(Math.max(0, fraction) * 86_400);
    const hours = Math.floor(totalSeconds / 3_600) % 24;
    const minutes = Math.floor((totalSeconds % 3_600) / 60);
    const seconds = totalSeconds % 60;
    return {
      year: dateParts.year,
      month: dateParts.month,
      day: dateParts.day,
      hours,
      minutes,
      seconds,
    };
  }

  numberToDate(value: number): Omit<HeadlessDateTime, "hours" | "minutes" | "seconds"> | undefined {
    const dateTime = this.numberToDateTime(value);
    if (!dateTime) {
      return undefined;
    }
    const { year, month, day } = dateTime;
    return { year, month, day };
  }

  numberToTime(value: number): Pick<HeadlessDateTime, "hours" | "minutes" | "seconds"> | undefined {
    const dateTime = this.numberToDateTime(value);
    if (!dateTime) {
      return undefined;
    }
    const { hours, minutes, seconds } = dateTime;
    return { hours, minutes, seconds };
  }

  setCellContents(
    address: HeadlessCellAddress,
    content: RawCellContent | HeadlessSheet,
  ): HeadlessChange[] {
    if (!this.isItPossibleToSetCellContents(address, content)) {
      throw new HeadlessOperationError("Cell contents cannot be set");
    }
    return this.captureChanges(undefined, () => {
      if (isHeadlessSheetMatrix(content)) {
        this.flushPendingBatchOps();
        this.applyMatrixContents(address, content);
        return;
      }
      const sheetName = this.sheetName(address.sheet);
      const a1 = this.a1(address);
      if (this.enqueueDeferredBatchLiteral(sheetName, a1, content)) {
        return;
      }
      this.flushPendingBatchOps();
      this.applyRawContent(sheetName, a1, content, address.sheet);
    });
  }

  swapRowIndexes(sheetId: number, rowA: number, rowB: number): HeadlessChange[];
  swapRowIndexes(
    sheetId: number,
    rowMappings: readonly HeadlessAxisSwapMapping[],
  ): HeadlessChange[];
  swapRowIndexes(
    sheetId: number,
    rowAOrMappings: number | readonly HeadlessAxisSwapMapping[],
    rowB?: number,
  ): HeadlessChange[] {
    const mappings = this.normalizeAxisSwapMappings("row", rowAOrMappings, rowB);
    if (!this.isItPossibleToSwapRowIndexes(sheetId, mappings)) {
      throw new HeadlessOperationError("Rows cannot be swapped");
    }
    return this.batch(() => {
      mappings.forEach(([rowA, mappedRowB]) => {
        if (rowA === mappedRowB) {
          return;
        }
        if (rowA < mappedRowB) {
          this.moveRows(sheetId, rowA, 1, mappedRowB);
          this.moveRows(sheetId, mappedRowB - 1, 1, rowA);
        } else {
          this.moveRows(sheetId, rowA, 1, mappedRowB);
          this.moveRows(sheetId, mappedRowB + 1, 1, rowA);
        }
      });
    });
  }

  setRowOrder(sheetId: number, rowOrder: readonly number[]): HeadlessChange[] {
    if (!this.isItPossibleToSetRowOrder(sheetId, rowOrder)) {
      throw new HeadlessOperationError("Row order is invalid");
    }
    const current = rowOrder.toSorted((left, right) => left - right);
    return this.batch(() => {
      rowOrder.forEach((targetOriginalIndex, targetIndex) => {
        const currentIndex = current.indexOf(targetOriginalIndex);
        if (currentIndex === targetIndex) {
          return;
        }
        this.moveRows(sheetId, currentIndex, 1, targetIndex);
        const [moved] = current.splice(currentIndex, 1);
        current.splice(targetIndex, 0, moved!);
      });
    });
  }

  swapColumnIndexes(sheetId: number, columnA: number, columnB: number): HeadlessChange[];
  swapColumnIndexes(
    sheetId: number,
    columnMappings: readonly HeadlessAxisSwapMapping[],
  ): HeadlessChange[];
  swapColumnIndexes(
    sheetId: number,
    columnAOrMappings: number | readonly HeadlessAxisSwapMapping[],
    columnB?: number,
  ): HeadlessChange[] {
    const mappings = this.normalizeAxisSwapMappings("column", columnAOrMappings, columnB);
    if (!this.isItPossibleToSwapColumnIndexes(sheetId, mappings)) {
      throw new HeadlessOperationError("Columns cannot be swapped");
    }
    return this.batch(() => {
      mappings.forEach(([columnA, mappedColumnB]) => {
        if (columnA === mappedColumnB) {
          return;
        }
        if (columnA < mappedColumnB) {
          this.moveColumns(sheetId, columnA, 1, mappedColumnB);
          this.moveColumns(sheetId, mappedColumnB - 1, 1, columnA);
        } else {
          this.moveColumns(sheetId, columnA, 1, mappedColumnB);
          this.moveColumns(sheetId, mappedColumnB + 1, 1, columnA);
        }
      });
    });
  }

  setColumnOrder(sheetId: number, columnOrder: readonly number[]): HeadlessChange[] {
    if (!this.isItPossibleToSetColumnOrder(sheetId, columnOrder)) {
      throw new HeadlessOperationError("Column order is invalid");
    }
    const current = columnOrder.toSorted((left, right) => left - right);
    return this.batch(() => {
      columnOrder.forEach((targetOriginalIndex, targetIndex) => {
        const currentIndex = current.indexOf(targetOriginalIndex);
        if (currentIndex === targetIndex) {
          return;
        }
        this.moveColumns(sheetId, currentIndex, 1, targetIndex);
        const [moved] = current.splice(currentIndex, 1);
        current.splice(targetIndex, 0, moved!);
      });
    });
  }

  addRows(sheetId: number, start: number, count?: number): HeadlessChange[];
  addRows(sheetId: number, ...indexes: readonly HeadlessAxisInterval[]): HeadlessChange[];
  addRows(
    sheetId: number,
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    ...restIntervals: readonly HeadlessAxisInterval[]
  ): HeadlessChange[] {
    const indexes = this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals);
    if (!this.isItPossibleToAddRows(sheetId, ...indexes)) {
      throw new HeadlessOperationError("Rows cannot be added");
    }
    return this.batch(() => {
      indexes.forEach(([start, amount]) => {
        this.engine.insertRows(this.sheetName(sheetId), start, amount);
      });
    });
  }

  removeRows(sheetId: number, start: number, count?: number): HeadlessChange[];
  removeRows(sheetId: number, ...indexes: readonly HeadlessAxisInterval[]): HeadlessChange[];
  removeRows(
    sheetId: number,
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    ...restIntervals: readonly HeadlessAxisInterval[]
  ): HeadlessChange[] {
    const indexes = this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals);
    if (!this.isItPossibleToRemoveRows(sheetId, ...indexes)) {
      throw new HeadlessOperationError("Rows cannot be removed");
    }
    return this.batch(() => {
      indexes
        .toSorted((left, right) => right[0] - left[0])
        .forEach(([start, amount]) => {
          this.engine.deleteRows(this.sheetName(sheetId), start, amount);
        });
    });
  }

  addColumns(sheetId: number, start: number, count?: number): HeadlessChange[];
  addColumns(sheetId: number, ...indexes: readonly HeadlessAxisInterval[]): HeadlessChange[];
  addColumns(
    sheetId: number,
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    ...restIntervals: readonly HeadlessAxisInterval[]
  ): HeadlessChange[] {
    const indexes = this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals);
    if (!this.isItPossibleToAddColumns(sheetId, ...indexes)) {
      throw new HeadlessOperationError("Columns cannot be added");
    }
    return this.batch(() => {
      indexes.forEach(([start, amount]) => {
        this.engine.insertColumns(this.sheetName(sheetId), start, amount);
      });
    });
  }

  removeColumns(sheetId: number, start: number, count?: number): HeadlessChange[];
  removeColumns(sheetId: number, ...indexes: readonly HeadlessAxisInterval[]): HeadlessChange[];
  removeColumns(
    sheetId: number,
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    ...restIntervals: readonly HeadlessAxisInterval[]
  ): HeadlessChange[] {
    const indexes = this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals);
    if (!this.isItPossibleToRemoveColumns(sheetId, ...indexes)) {
      throw new HeadlessOperationError("Columns cannot be removed");
    }
    return this.batch(() => {
      indexes
        .toSorted((left, right) => right[0] - left[0])
        .forEach(([start, amount]) => {
          this.engine.deleteColumns(this.sheetName(sheetId), start, amount);
        });
    });
  }

  moveCells(source: HeadlessCellRange, target: HeadlessCellAddress): HeadlessChange[] {
    if (!this.isItPossibleToMoveCells(source, target)) {
      throw new HeadlessOperationError("Cells cannot be moved");
    }
    const sourceHeight = source.end.row - source.start.row;
    const sourceWidth = source.end.col - source.start.col;
    return this.captureChanges(undefined, () => {
      this.engine.moveRange(sourceRangeRef(this.sheetName(source.start.sheet), source), {
        sheetName: this.sheetName(target.sheet),
        startAddress: formatAddress(target.row, target.col),
        endAddress: formatAddress(target.row + sourceHeight, target.col + sourceWidth),
      });
    });
  }

  moveRows(sheetId: number, start: number, count: number, target: number): HeadlessChange[] {
    if (!this.isItPossibleToMoveRows(sheetId, start, count, target)) {
      throw new HeadlessOperationError("Rows cannot be moved");
    }
    return this.captureChanges(undefined, () => {
      this.engine.moveRows(this.sheetName(sheetId), start, count, target);
    });
  }

  moveColumns(sheetId: number, start: number, count: number, target: number): HeadlessChange[] {
    if (!this.isItPossibleToMoveColumns(sheetId, start, count, target)) {
      throw new HeadlessOperationError("Columns cannot be moved");
    }
    return this.captureChanges(undefined, () => {
      this.engine.moveColumns(this.sheetName(sheetId), start, count, target);
    });
  }

  addSheet(sheetName?: string): string {
    this.assertNotDisposed();
    const name = sheetName?.trim() || this.nextSheetName();
    if (!this.isItPossibleToAddSheet(name)) {
      throw new SheetNameAlreadyTakenError(name);
    }
    const beforeVisibility = this.ensureVisibilityCache();
    const beforeNames = this.ensureNamedExpressionValueCache();
    this.drainTrackedEngineEvents();
    this.engine.createSheet(name);
    const sheetId = this.requireSheetId(name);
    const payload: HeadlessWorkbookDetailedEventMap["sheetAdded"] = { sheetId, sheetName: name };
    if (this.shouldSuppressEvents()) {
      this.queuedEvents.push({ eventName: "sheetAdded", payload });
    } else {
      this.emitter.emitDetailed({ eventName: "sheetAdded", payload });
    }
    const changes = this.computeChangesAfterMutation(beforeVisibility, beforeNames);
    if (!this.shouldSuppressEvents() && changes.length > 0) {
      this.emitter.emitDetailed({ eventName: "valuesUpdated", payload: { changes } });
    }
    return name;
  }

  removeSheet(sheetId: number): HeadlessChange[] {
    if (!this.isItPossibleToRemoveSheet(sheetId)) {
      throw new HeadlessSheetError(`Sheet '${sheetId}' cannot be removed`);
    }
    const sheetName = this.sheetName(sheetId);
    return this.captureChanges(
      {
        eventName: "sheetRemoved",
        payload: {
          sheetId,
          sheetName,
          changes: [],
        },
      },
      () => {
        this.engine.deleteSheet(sheetName);
      },
    );
  }

  clearSheet(sheetId: number): HeadlessChange[] {
    if (!this.isItPossibleToClearSheet(sheetId)) {
      throw new HeadlessSheetError(`Sheet '${sheetId}' cannot be cleared`);
    }
    return this.captureChanges(undefined, () => {
      const dimensions = this.getSheetDimensions(sheetId);
      if (dimensions.width === 0 || dimensions.height === 0) {
        return;
      }
      this.engine.clearRange({
        sheetName: this.sheetName(sheetId),
        startAddress: "A1",
        endAddress: formatAddress(dimensions.height - 1, dimensions.width - 1),
      });
    });
  }

  setSheetContent(sheetId: number, content: HeadlessSheet): HeadlessChange[] {
    if (!this.isItPossibleToReplaceSheetContent(sheetId, content)) {
      throw new HeadlessSheetError(`Sheet '${sheetId}' cannot be replaced`);
    }
    return this.captureChanges(undefined, () => {
      this.replaceSheetContentInternal(sheetId, content, { duringInitialization: false });
    });
  }

  renameSheet(sheetId: number, nextName: string): HeadlessChange[] {
    if (!this.isItPossibleToRenameSheet(sheetId, nextName)) {
      throw new HeadlessSheetError(`Sheet '${sheetId}' cannot be renamed to '${nextName}'`);
    }
    const oldName = this.sheetName(sheetId);
    const newName = nextName.trim();
    return this.captureChanges(
      {
        eventName: "sheetRenamed",
        payload: {
          sheetId,
          oldName,
          newName,
        },
      },
      () => {
        this.engine.renameSheet(oldName, newName);
      },
    );
  }

  isItPossibleToSetCellContents(
    address: HeadlessCellAddress,
    content?: RawCellContent | HeadlessSheet,
  ): boolean;
  isItPossibleToSetCellContents(range: HeadlessCellRange): boolean;
  isItPossibleToSetCellContents(
    addressOrRange: HeadlessAddressLike,
    content?: RawCellContent | HeadlessSheet,
  ): boolean {
    this.assertNotDisposed();
    if (isCellRange(addressOrRange)) {
      assertRange(addressOrRange);
      this.sheetRecord(addressOrRange.start.sheet);
      return (
        addressOrRange.end.row < (this.config.maxRows ?? MAX_ROWS) &&
        addressOrRange.end.col < (this.config.maxColumns ?? MAX_COLS)
      );
    }
    this.sheetRecord(addressOrRange.sheet);
    assertRowAndColumn(addressOrRange.row, "address.row");
    assertRowAndColumn(addressOrRange.col, "address.col");
    if (content === undefined) {
      return (
        addressOrRange.row < (this.config.maxRows ?? MAX_ROWS) &&
        addressOrRange.col < (this.config.maxColumns ?? MAX_COLS)
      );
    }
    if (Array.isArray(content)) {
      if (!content.every((row) => Array.isArray(row))) {
        throw new HeadlessArgumentError("Content matrix must be a two-dimensional array");
      }
      const height = content.length;
      const width = Math.max(0, ...content.map((row) => row.length));
      return (
        addressOrRange.row + height <= (this.config.maxRows ?? MAX_ROWS) &&
        addressOrRange.col + width <= (this.config.maxColumns ?? MAX_COLS)
      );
    }
    return (
      addressOrRange.row < (this.config.maxRows ?? MAX_ROWS) &&
      addressOrRange.col < (this.config.maxColumns ?? MAX_COLS)
    );
  }

  isItPossibleToSwapRowIndexes(sheetId: number, rowA: number, rowB: number): boolean;
  isItPossibleToSwapRowIndexes(
    sheetId: number,
    rowMappings: readonly HeadlessAxisSwapMapping[],
  ): boolean;
  isItPossibleToSwapRowIndexes(
    sheetId: number,
    rowAOrMappings: number | readonly HeadlessAxisSwapMapping[],
    rowB?: number,
  ): boolean {
    this.sheetRecord(sheetId);
    const mappings = this.normalizeAxisSwapMappings("row", rowAOrMappings, rowB);
    return mappings.every(([rowA, mappedRowB]) => {
      assertRowAndColumn(rowA, "rowA");
      assertRowAndColumn(mappedRowB, "rowB");
      return true;
    });
  }

  isItPossibleToSetRowOrder(sheetId: number, rowOrder: readonly number[]): boolean {
    this.sheetRecord(sheetId);
    if (
      new Set(rowOrder).size !== rowOrder.length ||
      rowOrder.some((value) => !Number.isInteger(value) || value < 0)
    ) {
      return false;
    }
    return true;
  }

  isItPossibleToSwapColumnIndexes(sheetId: number, columnA: number, columnB: number): boolean;
  isItPossibleToSwapColumnIndexes(
    sheetId: number,
    columnMappings: readonly HeadlessAxisSwapMapping[],
  ): boolean;
  isItPossibleToSwapColumnIndexes(
    sheetId: number,
    columnAOrMappings: number | readonly HeadlessAxisSwapMapping[],
    columnB?: number,
  ): boolean {
    this.sheetRecord(sheetId);
    const mappings = this.normalizeAxisSwapMappings("column", columnAOrMappings, columnB);
    return mappings.every(([columnA, mappedColumnB]) => {
      assertRowAndColumn(columnA, "columnA");
      assertRowAndColumn(mappedColumnB, "columnB");
      return true;
    });
  }

  isItPossibleToSetColumnOrder(sheetId: number, columnOrder: readonly number[]): boolean {
    this.sheetRecord(sheetId);
    if (
      new Set(columnOrder).size !== columnOrder.length ||
      columnOrder.some((value) => !Number.isInteger(value) || value < 0)
    ) {
      return false;
    }
    return true;
  }

  isItPossibleToAddRows(sheetId: number, start: number, count?: number): boolean;
  isItPossibleToAddRows(sheetId: number, ...indexes: readonly HeadlessAxisInterval[]): boolean;
  isItPossibleToAddRows(
    sheetId: number,
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    ...restIntervals: readonly HeadlessAxisInterval[]
  ): boolean {
    this.sheetRecord(sheetId);
    return this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(
      ([start, count]) => {
        assertRowAndColumn(start, "start");
        assertRowAndColumn(count, "count");
        return count > 0 && start + count <= (this.config.maxRows ?? MAX_ROWS);
      },
    );
  }

  isItPossibleToRemoveRows(sheetId: number, start: number, count?: number): boolean;
  isItPossibleToRemoveRows(sheetId: number, ...indexes: readonly HeadlessAxisInterval[]): boolean;
  isItPossibleToRemoveRows(
    sheetId: number,
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    ...restIntervals: readonly HeadlessAxisInterval[]
  ): boolean {
    this.sheetRecord(sheetId);
    return this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(
      ([start, count]) => {
        assertRowAndColumn(start, "start");
        assertRowAndColumn(count, "count");
        return count > 0;
      },
    );
  }

  isItPossibleToAddColumns(sheetId: number, start: number, count?: number): boolean;
  isItPossibleToAddColumns(sheetId: number, ...indexes: readonly HeadlessAxisInterval[]): boolean;
  isItPossibleToAddColumns(
    sheetId: number,
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    ...restIntervals: readonly HeadlessAxisInterval[]
  ): boolean {
    this.sheetRecord(sheetId);
    return this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(
      ([start, count]) => {
        assertRowAndColumn(start, "start");
        assertRowAndColumn(count, "count");
        return count > 0 && start + count <= (this.config.maxColumns ?? MAX_COLS);
      },
    );
  }

  isItPossibleToRemoveColumns(sheetId: number, start: number, count?: number): boolean;
  isItPossibleToRemoveColumns(
    sheetId: number,
    ...indexes: readonly HeadlessAxisInterval[]
  ): boolean;
  isItPossibleToRemoveColumns(
    sheetId: number,
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    ...restIntervals: readonly HeadlessAxisInterval[]
  ): boolean {
    this.sheetRecord(sheetId);
    return this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(
      ([start, count]) => {
        assertRowAndColumn(start, "start");
        assertRowAndColumn(count, "count");
        return count > 0;
      },
    );
  }

  isItPossibleToMoveCells(source: HeadlessCellRange, target: HeadlessCellAddress): boolean {
    assertRange(source);
    assertRowAndColumn(target.sheet, "target.sheet");
    assertRowAndColumn(target.row, "target.row");
    assertRowAndColumn(target.col, "target.col");
    return source.start.sheet === target.sheet;
  }

  isItPossibleToMoveRows(sheetId: number, start: number, count: number, target: number): boolean {
    this.sheetRecord(sheetId);
    assertRowAndColumn(start, "start");
    assertRowAndColumn(count, "count");
    assertRowAndColumn(target, "target");
    return count > 0;
  }

  isItPossibleToMoveColumns(
    sheetId: number,
    start: number,
    count: number,
    target: number,
  ): boolean {
    this.sheetRecord(sheetId);
    assertRowAndColumn(start, "start");
    assertRowAndColumn(count, "count");
    assertRowAndColumn(target, "target");
    return count > 0;
  }

  isItPossibleToAddSheet(sheetName: string): boolean {
    const trimmed = sheetName.trim();
    if (trimmed.length === 0) {
      throw new HeadlessArgumentError("Sheet name must be non-empty");
    }
    return !this.doesSheetExist(trimmed);
  }

  isItPossibleToRemoveSheet(sheetId: number): boolean {
    return this.engine.workbook.getSheetById(sheetId) !== undefined;
  }

  isItPossibleToClearSheet(sheetId: number): boolean {
    return this.engine.workbook.getSheetById(sheetId) !== undefined;
  }

  isItPossibleToReplaceSheetContent(sheetId: number, content: HeadlessSheet): boolean {
    this.sheetRecord(sheetId);
    if (!content.every((row) => Array.isArray(row))) {
      throw new HeadlessArgumentError("Sheet content must be a two-dimensional array");
    }
    const height = content.length;
    const width = Math.max(0, ...content.map((row) => row.length));
    return (
      height <= (this.config.maxRows ?? MAX_ROWS) && width <= (this.config.maxColumns ?? MAX_COLS)
    );
  }

  isItPossibleToRenameSheet(sheetId: number, nextName: string): boolean {
    this.sheetRecord(sheetId);
    const trimmed = nextName.trim();
    if (trimmed.length === 0) {
      throw new HeadlessArgumentError("Sheet name must be non-empty");
    }
    const existing = this.engine.workbook.getSheet(trimmed);
    return !existing || existing.id === sheetId;
  }

  isItPossibleToAddNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
  ): boolean {
    this.validateNamedExpression(expressionName, expression, scope);
    return !this.namedExpressions.has(makeNamedExpressionKey(expressionName, scope));
  }

  isItPossibleToChangeNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
  ): boolean {
    this.validateNamedExpression(expressionName, expression, scope);
    return this.namedExpressions.has(makeNamedExpressionKey(expressionName, scope));
  }

  isItPossibleToRemoveNamedExpression(expressionName: string, scope?: number): boolean {
    return this.namedExpressions.has(makeNamedExpressionKey(expressionName, scope));
  }

  destroy(): void {
    this.dispose();
  }

  dispose(): void {
    if (this.disposed) {
      return;
    }
    this.disposed = true;
    this.unsubscribeEngineEvents?.();
    this.unsubscribeEngineEvents = null;
    this.emitter.clear();
    this.clearFunctionBindings();
    this.clipboard = null;
    this.visibilityCache = null;
    this.namedExpressionValueCache = null;
    this.queuedEvents = [];
    this.trackedEngineEvents = [];
    this.namedExpressions.clear();
  }

  private attachEngineEventTracking(): void {
    this.unsubscribeEngineEvents?.();
    this.trackedEngineEvents = [];
    this.unsubscribeEngineEvents = this.engine.subscribe((event) => {
      this.trackedEngineEvents.push(cloneTrackedEngineEvent(event));
    });
  }

  private drainTrackedEngineEvents(): EngineEvent[] {
    const events = this.trackedEngineEvents;
    this.trackedEngineEvents = [];
    return events;
  }

  private primeChangeTrackingCaches(): void {
    this.visibilityCache = this.captureVisibilitySnapshot();
    this.namedExpressionValueCache = this.captureNamedExpressionValueSnapshot();
    this.drainTrackedEngineEvents();
  }

  private ensureVisibilityCache(): VisibilitySnapshot {
    if (!this.visibilityCache) {
      this.visibilityCache = this.captureVisibilitySnapshot();
    }
    return this.visibilityCache;
  }

  private ensureNamedExpressionValueCache(): NamedExpressionValueSnapshot {
    if (!this.namedExpressionValueCache) {
      this.namedExpressionValueCache = this.captureNamedExpressionValueSnapshot();
    }
    return this.namedExpressionValueCache;
  }

  private flushPendingBatchOps(): void {
    if (this.pendingBatchOps.length === 0) {
      return;
    }
    const ops = this.pendingBatchOps;
    const potentialNewCells = this.pendingBatchPotentialNewCells;
    this.pendingBatchOps = [];
    this.pendingBatchPotentialNewCells = 0;
    this.engine.applyOps(ops, {
      captureUndo: true,
      potentialNewCells: potentialNewCells > 0 ? potentialNewCells : undefined,
    });
  }

  private enqueueDeferredBatchLiteral(
    sheetName: string,
    address: string,
    content: RawCellContent,
  ): boolean {
    if (
      this.batchDepth === 0 ||
      !isDeferredBatchLiteralContent(content) ||
      isFormulaContent(content)
    ) {
      return false;
    }
    if (content === null) {
      this.pendingBatchOps.push({ kind: "clearCell", sheetName, address });
      return true;
    }
    this.pendingBatchOps.push({ kind: "setCellValue", sheetName, address, value: content });
    this.pendingBatchPotentialNewCells += 1;
    return true;
  }

  private prepareReadableState(): void {
    this.assertNotDisposed();
    this.flushPendingBatchOps();
  }

  private assertNotDisposed(): void {
    if (this.disposed) {
      throw new HeadlessOperationError("Workbook has been disposed");
    }
  }

  private assertReadable(): void {
    this.prepareReadableState();
    if (this.evaluationSuspended) {
      throw new HeadlessEvaluationSuspendedError();
    }
  }

  private sheetRecord(sheetId: number) {
    const sheet = this.engine.workbook.getSheetById(sheetId);
    if (!sheet) {
      throw new NoSheetWithIdError(sheetId);
    }
    return sheet;
  }

  private sheetName(sheetId: number): string {
    return this.sheetRecord(sheetId).name;
  }

  private requireSheetId(name: string): number {
    const sheetId = this.getSheetId(name);
    if (sheetId === undefined) {
      throw new NoSheetWithNameError(name);
    }
    return sheetId;
  }

  private a1(address: Pick<HeadlessCellAddress, "row" | "col">): string {
    return formatAddress(address.row, address.col);
  }

  private rangeRef(range: HeadlessCellRange): CellRangeRef {
    assertRange(range);
    return sourceRangeRef(this.sheetName(range.start.sheet), range);
  }

  private getDirectPrecedentStrings(address: HeadlessCellAddress): string[] {
    const precedents = new Set<string>(
      this.engine.getDependencies(this.sheetName(address.sheet), this.a1(address)).directPrecedents,
    );
    const formula = this.getCellFormula(address);
    if (formula) {
      this.getNamedExpressionsFromFormula(formula).forEach((name) => {
        precedents.add(name);
      });
    }
    return [...precedents];
  }

  private getDirectPrecedentRefs(address: HeadlessCellAddress): HeadlessDependencyRef[] {
    return this.toDependencyRefs(this.getDirectPrecedentStrings(address));
  }

  private listSheetRecords() {
    return [...this.engine.workbook.sheetsByName.values()].toSorted(
      (left, right) => left.order - right.order || left.name.localeCompare(right.name),
    );
  }

  private getDenseRange<Value>(
    range: HeadlessCellRange,
    read: (address: HeadlessCellAddress) => Value,
  ): Value[][] {
    assertRange(range);
    const height = range.end.row - range.start.row + 1;
    const width = range.end.col - range.start.col + 1;
    return Array.from({ length: height }, (_row, rowOffset) =>
      Array.from({ length: width }, (_column, colOffset) =>
        read({
          sheet: range.start.sheet,
          row: range.start.row + rowOffset,
          col: range.start.col + colOffset,
        }),
      ),
    );
  }

  private captureVisibilitySnapshot(): VisibilitySnapshot {
    const snapshot = new Map<number, SheetStateSnapshot>();
    this.listSheetRecords().forEach((sheet) => {
      const cells = new Map<string, CellValue>();
      sheet.grid.forEachCellEntry((_cellIndex: number, row: number, col: number) => {
        const address = formatAddress(row, col);
        const value = this.readStoredCellValue(sheet.name, row, col);
        if (value.tag === ValueTag.Empty) {
          return;
        }
        cells.set(address, value);
      });
      snapshot.set(sheet.id, {
        sheetId: sheet.id,
        sheetName: sheet.name,
        order: sheet.order,
        cells,
      });
    });
    return snapshot;
  }

  private readStoredCellValue(sheetName: string, row: number, col: number): CellValue {
    const sheet = this.engine.workbook.getSheet(sheetName);
    if (!sheet) {
      return emptyValue();
    }
    const cellIndex = sheet.grid.get(row, col);
    if (cellIndex === -1) {
      return emptyValue();
    }
    return cloneCellValue(
      this.engine.workbook.cellStore.getValue(cellIndex, (id) => this.engine.strings.get(id)),
    );
  }

  private captureNamedExpressionValueSnapshot(): NamedExpressionValueSnapshot {
    const snapshot = new Map<string, CellValue | CellValue[][]>();
    [...this.namedExpressions.values()].forEach((expression) => {
      snapshot.set(
        makeNamedExpressionKey(expression.publicName, expression.scope),
        cloneNamedExpressionValue(this.evaluateNamedExpression(expression)),
      );
    });
    return snapshot;
  }

  private collectTrackedCellRefs(events: readonly EngineEvent[]): TrackedCellRef[] | null {
    if (events.length === 0) {
      return [];
    }
    const refs = new Map<string, TrackedCellRef>();
    for (const event of events) {
      if (
        event.invalidation === "full" ||
        event.invalidatedRanges.length > 0 ||
        event.invalidatedRows.length > 0 ||
        event.invalidatedColumns.length > 0
      ) {
        return null;
      }
      for (let index = 0; index < event.changedCellIndices.length; index += 1) {
        const cellIndex = event.changedCellIndices[index]!;
        const qualifiedAddress = this.engine.workbook.getQualifiedAddress(cellIndex);
        if (qualifiedAddress.length === 0 || qualifiedAddress.startsWith("!")) {
          return null;
        }
        const separator = qualifiedAddress.indexOf("!");
        if (separator <= 0 || separator === qualifiedAddress.length - 1) {
          return null;
        }
        const sheetName = qualifiedAddress.slice(0, separator);
        const address = qualifiedAddress.slice(separator + 1);
        const sheetId = this.getSheetId(sheetName);
        if (sheetId === undefined) {
          return null;
        }
        const parsed = parseCellAddress(address, sheetName);
        refs.set(`${sheetId}:${address}`, {
          sheetId,
          sheetName,
          address,
          row: parsed.row,
          col: parsed.col,
        });
      }
    }
    return [...refs.values()];
  }

  private computeCellChanges(
    beforeVisibility: VisibilitySnapshot,
    afterVisibility: VisibilitySnapshot,
  ): HeadlessChange[] {
    const cellChanges: HeadlessChange[] = [];
    afterVisibility.forEach((afterSheet, sheetId) => {
      const beforeSheet = beforeVisibility.get(sheetId);
      const addresses = new Set<string>([
        ...(beforeSheet?.cells.keys() ?? []),
        ...afterSheet.cells.keys(),
      ]);
      [...addresses].toSorted(compareSheetNames).forEach((address) => {
        const beforeValue = beforeSheet?.cells.get(address) ?? emptyValue();
        const afterValue = afterSheet.cells.get(address) ?? emptyValue();
        if (valuesEqual(beforeValue, afterValue)) {
          return;
        }
        const parsed = parseCellAddress(address, afterSheet.sheetName);
        cellChanges.push({
          kind: "cell",
          address: { sheet: sheetId, row: parsed.row, col: parsed.col },
          sheetName: afterSheet.sheetName,
          a1: address,
          newValue: cloneCellValue(afterValue),
        });
      });
    });
    return cellChanges.toSorted(compareHeadlessCellChanges(this.listSheetRecords()));
  }

  private computeCellChangesFromTrackedEvents(
    beforeVisibility: VisibilitySnapshot,
    events: readonly EngineEvent[],
  ): { changes: HeadlessChange[]; nextVisibility: VisibilitySnapshot } | null {
    const refs = this.collectTrackedCellRefs(events);
    if (refs === null) {
      return null;
    }

    const nextVisibility = new Map(beforeVisibility);
    const mutableSheets = new Set<number>();
    const ensureMutableSheet = (ref: TrackedCellRef): SheetStateSnapshot => {
      const existing =
        nextVisibility.get(ref.sheetId) ??
        ({
          sheetId: ref.sheetId,
          sheetName: ref.sheetName,
          order: this.sheetRecord(ref.sheetId).order,
          cells: new Map<string, CellValue>(),
        } satisfies SheetStateSnapshot);
      if (!mutableSheets.has(ref.sheetId)) {
        nextVisibility.set(ref.sheetId, {
          sheetId: existing.sheetId,
          sheetName: existing.sheetName,
          order: existing.order,
          cells: new Map(
            [...existing.cells.entries()].map(([address, value]) => [
              address,
              cloneCellValue(value),
            ]),
          ),
        });
        mutableSheets.add(ref.sheetId);
      }
      return nextVisibility.get(ref.sheetId)!;
    };

    const changes: HeadlessChange[] = [];
    refs.forEach((ref) => {
      const beforeValue = beforeVisibility.get(ref.sheetId)?.cells.get(ref.address) ?? emptyValue();
      const afterValue = this.readStoredCellValue(ref.sheetName, ref.row, ref.col);
      if (valuesEqual(beforeValue, afterValue)) {
        return;
      }
      const sheet = ensureMutableSheet(ref);
      if (afterValue.tag === ValueTag.Empty) {
        sheet.cells.delete(ref.address);
      } else {
        sheet.cells.set(ref.address, afterValue);
      }
      changes.push({
        kind: "cell",
        address: { sheet: ref.sheetId, row: ref.row, col: ref.col },
        sheetName: ref.sheetName,
        a1: ref.address,
        newValue: cloneCellValue(afterValue),
      });
    });

    return {
      changes: changes.toSorted(compareHeadlessCellChanges(this.listSheetRecords())),
      nextVisibility,
    };
  }

  private computeNamedExpressionChanges(
    beforeNames: NamedExpressionValueSnapshot,
    afterNames: NamedExpressionValueSnapshot,
  ): HeadlessChange[] {
    const namedExpressionChanges: HeadlessChange[] = [];
    afterNames.forEach((afterValue, key) => {
      const beforeValue = beforeNames.get(key);
      if (matrixValuesEqual(beforeValue, afterValue)) {
        return;
      }
      const expression = this.namedExpressions.get(key);
      if (!expression) {
        return;
      }
      namedExpressionChanges.push({
        kind: "named-expression",
        name: expression.publicName,
        scope: expression.scope,
        newValue: cloneNamedExpressionValue(afterValue),
      });
    });
    return namedExpressionChanges.toSorted(compareHeadlessNamedExpressionChanges);
  }

  private computeChangesAfterMutation(
    beforeVisibility: VisibilitySnapshot,
    beforeNames: NamedExpressionValueSnapshot,
  ): HeadlessChange[] {
    const afterNames = this.captureNamedExpressionValueSnapshot();
    const fastPath = this.computeCellChangesFromTrackedEvents(
      beforeVisibility,
      this.drainTrackedEngineEvents(),
    );
    let cellChanges: HeadlessChange[];
    if (fastPath) {
      cellChanges = fastPath.changes;
      this.visibilityCache = fastPath.nextVisibility;
    } else {
      const afterVisibility = this.captureVisibilitySnapshot();
      cellChanges = this.computeCellChanges(beforeVisibility, afterVisibility);
      this.visibilityCache = afterVisibility;
    }
    this.namedExpressionValueCache = afterNames;
    return [...cellChanges, ...this.computeNamedExpressionChanges(beforeNames, afterNames)];
  }

  private captureChanges(
    semanticEvent: QueuedEvent | undefined,
    mutate: () => void,
  ): HeadlessChange[] {
    this.assertNotDisposed();
    if (semanticEvent !== undefined) {
      this.flushPendingBatchOps();
    }
    if (this.shouldSuppressEvents()) {
      try {
        mutate();
      } catch (error) {
        if (error instanceof Error && HEADLESS_PUBLIC_ERROR_NAMES.has(error.name)) {
          throw error;
        }
        throw new HeadlessOperationError(this.messageOf(error, "Mutation failed"));
      }
      if (semanticEvent) {
        this.queuedEvents.push(semanticEvent);
      }
      return [];
    }
    const beforeVisibility = this.ensureVisibilityCache();
    const beforeNames = this.ensureNamedExpressionValueCache();
    this.drainTrackedEngineEvents();
    try {
      mutate();
    } catch (error) {
      if (error instanceof Error && HEADLESS_PUBLIC_ERROR_NAMES.has(error.name)) {
        throw error;
      }
      throw new HeadlessOperationError(this.messageOf(error, "Mutation failed"));
    }
    const changes =
      semanticEvent === undefined
        ? this.computeChangesAfterMutation(beforeVisibility, beforeNames)
        : (() => {
            const afterVisibility = this.captureVisibilitySnapshot();
            const afterNames = this.captureNamedExpressionValueSnapshot();
            this.visibilityCache = afterVisibility;
            this.namedExpressionValueCache = afterNames;
            return [
              ...this.computeCellChanges(beforeVisibility, afterVisibility),
              ...this.computeNamedExpressionChanges(beforeNames, afterNames),
            ];
          })();
    if (semanticEvent) {
      const event = withEventChanges(semanticEvent, changes);
      if (this.shouldSuppressEvents()) {
        this.queuedEvents.push(event);
      } else {
        this.emitter.emitDetailed(event);
      }
    }
    if (!this.shouldSuppressEvents() && changes.length > 0) {
      this.emitter.emitDetailed({ eventName: "valuesUpdated", payload: { changes } });
    }
    return changes;
  }

  private shouldSuppressEvents(): boolean {
    return this.batchDepth > 0 || this.evaluationSuspended;
  }

  private flushQueuedEvents(): void {
    const events = [...this.queuedEvents];
    this.queuedEvents.length = 0;
    events.forEach((event) => {
      this.emitter.emitDetailed(event);
    });
  }

  private getUndoStack(): HistoryRecord[] {
    const stack = Reflect.get(this.engine, "undoStack");
    if (!isHistoryRecordArray(stack)) {
      return [];
    }
    return stack;
  }

  private getRedoStack(): HistoryRecord[] {
    const stack = Reflect.get(this.engine, "redoStack");
    if (!isHistoryRecordArray(stack)) {
      return [];
    }
    return stack;
  }

  private clearHistoryStacks(): void {
    this.getUndoStack().length = 0;
    this.getRedoStack().length = 0;
  }

  private mergeUndoHistory(startIndex: number): void {
    const undoStack = this.getUndoStack();
    if (undoStack.length - startIndex <= 1) {
      return;
    }
    const entries = undoStack.splice(startIndex);
    const merged: HistoryRecord = {
      forward: {
        ops: entries.flatMap((entry) => structuredClone(entry.forward.ops)),
        potentialNewCells: sumNumbers(entries.map((entry) => entry.forward.potentialNewCells)),
      },
      inverse: {
        ops: entries.toReversed().flatMap((entry) => structuredClone(entry.inverse.ops)),
        potentialNewCells: sumNumbers(entries.map((entry) => entry.inverse.potentialNewCells)),
      },
    };
    undoStack.push(merged);
  }

  private nextSheetName(): string {
    let index = 1;
    while (this.doesSheetExist(`Sheet${index}`)) {
      index += 1;
    }
    return `Sheet${index}`;
  }

  private buildNullMatrixForRange(range: HeadlessCellRange): RawCellContent[][] {
    const height = range.end.row - range.start.row + 1;
    const width = range.end.col - range.start.col + 1;
    return Array.from({ length: height }, () => Array.from({ length: width }, () => null));
  }

  private applySerializedMatrix(
    targetLeftCorner: HeadlessCellAddress,
    serialized: RawCellContent[][],
    sourceAnchor: HeadlessCellAddress,
  ): void {
    this.flushPendingBatchOps();
    if (matrixContainsFormulaContent(serialized)) {
      serialized.forEach((row, rowOffset) => {
        row.forEach((raw, columnOffset) => {
          const destination = {
            sheet: targetLeftCorner.sheet,
            row: targetLeftCorner.row + rowOffset,
            col: targetLeftCorner.col + columnOffset,
          };
          let nextValue = raw;
          if (typeof raw === "string" && raw.startsWith("=")) {
            nextValue = `=${translateFormulaReferences(
              raw.slice(1),
              destination.row - (sourceAnchor.row + rowOffset),
              destination.col - (sourceAnchor.col + columnOffset),
            )}`;
          }
          this.applyRawContent(
            this.sheetName(destination.sheet),
            this.a1(destination),
            nextValue,
            destination.sheet,
          );
        });
      });
      return;
    }

    const sheetName = this.sheetName(targetLeftCorner.sheet);
    const { ops, potentialNewCells } = buildMatrixMutationPlan({
      target: targetLeftCorner,
      targetSheetName: sheetName,
      content: serialized,
      rewriteFormula: (formula, destination, rowOffset, columnOffset) =>
        this.rewriteFormulaForStorage(
          translateFormulaReferences(
            stripLeadingEquals(formula),
            destination.row - (sourceAnchor.row + rowOffset),
            destination.col - (sourceAnchor.col + columnOffset),
          ),
          destination.sheet,
        ),
    });
    if (ops.length === 0) {
      return;
    }
    this.engine.applyOps(ops, { potentialNewCells });
  }

  private applyMatrixContents(
    address: HeadlessCellAddress,
    content: HeadlessSheet,
    options: {
      captureUndo?: boolean;
      skipNulls?: boolean;
    } = {},
  ): void {
    this.flushPendingBatchOps();
    if (matrixContainsFormulaContent(content)) {
      content.forEach((row, rowOffset) => {
        row.forEach((raw, columnOffset) => {
          if (raw === null && options.skipNulls) {
            return;
          }
          this.applyRawContent(
            this.sheetName(address.sheet),
            formatAddress(address.row + rowOffset, address.col + columnOffset),
            raw,
            address.sheet,
          );
        });
      });
      return;
    }

    const sheetName = this.sheetName(address.sheet);
    const { ops, potentialNewCells } = buildMatrixMutationPlan({
      target: address,
      targetSheetName: sheetName,
      content,
      skipNulls: options.skipNulls,
      rewriteFormula: (formula, destination) =>
        this.rewriteFormulaForStorage(stripLeadingEquals(formula), destination.sheet),
    });
    if (ops.length === 0) {
      return;
    }
    this.engine.applyOps(ops, {
      captureUndo: options.captureUndo,
      potentialNewCells,
    });
  }

  private replaceSheetContentInternal(
    sheetId: number,
    content: HeadlessSheet,
    options: { duringInitialization: boolean },
  ): void {
    const sheetName = this.sheetName(sheetId);
    const undoStackStart = options.duringInitialization ? 0 : this.getUndoStack().length;
    const dimensions = this.getSheetDimensions(sheetId);
    if (dimensions.width > 0 && dimensions.height > 0) {
      this.engine.clearRange({
        sheetName,
        startAddress: "A1",
        endAddress: formatAddress(dimensions.height - 1, dimensions.width - 1),
      });
    }
    this.applyMatrixContents({ sheet: sheetId, row: 0, col: 0 }, content, {
      captureUndo: !options.duringInitialization,
      skipNulls: true,
    });
    if (options.duringInitialization) {
      this.clearHistoryStacks();
      return;
    }
    this.mergeUndoHistory(undoStackStart);
  }

  private applyRawContent(
    sheetName: string,
    address: string,
    content: RawCellContent,
    ownerSheetId: number,
  ): void {
    if (content === null) {
      this.engine.clearCell(sheetName, address);
      return;
    }
    if (typeof content === "boolean" || typeof content === "number") {
      this.engine.setCellValue(sheetName, address, content);
      return;
    }
    if (typeof content === "string" && content.trim().startsWith("=")) {
      this.engine.setCellFormula(
        sheetName,
        address,
        this.rewriteFormulaForStorage(stripLeadingEquals(content), ownerSheetId),
      );
      return;
    }
    this.engine.setCellValue(sheetName, address, content);
  }

  private captureFunctionRegistry(): void {
    const allowedPluginIds =
      this.config.functionPlugins && this.config.functionPlugins.length > 0
        ? new Set(this.config.functionPlugins.map((plugin) => plugin.id))
        : undefined;
    HeadlessWorkbook.functionPluginRegistry.forEach((plugin) => {
      if (allowedPluginIds && !allowedPluginIds.has(plugin.id)) {
        return;
      }
      Object.keys(plugin.implementedFunctions).forEach((functionId) => {
        const normalized = functionId.trim().toUpperCase();
        const internalName = `__BILIG_HEADLESS_FN_${this.workbookId}_${normalized}`;
        const implementation = plugin.functions?.[normalized];
        const binding: InternalFunctionBinding = {
          pluginId: plugin.id,
          publicName: normalized,
          internalName,
          implementation,
        };
        this.functionSnapshot.set(normalized, binding);
        this.functionAliasLookup.set(normalized, binding);
        this.internalFunctionLookup.set(internalName, binding);
        if (implementation) {
          globalCustomFunctions.set(internalName, implementation);
        }
      });
      Object.entries(plugin.aliases ?? {}).forEach(([alias, target]) => {
        const binding = this.functionSnapshot.get(target.trim().toUpperCase());
        if (!binding) {
          return;
        }
        this.functionAliasLookup.set(alias.trim().toUpperCase(), binding);
      });
    });
  }

  private clearFunctionBindings(): void {
    this.internalFunctionLookup.forEach((_binding, internalName) => {
      globalCustomFunctions.delete(internalName);
    });
    this.functionSnapshot.clear();
    this.functionAliasLookup.clear();
    this.internalFunctionLookup.clear();
  }

  private rebuildWithConfig(nextConfig: HeadlessConfig): void {
    validateHeadlessConfig(nextConfig);
    const serializedSheets = this.getAllSheetsSerialized();
    Object.entries(serializedSheets).forEach(([sheetName, sheet]) => {
      validateSheetWithinLimits(sheetName, sheet, nextConfig);
    });
    const serializedNamedExpressions = this.getAllNamedExpressionsSerialized();
    const suspended = this.evaluationSuspended;
    const clipboard = this.clipboard
      ? {
          sourceAnchor: { ...this.clipboard.sourceAnchor },
          serialized: this.clipboard.serialized.map((row) => [...row]),
          values: this.clipboard.values.map((row) => row.map((value) => cloneCellValue(value))),
        }
      : null;

    this.clearFunctionBindings();
    this.namedExpressions.clear();
    this.engine = new SpreadsheetEngine({
      workbookName: "Workbook",
      useColumnIndex: this.config.useColumnIndex,
    });
    this.attachEngineEventTracking();
    this.config = cloneConfig(nextConfig);
    this.captureFunctionRegistry();

    Object.keys(serializedSheets).forEach((sheetName) => {
      this.engine.createSheet(sheetName);
    });
    serializedNamedExpressions.forEach((expression) => {
      this.upsertNamedExpressionInternal(expression, { duringInitialization: true });
    });
    Object.entries(serializedSheets).forEach(([sheetName, sheet]) => {
      const sheetId = this.requireSheetId(sheetName);
      this.replaceSheetContentInternal(sheetId, sheet, { duringInitialization: true });
    });
    this.clearHistoryStacks();
    this.primeChangeTrackingCaches();
    this.clipboard = clipboard;
    if (suspended) {
      this.suspendedVisibility = this.ensureVisibilityCache();
      this.suspendedNamedValues = this.ensureNamedExpressionValueCache();
    }
  }

  private normalizeAxisIntervals(
    startOrInterval: number | HeadlessAxisInterval,
    countOrInterval?: number | HeadlessAxisInterval,
    restIntervals: readonly HeadlessAxisInterval[] = [],
  ): Array<[number, number]> {
    if (typeof startOrInterval === "number") {
      if (Array.isArray(countOrInterval)) {
        throw new HeadlessArgumentError("Axis interval count must be a number");
      }
      const resolvedCount = typeof countOrInterval === "number" ? countOrInterval : 1;
      return [[startOrInterval, resolvedCount]];
    }
    if (typeof countOrInterval === "number") {
      throw new HeadlessArgumentError("Axis interval count is only valid with a numeric start");
    }
    return [startOrInterval, ...(countOrInterval ? [countOrInterval] : []), ...restIntervals].map(
      ([start, count]) => [start, count ?? 1] as [number, number],
    );
  }

  private normalizeAxisSwapMappings(
    label: "row" | "column",
    startOrMappings: number | readonly HeadlessAxisSwapMapping[],
    end?: number,
  ): HeadlessAxisSwapMapping[] {
    if (typeof startOrMappings === "number") {
      if (end === undefined) {
        throw new HeadlessArgumentError(`${label} swap requires two indexes`);
      }
      return [[startOrMappings, end]];
    }
    return [...startOrMappings];
  }

  private collectRangeDependencies(
    range: HeadlessCellRange,
    readDependencies: (address: HeadlessCellAddress) => readonly string[],
  ): HeadlessDependencyRef[] {
    assertRange(range);
    const seen = new Set<string>();
    const collected: HeadlessDependencyRef[] = [];
    this.getDenseRange(range, (address) => address).forEach((row) => {
      row.forEach((address) => {
        this.toDependencyRefs(readDependencies(address)).forEach((dependency) => {
          const key =
            dependency.kind === "cell"
              ? `cell:${dependency.address.sheet}:${dependency.address.row}:${dependency.address.col}`
              : dependency.kind === "range"
                ? `range:${dependency.range.start.sheet}:${dependency.range.start.row}:${dependency.range.start.col}:${dependency.range.end.row}:${dependency.range.end.col}`
                : `name:${dependency.name}`;
          if (seen.has(key)) {
            return;
          }
          seen.add(key);
          collected.push(dependency);
        });
      });
    });
    return collected;
  }

  private rewriteFormulaForStorage(formula: string, ownerSheetId: number): string {
    try {
      const transformed = transformFormulaNode(
        parseFormula(stripLeadingEquals(formula)),
        (node) => {
          if (node.kind === "NameRef") {
            return this.rewriteNameRefForStorage(node, ownerSheetId);
          }
          if (node.kind === "CallExpr") {
            return this.rewriteCallForStorage(node);
          }
          return node;
        },
      );
      return serializeFormula(transformed);
    } catch (error) {
      throw new HeadlessParseError(this.messageOf(error, "Unable to store formula"));
    }
  }

  private restorePublicFormula(formula: string, ownerSheetId: number): string {
    const transformed = transformFormulaNode(parseFormula(formula), (node) => {
      if (node.kind === "NameRef") {
        return this.rewriteNameRefForPublic(node, ownerSheetId);
      }
      if (node.kind === "CallExpr") {
        return this.rewriteCallForPublic(node);
      }
      return node;
    });
    return serializeFormula(transformed);
  }

  private rewriteNameRefForStorage(node: NameRefNode, ownerSheetId: number): FormulaNode {
    const scoped = this.namedExpressions.get(makeNamedExpressionKey(node.name, ownerSheetId));
    if (scoped) {
      return { ...node, name: scoped.internalName };
    }
    const workbookScoped = this.namedExpressions.get(makeNamedExpressionKey(node.name));
    if (workbookScoped) {
      return { ...node, name: workbookScoped.internalName };
    }
    return node;
  }

  private rewriteNameRefForPublic(node: NameRefNode, ownerSheetId: number): FormulaNode {
    const exact = [...this.namedExpressions.values()].find(
      (expression) => expression.internalName === node.name && expression.scope === ownerSheetId,
    );
    if (exact) {
      return { ...node, name: exact.publicName };
    }
    return node;
  }

  private rewriteCallForStorage(node: CallExprNode): FormulaNode {
    const binding = this.functionAliasLookup.get(node.callee.trim().toUpperCase());
    if (!binding) {
      return node;
    }
    return { ...node, callee: binding.internalName };
  }

  private rewriteCallForPublic(node: CallExprNode): FormulaNode {
    const binding = this.internalFunctionLookup.get(node.callee.trim().toUpperCase());
    if (!binding) {
      return node;
    }
    return { ...node, callee: binding.publicName };
  }

  private validateNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
  ): void {
    const trimmed = expressionName.trim();
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed) || isCellReferenceText(trimmed)) {
      throw new NamedExpressionNameIsInvalidError(expressionName);
    }
    if (scope !== undefined) {
      this.sheetRecord(scope);
    }
    if (isFormulaContent(expression)) {
      try {
        const parsed = parseFormula(stripLeadingEquals(expression));
        if (formulaHasRelativeReferences(parsed)) {
          throw new NoRelativeAddressesAllowedError();
        }
      } catch (error) {
        if (error instanceof NoRelativeAddressesAllowedError) {
          throw error;
        }
        throw new UnableToParseError({
          expressionName,
          reason: this.messageOf(error, `Invalid named expression formula for '${expressionName}'`),
        });
      }
    }
  }

  private upsertNamedExpressionInternal(
    expression: SerializedHeadlessNamedExpression,
    options: { duringInitialization: boolean },
  ): void {
    this.validateNamedExpression(expression.name, expression.expression, expression.scope);
    const trimmed = expression.name.trim();
    const internalName =
      expression.scope === undefined ? trimmed : makeInternalScopedName(expression.scope, trimmed);
    const record: InternalNamedExpression = {
      publicName: trimmed,
      normalizedName: normalizeName(trimmed),
      internalName,
      scope: expression.scope,
      expression: expression.expression,
      options: expression.options ? structuredClone(expression.options) : undefined,
    };
    this.namedExpressions.set(makeNamedExpressionKey(trimmed, expression.scope), record);
    this.engine.setDefinedName(
      internalName,
      this.toDefinedNameSnapshot(record.expression, record.scope),
    );
    if (options.duringInitialization) {
      this.clearHistoryStacks();
    }
  }

  private toDefinedNameSnapshot(
    expression: RawCellContent,
    scope?: number,
  ): WorkbookDefinedNameValueSnapshot {
    if (expression === null || typeof expression === "number" || typeof expression === "boolean") {
      return expression;
    }
    if (typeof expression === "string" && expression.trim().startsWith("=")) {
      return {
        kind: "formula",
        formula: `=${this.rewriteFormulaForStorage(stripLeadingEquals(expression), scope ?? this.listSheetRecords()[0]?.id ?? 1)}`,
      };
    }
    return expression;
  }

  private namedExpressionRecord(name: string, scope?: number): InternalNamedExpression {
    const direct = this.namedExpressions.get(makeNamedExpressionKey(name, scope));
    if (direct) {
      return direct;
    }
    throw new NamedExpressionDoesNotExistError(name);
  }

  private evaluateNamedExpression(expression: InternalNamedExpression): CellValue | CellValue[][] {
    const raw = expression.expression;
    if (raw === null || typeof raw === "number" || typeof raw === "boolean") {
      return scalarValueFromLiteral(raw);
    }
    if (typeof raw === "string" && !raw.trim().startsWith("=")) {
      return scalarValueFromLiteral(raw);
    }
    return this.calculateFormula(raw, expression.scope);
  }

  private cellSnapshotToRawContent(cell: CellSnapshot, ownerSheetId: number): RawCellContent {
    if (cell.formula) {
      return `=${this.restorePublicFormula(cell.formula, ownerSheetId)}`;
    }
    if (cell.input !== undefined) {
      return cell.input;
    }
    switch (cell.value.tag) {
      case ValueTag.Empty:
      case ValueTag.Error:
        return null;
      case ValueTag.Number:
        return cell.value.value;
      case ValueTag.Boolean:
        return cell.value.value;
      case ValueTag.String:
        return cell.value.value;
    }
  }

  private toDependencyRefs(values: readonly string[]): HeadlessDependencyRef[] {
    return values.map((value) => {
      try {
        const parsedCell = parseCellAddress(value);
        return {
          kind: "cell",
          address: {
            sheet: this.requireSheetId(parsedCell.sheetName ?? this.listSheetRecords()[0]!.name),
            row: parsedCell.row,
            col: parsedCell.col,
          },
        } satisfies HeadlessDependencyRef;
      } catch {
        try {
          const parsedRange = parseRangeAddress(value);
          if (parsedRange.kind === "cells") {
            return {
              kind: "range",
              range: {
                start: {
                  sheet: this.requireSheetId(
                    parsedRange.sheetName ?? this.listSheetRecords()[0]!.name,
                  ),
                  row: parsedRange.start.row,
                  col: parsedRange.start.col,
                },
                end: {
                  sheet: this.requireSheetId(
                    parsedRange.sheetName ?? this.listSheetRecords()[0]!.name,
                  ),
                  row: parsedRange.end.row,
                  col: parsedRange.end.col,
                },
              },
            } satisfies HeadlessDependencyRef;
          }
        } catch {
          return { kind: "name", name: value } satisfies HeadlessDependencyRef;
        }
      }
      return { kind: "name", name: value } satisfies HeadlessDependencyRef;
    });
  }

  private messageOf(error: unknown, fallback: string): string {
    return error instanceof Error && error.message.length > 0 ? error.message : fallback;
  }
}

function cloneNamedExpressionValue(value: CellValue | CellValue[][]): CellValue | CellValue[][] {
  if (!Array.isArray(value)) {
    return cloneCellValue(value);
  }
  return value.map((row) => row.map((cell) => cloneCellValue(cell)));
}

function compareHeadlessNamedExpressionChanges(
  left: HeadlessChange,
  right: HeadlessChange,
): number {
  if (left.kind !== "named-expression" || right.kind !== "named-expression") {
    return 0;
  }
  return (left.scope ?? -1) - (right.scope ?? -1) || left.name.localeCompare(right.name);
}

function compareHeadlessCellChanges(
  sheets: readonly { id: number; order: number }[],
): (left: HeadlessChange, right: HeadlessChange) => number {
  const orderBySheet = new Map(sheets.map((sheet) => [sheet.id, sheet.order]));
  return (left, right) => {
    if (left.kind !== "cell" || right.kind !== "cell") {
      return 0;
    }
    return (
      (orderBySheet.get(left.address.sheet) ?? 0) - (orderBySheet.get(right.address.sheet) ?? 0) ||
      left.address.row - right.address.row ||
      left.address.col - right.address.col
    );
  };
}

function sourceRangeRef(sheetName: string, range: HeadlessCellRange): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(range.start.row, range.start.col),
    endAddress: formatAddress(range.end.row, range.end.col),
  };
}

function sumNumbers(values: Array<number | undefined>): number | undefined {
  const filtered = values.filter((value): value is number => typeof value === "number");
  if (filtered.length === 0) {
    return undefined;
  }
  return filtered.reduce((sum, value) => sum + value, 0);
}
