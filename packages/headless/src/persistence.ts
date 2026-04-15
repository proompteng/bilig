import type { LiteralInput } from "@bilig/protocol";
import { WorkPaper } from "./work-paper.js";
import type {
  RawCellContent,
  SerializedWorkPaperNamedExpression,
  WorkPaperChooseAddressMappingPolicy,
  WorkPaperConfig,
  WorkPaperContextValue,
  WorkPaperSheet,
} from "./work-paper-types.js";

export const WORK_PAPER_DOCUMENT_FORMAT = "bilig.headless.work-paper.document.v1" as const;

export const PERSISTABLE_WORK_PAPER_CONFIG_KEYS = [
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
  "precisionEpsilon",
  "precisionRounding",
  "smartRounding",
  "thousandSeparator",
  "timeFormats",
  "useArrayArithmetic",
  "useColumnIndex",
  "useStats",
  "undoLimit",
  "useRegularExpressions",
  "useWildcards",
] as const satisfies readonly (keyof WorkPaperConfig)[];

type PersistableWorkPaperConfigKey = (typeof PERSISTABLE_WORK_PAPER_CONFIG_KEYS)[number];

export type PersistableWorkPaperConfig = Pick<WorkPaperConfig, PersistableWorkPaperConfigKey>;

export interface PersistedWorkPaperNamedExpression {
  name: string;
  expression: RawCellContent;
  scopeSheetName?: string;
  options?: Record<string, string | number | boolean>;
}

export interface PersistedWorkPaperSheet {
  name: string;
  content: WorkPaperSheet;
}

export interface PersistedWorkPaperDocument {
  format: typeof WORK_PAPER_DOCUMENT_FORMAT;
  sheets: PersistedWorkPaperSheet[];
  namedExpressions: PersistedWorkPaperNamedExpression[];
  config?: PersistableWorkPaperConfig;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "boolean" ||
    typeof value === "number" ||
    typeof value === "string"
  );
}

function isRawCellContent(value: unknown): value is RawCellContent {
  return isLiteralInput(value) || typeof value === "string";
}

function isWorkPaperSheet(value: unknown): value is WorkPaperSheet {
  return (
    Array.isArray(value) &&
    value.every((row) => Array.isArray(row) && row.every((cell) => isRawCellContent(cell)))
  );
}

function isNamedExpressionOptions(
  value: unknown,
): value is Record<string, string | number | boolean> {
  return (
    isRecord(value) &&
    Object.values(value).every(
      (entry) =>
        typeof entry === "string" || typeof entry === "number" || typeof entry === "boolean",
    )
  );
}

function isWorkPaperContextValue(value: unknown): value is WorkPaperContextValue {
  return (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    (Array.isArray(value) && value.every((entry) => isWorkPaperContextValue(entry))) ||
    (isRecord(value) && Object.values(value).every((entry) => isWorkPaperContextValue(entry)))
  );
}

function isChooseAddressMappingPolicy(
  value: unknown,
): value is WorkPaperChooseAddressMappingPolicy {
  return isRecord(value) && (value["mode"] === "dense" || value["mode"] === "sparse");
}

function isPersistableWorkPaperConfig(value: unknown): value is PersistableWorkPaperConfig {
  if (!isRecord(value)) {
    return false;
  }
  return Object.entries(value).every(([key, entry]) => {
    if (!(PERSISTABLE_WORK_PAPER_CONFIG_KEYS as readonly string[]).includes(key)) {
      return false;
    }
    if (key === "chooseAddressMappingPolicy") {
      return isChooseAddressMappingPolicy(entry);
    }
    if (key === "currencySymbol" || key === "dateFormats" || key === "timeFormats") {
      return Array.isArray(entry) && entry.every((item) => typeof item === "string");
    }
    if (key === "nullDate") {
      return (
        isRecord(entry) &&
        typeof entry["year"] === "number" &&
        typeof entry["month"] === "number" &&
        typeof entry["day"] === "number"
      );
    }
    if (key === "context") {
      return isWorkPaperContextValue(entry);
    }
    return typeof entry === "string" || typeof entry === "number" || typeof entry === "boolean";
  });
}

function isPersistedWorkPaperNamedExpression(
  value: unknown,
): value is PersistedWorkPaperNamedExpression {
  return (
    isRecord(value) &&
    typeof value["name"] === "string" &&
    isRawCellContent(value["expression"]) &&
    (value["scopeSheetName"] === undefined || typeof value["scopeSheetName"] === "string") &&
    (value["options"] === undefined || isNamedExpressionOptions(value["options"]))
  );
}

function isPersistedWorkPaperSheet(value: unknown): value is PersistedWorkPaperSheet {
  return isRecord(value) && typeof value["name"] === "string" && isWorkPaperSheet(value["content"]);
}

/**
 * Checks whether a value matches the persisted WorkPaper document format.
 */
export function isPersistedWorkPaperDocument(value: unknown): value is PersistedWorkPaperDocument {
  return (
    isRecord(value) &&
    value["format"] === WORK_PAPER_DOCUMENT_FORMAT &&
    Array.isArray(value["sheets"]) &&
    value["sheets"].every((sheet) => isPersistedWorkPaperSheet(sheet)) &&
    Array.isArray(value["namedExpressions"]) &&
    value["namedExpressions"].every((expression) =>
      isPersistedWorkPaperNamedExpression(expression),
    ) &&
    (value["config"] === undefined || isPersistableWorkPaperConfig(value["config"]))
  );
}

function assertPersistedWorkPaperDocument(
  value: unknown,
): asserts value is PersistedWorkPaperDocument {
  if (!isPersistedWorkPaperDocument(value)) {
    throw new Error("Invalid persisted WorkPaper document");
  }
}

function setPersistableWorkPaperConfigValue<Key extends PersistableWorkPaperConfigKey>(
  target: PersistableWorkPaperConfig,
  key: Key,
  value: PersistableWorkPaperConfig[Key],
): void {
  target[key] = structuredClone(value);
}

/**
 * Clones the documented JSON-safe subset of WorkPaper configuration values.
 */
export function pickPersistableWorkPaperConfig(
  config: WorkPaperConfig,
): PersistableWorkPaperConfig {
  const picked: PersistableWorkPaperConfig = {};
  for (const key of PERSISTABLE_WORK_PAPER_CONFIG_KEYS) {
    const value = config[key];
    if (value !== undefined) {
      setPersistableWorkPaperConfigValue(picked, key, value);
    }
  }
  return picked;
}

/**
 * Exports sheets, named expressions, and optional config from a WorkPaper.
 */
export function exportWorkPaperDocument(
  workbook: WorkPaper,
  options: { includeConfig?: boolean } = {},
): PersistedWorkPaperDocument {
  const { includeConfig = true } = options;
  const sheets = workbook.getSheetNames().map((name) => {
    const sheetId = workbook.getSheetId(name);
    if (sheetId === undefined) {
      throw new Error(`Missing sheet id for ${name}`);
    }
    return {
      name,
      content: workbook.getSheetSerialized(sheetId),
    } satisfies PersistedWorkPaperSheet;
  });
  const namedExpressions = workbook
    .getAllNamedExpressionsSerialized()
    .map((expression) => serializeNamedExpression(workbook, expression));
  const document: PersistedWorkPaperDocument = {
    format: WORK_PAPER_DOCUMENT_FORMAT,
    sheets,
    namedExpressions,
  };
  if (includeConfig) {
    document.config = pickPersistableWorkPaperConfig(workbook.getConfig());
  }
  return document;
}

/**
 * Creates a WorkPaper instance from a validated persisted document.
 */
export function createWorkPaperFromDocument(document: PersistedWorkPaperDocument): WorkPaper {
  assertPersistedWorkPaperDocument(document);
  const workbook = WorkPaper.buildEmpty(document.config ?? {});
  document.sheets.forEach((sheet) => {
    workbook.addSheet(sheet.name);
  });
  document.namedExpressions.forEach((expression) => {
    const scope = expression.scopeSheetName
      ? workbook.getSheetId(expression.scopeSheetName)
      : undefined;
    if (expression.scopeSheetName && scope === undefined) {
      throw new Error(`Missing scoped sheet ${expression.scopeSheetName}`);
    }
    workbook.addNamedExpression(expression.name, expression.expression, scope, expression.options);
  });
  document.sheets.forEach((sheet) => {
    const sheetId = workbook.getSheetId(sheet.name);
    if (sheetId === undefined) {
      throw new Error(`Missing restored sheet ${sheet.name}`);
    }
    workbook.setSheetContent(sheetId, sheet.content);
  });
  workbook.clearUndoStack();
  workbook.clearRedoStack();
  return workbook;
}

/**
 * Serializes a validated WorkPaper document to JSON.
 */
export function serializeWorkPaperDocument(document: PersistedWorkPaperDocument): string {
  assertPersistedWorkPaperDocument(document);
  return JSON.stringify(document);
}

/**
 * Parses and validates a WorkPaper document from JSON.
 */
export function parseWorkPaperDocument(json: string): PersistedWorkPaperDocument {
  const parsed = JSON.parse(json) as unknown;
  assertPersistedWorkPaperDocument(parsed);
  return parsed;
}

function serializeNamedExpression(
  workbook: WorkPaper,
  expression: SerializedWorkPaperNamedExpression,
): PersistedWorkPaperNamedExpression {
  if (expression.scope === undefined) {
    return {
      name: expression.name,
      expression: expression.expression,
      options: expression.options ? structuredClone(expression.options) : undefined,
    };
  }
  const scopeSheetName = workbook.getSheetName(expression.scope);
  if (!scopeSheetName) {
    throw new Error(`Missing scope sheet for named expression ${expression.name}`);
  }
  return {
    name: expression.name,
    expression: expression.expression,
    scopeSheetName,
    options: expression.options ? structuredClone(expression.options) : undefined,
  };
}
