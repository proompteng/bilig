import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getExternalLookupFunction } from "../external-function-adapter.js";
import type { ArrayValue, EvaluationResult } from "../runtime-values.js";
import { createLookupArrayShapeBuiltins } from "./lookup-array-shape-builtins.js";
import { createLookupCriteriaBuiltins } from "./lookup-criteria-builtins.js";
import { createLookupDatabaseBuiltins } from "./lookup-database-builtins.js";
import { createLookupFinancialBuiltins } from "./lookup-financial-builtins.js";
import { createLookupHypothesisBuiltins } from "./lookup-hypothesis-builtins.js";
import { createLookupMatrixBuiltins } from "./lookup-matrix-builtins.js";
import { createLookupOrderStatisticsBuiltins } from "./lookup-order-statistics-builtins.js";
import { createLookupReferenceBuiltins } from "./lookup-reference-builtins.js";
import { createLookupRegressionBuiltins } from "./lookup-regression-builtins.js";
import { createLookupSortFilterBuiltins } from "./lookup-sort-filter-builtins.js";

export interface RangeBuiltinArgument {
  kind: "range";
  values: CellValue[];
  refKind: "cells" | "rows" | "cols";
  rows: number;
  cols: number;
}

export type LookupBuiltinArgument = CellValue | RangeBuiltinArgument;
export type LookupBuiltin = (...args: LookupBuiltinArgument[]) => EvaluationResult;

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function isError(
  value: LookupBuiltinArgument | undefined,
): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return value !== undefined && !isRangeArg(value) && value.tag === ValueTag.Error;
}

function isRangeArg(value: LookupBuiltinArgument | undefined): value is RangeBuiltinArgument {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "range";
}

function isCriteriaOperator(value: string): value is CriteriaOperator {
  return (
    value === "=" ||
    value === "<>" ||
    value === ">" ||
    value === ">=" ||
    value === "<" ||
    value === "<="
  );
}

function findFirstNonRange(
  values: readonly (RangeBuiltinArgument | CellValue)[],
): CellValue | undefined {
  for (const value of values) {
    if (!isRangeArg(value)) {
      return value;
    }
  }
  return undefined;
}

function areRangeArgs(
  values: readonly (RangeBuiltinArgument | CellValue)[],
): values is RangeBuiltinArgument[] {
  return values.every((value) => isRangeArg(value));
}

function toNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    case ValueTag.String:
    case ValueTag.Error:
      return undefined;
    default:
      return undefined;
  }
}

function toInteger(value: CellValue): number | undefined {
  const numeric = toNumber(value);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  return Math.trunc(numeric);
}

function toBoolean(value: CellValue): boolean | undefined {
  switch (value.tag) {
    case ValueTag.Boolean:
      return value.value;
    case ValueTag.Number:
      return value.value !== 0;
    case ValueTag.Empty:
      return false;
    case ValueTag.String:
    case ValueTag.Error:
      return undefined;
    default:
      return undefined;
  }
}

function toStringValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return value.value;
    case ValueTag.Error:
      return "";
  }
}

function compareScalars(left: CellValue, right: CellValue): number | undefined {
  if (
    (left.tag === ValueTag.String || left.tag === ValueTag.Empty) &&
    (right.tag === ValueTag.String || right.tag === ValueTag.Empty)
  ) {
    const normalizedLeft = toStringValue(left).toUpperCase();
    const normalizedRight = toStringValue(right).toUpperCase();
    if (normalizedLeft === normalizedRight) {
      return 0;
    }
    return normalizedLeft < normalizedRight ? -1 : 1;
  }

  const leftNum = toNumber(left);
  const rightNum = toNumber(right);
  if (leftNum === undefined || rightNum === undefined) {
    return undefined;
  }
  if (leftNum === rightNum) {
    return 0;
  }
  return leftNum < rightNum ? -1 : 1;
}

function requireCellVector(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg)) {
    return errorValue(ErrorCode.Value);
  }
  if (arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  if (arg.rows !== 1 && arg.cols !== 1) {
    return errorValue(ErrorCode.NA);
  }
  return arg;
}

function requireCellRange(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg) || arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  return arg;
}

function getRangeValue(range: RangeBuiltinArgument, row: number, col: number): CellValue {
  const index = row * range.cols + col;
  return range.values[index] ?? { tag: ValueTag.Empty };
}

function arrayResult(values: CellValue[], rows: number, cols: number): ArrayValue {
  return { kind: "array", values, rows, cols };
}

function collectNumericSeries(
  arg: LookupBuiltinArgument,
  mode: "lenient" | "strict",
): number[] | CellValue {
  const values: number[] = [];
  const cells = isRangeArg(arg) ? arg.values : [arg];
  if (isRangeArg(arg) && arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  for (const cell of cells) {
    if (cell.tag === ValueTag.Error) {
      return cell;
    }
    if (cell.tag === ValueTag.Number) {
      values.push(cell.value);
      continue;
    }
    if (mode === "strict") {
      return errorValue(ErrorCode.Value);
    }
  }
  return values;
}

function numericAggregateCandidate(value: CellValue): number | undefined {
  return value.tag === ValueTag.Number ? value.value : undefined;
}

function toCellRange(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg)) {
    return { kind: "range", values: [arg], refKind: "cells", rows: 1, cols: 1 };
  }
  if (arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  return arg;
}

function toNumericMatrix(arg: LookupBuiltinArgument): number[][] | CellValue {
  const range = toCellRange(arg);
  if (!isRangeArg(range)) {
    return range;
  }
  const matrix: number[][] = [];
  for (let row = 0; row < range.rows; row += 1) {
    const rowValues: number[] = [];
    for (let col = 0; col < range.cols; col += 1) {
      const numeric = toNumber(getRangeValue(range, row, col));
      if (numeric === undefined) {
        return errorValue(ErrorCode.Value);
      }
      rowValues.push(numeric);
    }
    matrix.push(rowValues);
  }
  return matrix;
}

function flattenNumbers(arg: LookupBuiltinArgument): number[] | CellValue {
  if (!isRangeArg(arg)) {
    const numeric = toNumber(arg);
    return numeric === undefined ? errorValue(ErrorCode.Value) : [numeric];
  }
  const values: number[] = [];
  for (const value of arg.values) {
    const numeric = toNumber(value);
    if (numeric === undefined) {
      return errorValue(ErrorCode.Value);
    }
    values.push(numeric);
  }
  return values;
}

function pickRangeRow(range: RangeBuiltinArgument, row: number): CellValue[] {
  const values: CellValue[] = [];
  for (let col = 0; col < range.cols; col += 1) {
    values.push(getRangeValue(range, row, col));
  }
  return values;
}

function hasWildcardPattern(pattern: string): boolean {
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index];
    if (char === "~") {
      index += 1;
      continue;
    }
    if (char === "*" || char === "?") {
      return true;
    }
  }
  return false;
}

function wildcardPatternToRegExp(pattern: string): RegExp {
  let source = "^";
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index];
    if (char === undefined) {
      continue;
    }
    if (char === "~") {
      const escaped = pattern[index + 1];
      if (escaped !== undefined) {
        source += escapeRegexFragment(escaped);
        index += 1;
        continue;
      }
      source += escapeRegexFragment(char);
      continue;
    }
    if (char === "*") {
      source += ".*";
      continue;
    }
    if (char === "?") {
      source += ".";
      continue;
    }
    source += escapeRegexFragment(char);
  }
  source += "$";
  return new RegExp(source, "i");
}

function escapeRegexFragment(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function unescapeCriteriaPattern(pattern: string): string {
  let unescaped = "";
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index];
    if (char === undefined) {
      continue;
    }
    if (char === "~") {
      const escaped = pattern[index + 1];
      if (escaped !== undefined) {
        unescaped += escaped;
        index += 1;
        continue;
      }
    }
    unescaped += char;
  }
  return unescaped;
}

function firstLookupError(args: readonly LookupBuiltinArgument[]): CellValue | undefined {
  return args.find((arg) => isError(arg));
}

const externalLookupBuiltinNames = ["FILTERXML", "STOCKHISTORY"] as const;

function createExternalLookupBuiltin(name: string): LookupBuiltin {
  return (...args) => {
    const existingError = firstLookupError(args);
    if (existingError) {
      return existingError;
    }
    const external = getExternalLookupFunction(name);
    return external ? external(...args) : errorValue(ErrorCode.Blocked);
  };
}

const externalLookupBuiltins = Object.fromEntries(
  externalLookupBuiltinNames.map((name) => [name, createExternalLookupBuiltin(name)]),
) as Record<string, LookupBuiltin>;

const lookupRegressionBuiltins = createLookupRegressionBuiltins({
  errorValue,
  numberResult,
  isRangeArg,
  toNumber,
  toBoolean,
  flattenNumbers,
});

const lookupOrderStatisticsBuiltins = createLookupOrderStatisticsBuiltins({
  errorValue,
  numberResult,
  arrayResult,
  requireCellRange,
  isError,
  isRangeArg,
  toNumber,
  toInteger,
  flattenNumbers,
});

const lookupFinancialBuiltins = createLookupFinancialBuiltins({
  errorValue,
  numberResult,
  isRangeArg,
  toNumber,
  collectNumericSeries,
});

const lookupHypothesisBuiltins = createLookupHypothesisBuiltins({
  errorValue,
  isRangeArg,
  toNumber,
  toNumericMatrix,
});

const lookupMatrixBuiltins = createLookupMatrixBuiltins({
  errorValue,
  numberResult,
  arrayResult,
  isRangeArg,
  requireCellRange,
  findFirstNonRange,
  areRangeArgs,
  toNumber,
  toNumericMatrix,
  flattenNumbers,
});

const lookupSortFilterBuiltins = createLookupSortFilterBuiltins({
  errorValue,
  arrayResult,
  isError,
  isRangeArg,
  toBoolean,
  toInteger,
  requireCellRange,
  toCellRange,
  compareScalars,
  getRangeValue,
  pickRangeRow,
});

const lookupDatabaseBuiltins = createLookupDatabaseBuiltins({
  errorValue,
  numberResult,
  isError,
  isRangeArg,
  toNumber,
  toStringValue,
  requireCellRange,
  getRangeValue,
  matchesCriteria,
});

const lookupCriteriaBuiltins = createLookupCriteriaBuiltins({
  errorValue,
  numberResult,
  isError,
  isRangeArg,
  toNumber,
  requireCellRange,
  matchesCriteria,
  numericAggregateCandidate,
});

const lookupReferenceBuiltins = createLookupReferenceBuiltins({
  errorValue,
  numberResult,
  isError,
  isRangeArg,
  toBoolean,
  toInteger,
  requireCellVector,
  toCellRange,
  compareScalars,
  getRangeValue,
});

const lookupArrayShapeBuiltins = createLookupArrayShapeBuiltins({
  errorValue,
  arrayResult,
  isError,
  isRangeArg,
  toBoolean,
  toInteger,
  requireCellRange,
  toCellRange,
  getRangeValue,
  findFirstNonRange,
  areRangeArgs,
  pickRangeRow,
});

export const lookupBuiltins: Record<string, LookupBuiltin> = {
  ...lookupArrayShapeBuiltins,
  ...lookupReferenceBuiltins,
  ...lookupCriteriaBuiltins,
  ...lookupDatabaseBuiltins,
  ...lookupFinancialBuiltins,
  ...lookupHypothesisBuiltins,
  ...lookupRegressionBuiltins,
  ...lookupOrderStatisticsBuiltins,
  ...lookupMatrixBuiltins,
  ...lookupSortFilterBuiltins,
  ...externalLookupBuiltins,
};

type CriteriaOperator = "=" | "<>" | ">" | ">=" | "<" | "<=";

function matchesCriteria(value: CellValue, criteria: CellValue): boolean {
  if (isError(value)) {
    return false;
  }
  let { operator, operand } = parseCriteria(criteria);
  if (
    operand.tag === ValueTag.String &&
    (operator === "=" || operator === "<>") &&
    hasWildcardPattern(operand.value)
  ) {
    const matches = wildcardPatternToRegExp(operand.value).test(toStringValue(value));
    return operator === "=" ? matches : !matches;
  }
  if (operand.tag === ValueTag.String && operand.value.includes("~")) {
    operand = {
      tag: ValueTag.String,
      value: unescapeCriteriaPattern(operand.value),
      stringId: operand.stringId,
    };
  }
  const comparison = compareScalars(value, operand);
  if (comparison === undefined) {
    return false;
  }
  switch (operator) {
    case "=":
      return comparison === 0;
    case "<>":
      return comparison !== 0;
    case ">":
      return comparison > 0;
    case ">=":
      return comparison >= 0;
    case "<":
      return comparison < 0;
    case "<=":
      return comparison <= 0;
  }
}

function parseCriteria(criteria: CellValue): { operator: CriteriaOperator; operand: CellValue } {
  if (criteria.tag !== ValueTag.String) {
    return { operator: "=", operand: criteria };
  }

  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(criteria.value);
  if (!match) {
    return { operator: "=", operand: criteria };
  }

  const operator = match[1] ?? "=";
  return {
    operator: isCriteriaOperator(operator) ? operator : "=",
    operand: parseCriteriaOperand(match[2] ?? ""),
  };
}

function parseCriteriaOperand(raw: string): CellValue {
  const trimmed = raw.trim();
  if (trimmed === "") {
    return { tag: ValueTag.String, value: "", stringId: 0 };
  }
  const upper = trimmed.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") {
    return { tag: ValueTag.Boolean, value: upper === "TRUE" };
  }
  const numeric = Number(trimmed);
  if (Number.isFinite(numeric)) {
    return { tag: ValueTag.Number, value: numeric };
  }
  return { tag: ValueTag.String, value: trimmed, stringId: 0 };
}

export function getLookupBuiltin(name: string): LookupBuiltin | undefined {
  const upper = name.toUpperCase();
  if (upper === "USE.THE.COUNTIF") {
    return lookupBuiltins["COUNTIF"];
  }
  return lookupBuiltins[upper] ?? getExternalLookupFunction(name);
}
