import { ErrorCode, ValueTag, formatErrorCode, type CellValue } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { indexToColumn, parseCellAddress, parseRangeAddress } from "./addressing.js";
import { getBuiltin, hasBuiltin } from "./builtins.js";
import { getLookupBuiltin, type RangeBuiltinArgument } from "./builtins/lookup.js";
import { evaluateGroupBy, evaluatePivotBy, type MatrixValue } from "./group-pivot-evaluator.js";
import {
  isArrayValue,
  scalarFromEvaluationResult,
  type ArrayValue,
  type EvaluationResult,
  type RangeLikeValue,
} from "./runtime-values.js";
import { rewriteSpecialCall } from "./special-call-rewrites.js";

export interface EvaluationContext {
  sheetName: string;
  currentAddress?: string;
  resolveCell: (sheetName: string, address: string) => CellValue;
  resolveRange: (
    sheetName: string,
    start: string,
    end: string,
    refKind: "cells" | "rows" | "cols",
  ) => CellValue[];
  resolveName?: (name: string) => CellValue;
  resolveFormula?: (sheetName: string, address: string) => string | undefined;
  resolvePivotData?: (request: {
    dataField: string;
    sheetName: string;
    address: string;
    filters: ReadonlyArray<{ field: string; item: CellValue }>;
  }) => CellValue | undefined;
  resolveMultipleOperations?: (request: {
    formulaSheetName: string;
    formulaAddress: string;
    rowCellSheetName: string;
    rowCellAddress: string;
    rowReplacementSheetName: string;
    rowReplacementAddress: string;
    columnCellSheetName?: string;
    columnCellAddress?: string;
    columnReplacementSheetName?: string;
    columnReplacementAddress?: string;
  }) => CellValue | undefined;
  listSheetNames?: () => string[];
  resolveBuiltin?: (name: string) => ((...args: CellValue[]) => EvaluationResult) | undefined;
}

interface ReferenceOperand {
  kind: "cell" | "range" | "row" | "col";
  sheetName?: string;
  address?: string;
  start?: string;
  end?: string;
  refKind?: "cells" | "rows" | "cols";
}

export type JsPlanInstruction =
  | { opcode: "push-number"; value: number }
  | { opcode: "push-boolean"; value: boolean }
  | { opcode: "push-string"; value: string }
  | { opcode: "push-error"; code: ErrorCode }
  | { opcode: "push-name"; name: string }
  | { opcode: "push-cell"; sheetName?: string; address: string }
  | {
      opcode: "push-range";
      sheetName?: string;
      start: string;
      end: string;
      refKind: "cells" | "rows" | "cols";
    }
  | { opcode: "push-lambda"; params: string[]; body: JsPlanInstruction[] }
  | { opcode: "unary"; operator: "+" | "-" }
  | {
      opcode: "binary";
      operator: "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | ">" | ">=" | "<" | "<=";
    }
  | {
      opcode: "call";
      callee: string;
      argc: number;
      argRefs?: Array<ReferenceOperand | undefined>;
    }
  | { opcode: "invoke"; argc: number }
  | { opcode: "begin-scope" }
  | { opcode: "bind-name"; name: string }
  | { opcode: "end-scope" }
  | { opcode: "jump-if-false"; target: number }
  | { opcode: "jump"; target: number }
  | { opcode: "return" };

type BinaryOperator = Extract<JsPlanInstruction, { opcode: "binary" }>["operator"];

type StackValue =
  | { kind: "scalar"; value: CellValue }
  | { kind: "omitted" }
  | {
      kind: "range";
      values: CellValue[];
      refKind: "cells" | "rows" | "cols";
      rows: number;
      cols: number;
    }
  | {
      kind: "lambda";
      params: string[];
      body: JsPlanInstruction[];
      scopes: Array<Map<string, StackValue>>;
    }
  | ArrayValue;

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

function error(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
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
      return formatErrorCode(value.code);
  }
}

function isTextLike(value: CellValue): boolean {
  return value.tag === ValueTag.String || value.tag === ValueTag.Empty;
}

function compareText(left: string, right: string): number {
  const normalizedLeft = left.toUpperCase();
  const normalizedRight = right.toUpperCase();
  if (normalizedLeft === normalizedRight) {
    return 0;
  }
  return normalizedLeft < normalizedRight ? -1 : 1;
}

function compareScalars(left: CellValue, right: CellValue): number | undefined {
  if (isTextLike(left) && isTextLike(right)) {
    return compareText(toStringValue(left), toStringValue(right));
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

function truthy(value: CellValue): boolean {
  return (toNumber(value) ?? 0) !== 0;
}

function popScalar(stack: StackValue[]): CellValue {
  const value = stack.pop();
  if (!value) {
    return error(ErrorCode.Value);
  }
  if (value.kind === "scalar") {
    return value.value;
  }
  if (value.kind === "omitted") {
    return error(ErrorCode.Value);
  }
  if (value.kind === "lambda") {
    return error(ErrorCode.Value);
  }
  return value.values[0] ?? emptyValue();
}

function popArgument(stack: StackValue[]): StackValue {
  return stack.pop() ?? { kind: "scalar", value: error(ErrorCode.Value) };
}

function toEvaluationResult(value: StackValue | undefined): EvaluationResult {
  if (!value) {
    return error(ErrorCode.Value);
  }
  if (value.kind === "scalar") {
    return value.value;
  }
  if (value.kind === "omitted") {
    return error(ErrorCode.Value);
  }
  if (value.kind === "lambda") {
    return error(ErrorCode.Value);
  }
  if (value.kind === "range") {
    return {
      kind: "array",
      rows: value.rows,
      cols: value.cols,
      values: value.values,
    };
  }
  return value;
}

function cloneStackValue(value: StackValue): StackValue {
  if (value.kind === "scalar") {
    return { kind: "scalar", value: value.value };
  }
  if (value.kind === "omitted") {
    return { kind: "omitted" };
  }
  if (value.kind === "range") {
    return {
      kind: "range",
      values: value.values,
      refKind: value.refKind,
      rows: value.rows,
      cols: value.cols,
    };
  }
  if (value.kind === "lambda") {
    return {
      kind: "lambda",
      params: [...value.params],
      body: value.body,
      scopes: cloneScopes(value.scopes),
    };
  }
  return { kind: "array", values: value.values, rows: value.rows, cols: value.cols };
}

function cloneScopes(scopes: readonly Map<string, StackValue>[]): Array<Map<string, StackValue>> {
  return scopes.map(
    (scope) => new Map([...scope.entries()].map(([name, value]) => [name, cloneStackValue(value)])),
  );
}

function toRangeLike(value: StackValue): RangeLikeValue {
  if (value.kind === "omitted") {
    return { kind: "range", values: [error(ErrorCode.Value)], rows: 1, cols: 1, refKind: "cells" };
  }
  if (value.kind === "lambda") {
    return { kind: "range", values: [error(ErrorCode.Value)], rows: 1, cols: 1, refKind: "cells" };
  }
  if (value.kind === "range") {
    return value;
  }
  if (value.kind === "array") {
    return {
      kind: "range",
      values: value.values,
      rows: value.rows,
      cols: value.cols,
      refKind: "cells",
    };
  }
  return { kind: "range", values: [value.value], rows: 1, cols: 1, refKind: "cells" };
}

function scalarBinary(
  operator: BinaryOperator,
  leftValue: CellValue,
  rightValue: CellValue,
): CellValue {
  if (leftValue.tag === ValueTag.Error) {
    return leftValue;
  }
  if (rightValue.tag === ValueTag.Error) {
    return rightValue;
  }

  if (operator === "&") {
    return {
      tag: ValueTag.String,
      value: `${toStringValue(leftValue)}${toStringValue(rightValue)}`,
      stringId: 0,
    };
  }

  if (["+", "-", "*", "/", "^"].includes(operator)) {
    const left = toNumber(leftValue);
    const right = toNumber(rightValue);
    if (left === undefined || right === undefined) {
      return error(ErrorCode.Value);
    }
    if (operator === "/" && right === 0) {
      return error(ErrorCode.Div0);
    }
    const value =
      operator === "+"
        ? left + right
        : operator === "-"
          ? left - right
          : operator === "*"
            ? left * right
            : operator === "/"
              ? left / right
              : left ** right;
    return { tag: ValueTag.Number, value };
  }

  const comparison = compareScalars(leftValue, rightValue);
  if (comparison === undefined) {
    return error(ErrorCode.Value);
  }
  return {
    tag: ValueTag.Boolean,
    value:
      operator === "="
        ? comparison === 0
        : operator === "<>"
          ? comparison !== 0
          : operator === ">"
            ? comparison > 0
            : operator === ">="
              ? comparison >= 0
              : operator === "<"
                ? comparison < 0
                : comparison <= 0,
  };
}

function evaluateBinary(
  operator: BinaryOperator,
  leftValue: StackValue,
  rightValue: StackValue,
): EvaluationResult {
  if (leftValue.kind === "scalar" && rightValue.kind === "scalar") {
    return scalarBinary(operator, leftValue.value, rightValue.value);
  }

  const leftRange = toRangeLike(leftValue);
  const rightRange = toRangeLike(rightValue);
  const rows =
    leftRange.rows === rightRange.rows
      ? leftRange.rows
      : leftRange.rows === 1
        ? rightRange.rows
        : rightRange.rows === 1
          ? leftRange.rows
          : 0;
  const cols =
    leftRange.cols === rightRange.cols
      ? leftRange.cols
      : leftRange.cols === 1
        ? rightRange.cols
        : rightRange.cols === 1
          ? leftRange.cols
          : 0;
  if (rows === 0 || cols === 0) {
    return error(ErrorCode.Value);
  }

  const values: CellValue[] = [];
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      const leftIndex =
        Math.min(row, leftRange.rows - 1) * leftRange.cols + Math.min(col, leftRange.cols - 1);
      const rightIndex =
        Math.min(row, rightRange.rows - 1) * rightRange.cols + Math.min(col, rightRange.cols - 1);
      values.push(
        scalarBinary(
          operator,
          leftRange.values[leftIndex] ?? emptyValue(),
          rightRange.values[rightIndex] ?? emptyValue(),
        ),
      );
    }
  }
  return rows === 1 && cols === 1
    ? (values[0] ?? emptyValue())
    : { kind: "array", values, rows, cols };
}

function stackScalar(value: CellValue): StackValue {
  return { kind: "scalar", value };
}

function normalizeScopeName(name: string): string {
  return name.toUpperCase();
}

function isSingleCellValue(value: StackValue): CellValue | undefined {
  if (value.kind === "scalar") {
    return value.value;
  }
  if (value.kind === "omitted") {
    return undefined;
  }
  if (value.kind === "lambda") {
    return undefined;
  }
  return value.rows * value.cols === 1 ? (value.values[0] ?? emptyValue()) : undefined;
}

function toRangeArgument(value: StackValue): CellValue | RangeBuiltinArgument {
  if (value.kind === "scalar") {
    return value.value;
  }
  if (value.kind === "omitted") {
    return error(ErrorCode.Value);
  }
  if (value.kind === "lambda") {
    return error(ErrorCode.Value);
  }
  return {
    kind: "range",
    values: value.values,
    refKind: value.kind === "range" ? value.refKind : "cells",
    rows: value.rows,
    cols: value.cols,
  };
}

function toPositiveInteger(value: StackValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const scalar = isSingleCellValue(value);
  const numeric = scalar ? toNumber(scalar) : undefined;
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  const integer = Math.trunc(numeric);
  return integer >= 1 ? integer : undefined;
}

function referenceOperandFromNode(node: FormulaNode): ReferenceOperand | undefined {
  switch (node.kind) {
    case "CellRef":
      return {
        kind: "cell",
        ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
        address: node.ref,
      };
    case "RangeRef":
      return {
        kind: "range",
        ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
        start: node.start,
        end: node.end,
        refKind: node.refKind,
      };
    case "RowRef":
      return {
        kind: "row",
        ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
        address: node.ref,
      };
    case "ColumnRef":
      return {
        kind: "col",
        ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
        address: node.ref,
      };
    case "BinaryExpr":
    case "BooleanLiteral":
    case "CallExpr":
    case "ErrorLiteral":
    case "InvokeExpr":
    case "NameRef":
    case "NumberLiteral":
    case "SpillRef":
    case "StringLiteral":
    case "StructuredRef":
    case "UnaryExpr":
      return undefined;
    default:
      return undefined;
  }
}

function currentCellReference(context: EvaluationContext): ReferenceOperand | undefined {
  return context.currentAddress
    ? { kind: "cell", sheetName: context.sheetName, address: context.currentAddress }
    : undefined;
}

function referenceSheetName(
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
): string | undefined {
  return ref?.sheetName ?? context.sheetName;
}

function referenceTopLeftAddress(ref: ReferenceOperand | undefined): string | undefined {
  if (!ref) {
    return undefined;
  }
  switch (ref.kind) {
    case "cell":
    case "row":
    case "col":
      return ref.address;
    case "range":
      return ref.start;
  }
}

function referenceRowNumber(
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
): number | undefined {
  const target = ref ?? currentCellReference(context);
  if (!target) {
    return undefined;
  }
  switch (target.kind) {
    case "cell":
      return parseCellAddress(target.address!, referenceSheetName(target, context)).row + 1;
    case "range":
      if (target.refKind === "rows") {
        return Number.parseInt(target.start!, 10);
      }
      if (target.refKind === "cells") {
        return parseCellAddress(target.start!, referenceSheetName(target, context)).row + 1;
      }
      return undefined;
    case "row":
      return Number.parseInt(target.address!, 10);
    case "col":
      return undefined;
  }
}

function referenceColumnNumber(
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
): number | undefined {
  const target = ref ?? currentCellReference(context);
  if (!target) {
    return undefined;
  }
  switch (target.kind) {
    case "cell":
      return parseCellAddress(target.address!, referenceSheetName(target, context)).col + 1;
    case "range":
      if (target.refKind === "cols") {
        return parseCellAddress(`${target.start!}1`, referenceSheetName(target, context)).col + 1;
      }
      if (target.refKind === "cells") {
        return parseCellAddress(target.start!, referenceSheetName(target, context)).col + 1;
      }
      return undefined;
    case "row":
      return undefined;
    case "col":
      return parseCellAddress(`${target.address!}1`, referenceSheetName(target, context)).col + 1;
  }
}

function absoluteAddress(
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
): string | undefined {
  const row = referenceRowNumber(ref, context);
  const col = referenceColumnNumber(ref, context);
  return row === undefined || col === undefined ? undefined : `$${indexToColumn(col - 1)}$${row}`;
}

function cellTypeCode(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "b";
    case ValueTag.String:
      return "l";
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.Error:
      return "v";
  }
}

function sheetNames(context: EvaluationContext): string[] {
  return context.listSheetNames?.() ?? [context.sheetName];
}

function sheetIndexByName(name: string, context: EvaluationContext): number | undefined {
  const index = sheetNames(context).findIndex(
    (sheetName) => sheetName.trim().toUpperCase() === name.trim().toUpperCase(),
  );
  return index === -1 ? undefined : index + 1;
}

function getRangeCell(range: RangeLikeValue, row: number, col: number): CellValue {
  return range.values[row * range.cols + col] ?? emptyValue();
}

function getBroadcastShape(
  values: readonly StackValue[],
): { rows: number; cols: number } | undefined {
  const ranges = values.map(toRangeLike);
  const rows = Math.max(...ranges.map((range) => range.rows));
  const cols = Math.max(...ranges.map((range) => range.cols));
  const compatible = ranges.every(
    (range) =>
      (range.rows === rows || range.rows === 1) && (range.cols === cols || range.cols === 1),
  );
  return compatible ? { rows, cols } : undefined;
}

function coerceScalarTextArgument(value: StackValue | undefined): string | CellValue {
  if (value === undefined) {
    return error(ErrorCode.Value);
  }
  const scalar = isSingleCellValue(value);
  if (!scalar) {
    return error(ErrorCode.Value);
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar;
  }
  return toStringValue(scalar);
}

function coerceOptionalBooleanArgument(
  value: StackValue | undefined,
  fallback: boolean,
): boolean | CellValue {
  if (value === undefined) {
    return fallback;
  }
  const scalar = isSingleCellValue(value);
  if (!scalar) {
    return error(ErrorCode.Value);
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar;
  }
  if (scalar.tag === ValueTag.Boolean) {
    return scalar.value;
  }
  const numeric = toNumber(scalar);
  return numeric === undefined ? error(ErrorCode.Value) : numeric !== 0;
}

function coerceOptionalMatchModeArgument(
  value: StackValue | undefined,
  fallback: 0 | 1,
): 0 | 1 | CellValue {
  if (value === undefined) {
    return fallback;
  }
  const scalar = isSingleCellValue(value);
  if (!scalar) {
    return error(ErrorCode.Value);
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar;
  }
  const numeric = toNumber(scalar);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return error(ErrorCode.Value);
  }
  const integer = Math.trunc(numeric);
  return integer === 0 || integer === 1 ? integer : error(ErrorCode.Value);
}

function coerceOptionalPositiveIntegerArgument(
  value: StackValue | undefined,
  fallback: number,
): number | CellValue {
  if (value === undefined) {
    return fallback;
  }
  const scalar = isSingleCellValue(value);
  if (!scalar) {
    return error(ErrorCode.Value);
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar;
  }
  const numeric = toNumber(scalar);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return error(ErrorCode.Value);
  }
  const integer = Math.trunc(numeric);
  return integer >= 1 ? integer : error(ErrorCode.Value);
}

function coerceOptionalTrimModeArgument(
  value: StackValue | undefined,
  fallback: 0 | 1 | 2 | 3,
): 0 | 1 | 2 | 3 | CellValue {
  if (value === undefined) {
    return fallback;
  }
  const scalar = isSingleCellValue(value);
  if (!scalar) {
    return error(ErrorCode.Value);
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar;
  }
  const numeric = toNumber(scalar);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return error(ErrorCode.Value);
  }
  const integer = Math.trunc(numeric);
  switch (integer) {
    case 0:
    case 1:
    case 2:
    case 3:
      return integer;
    default:
      return error(ErrorCode.Value);
  }
}

function isCellValueError(value: number | boolean | string | CellValue): value is CellValue {
  return typeof value === "object" && value !== null && "tag" in value;
}

function indexOfWithMatchMode(
  text: string,
  delimiter: string,
  startIndex: number,
  matchMode: 0 | 1,
): number {
  if (matchMode === 1) {
    return text.toLowerCase().indexOf(delimiter.toLowerCase(), startIndex);
  }
  return text.indexOf(delimiter, startIndex);
}

function splitTextByDelimiter(text: string, delimiter: string, matchMode: 0 | 1): string[] {
  if (delimiter === "") {
    return [text];
  }
  const parts: string[] = [];
  let cursor = 0;
  while (cursor <= text.length) {
    const found = indexOfWithMatchMode(text, delimiter, cursor, matchMode);
    if (found === -1) {
      parts.push(text.slice(cursor));
      break;
    }
    parts.push(text.slice(cursor, found));
    cursor = found + delimiter.length;
  }
  return parts;
}

function makeArrayStack(rows: number, cols: number, values: CellValue[]): StackValue {
  return { kind: "array", rows, cols, values };
}

function matrixFromStackValue(value: StackValue): MatrixValue | undefined {
  if (value.kind === "omitted" || value.kind === "lambda") {
    return undefined;
  }
  if (value.kind === "scalar") {
    return { rows: 1, cols: 1, values: [value.value] };
  }
  return { rows: value.rows, cols: value.cols, values: value.values };
}

function scalarIntegerArgument(value: StackValue | undefined): number | undefined {
  const scalar = value ? isSingleCellValue(value) : undefined;
  const numeric = scalar ? toNumber(scalar) : undefined;
  return numeric === undefined || !Number.isFinite(numeric) ? undefined : Math.trunc(numeric);
}

function vectorIntegerArgument(value: StackValue | undefined): number[] | undefined {
  if (!value) {
    return undefined;
  }
  const matrix = matrixFromStackValue(value);
  if (!matrix || !(matrix.rows === 1 || matrix.cols === 1)) {
    return undefined;
  }
  const values: number[] = [];
  for (let index = 0; index < matrix.rows * matrix.cols; index += 1) {
    const numeric = toNumber(matrix.values[index] ?? emptyValue());
    if (numeric === undefined || !Number.isFinite(numeric)) {
      return undefined;
    }
    values.push(Math.trunc(numeric));
  }
  return values;
}

function aggregateRangeSubset(
  functionArg: StackValue,
  subset: readonly CellValue[],
  context: EvaluationContext,
  totalSet?: readonly CellValue[],
): CellValue {
  if (functionArg.kind === "lambda") {
    const args: StackValue[] = [makeArrayStack(Math.max(subset.length, 1), 1, [...subset])];
    if (functionArg.params.length >= 2) {
      args.push(
        makeArrayStack(Math.max(totalSet?.length ?? 0, 1), 1, [...(totalSet ?? [emptyValue()])]),
      );
    }
    const result = applyLambda(functionArg, args, context);
    return isSingleCellValue(result) ?? error(ErrorCode.Value);
  }
  const scalar = isSingleCellValue(functionArg);
  if (scalar?.tag !== ValueTag.String) {
    return scalar?.tag === ValueTag.Error ? scalar : error(ErrorCode.Value);
  }
  const name = scalar.value.trim().toUpperCase();
  if (subset.length === 0) {
    if (name === "SUM" || name === "COUNT" || name === "COUNTA") {
      return numberValue(0);
    }
    if (name === "AVERAGE" || name === "AVG") {
      return error(ErrorCode.Div0);
    }
    return numberValue(0);
  }
  const builtin = context.resolveBuiltin?.(name) ?? getBuiltin(name);
  if (!builtin) {
    return error(ErrorCode.Name);
  }
  const result = builtin(...subset);
  return isArrayValue(result) ? scalarFromEvaluationResult(result) : result;
}

function isTrimRangeEmptyCell(value: CellValue): boolean {
  return value.tag === ValueTag.Empty;
}

function applyLambda(
  lambdaValue: StackValue,
  args: StackValue[],
  context: EvaluationContext,
): StackValue {
  if (lambdaValue.kind !== "lambda") {
    return stackScalar(
      lambdaValue.kind === "scalar" && lambdaValue.value.tag === ValueTag.Error
        ? lambdaValue.value
        : error(ErrorCode.Value),
    );
  }
  if (args.length > lambdaValue.params.length) {
    return stackScalar(error(ErrorCode.Value));
  }
  const parameterScope = new Map<string, StackValue>();
  lambdaValue.params.forEach((name, index) => {
    parameterScope.set(
      normalizeScopeName(name),
      index < args.length ? cloneStackValue(args[index]!) : { kind: "omitted" },
    );
  });
  return (
    executePlan(lambdaValue.body, context, [...cloneScopes(lambdaValue.scopes), parameterScope]) ??
    stackScalar(error(ErrorCode.Value))
  );
}

function evaluateSpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[] = [],
): StackValue | undefined {
  switch (callee) {
    case "ROW": {
      if (rawArgs.length > 1) {
        return stackScalar(error(ErrorCode.Value));
      }
      const row = referenceRowNumber(argRefs[0], context);
      return stackScalar(row === undefined ? error(ErrorCode.Value) : numberValue(row));
    }
    case "COLUMN": {
      if (rawArgs.length > 1) {
        return stackScalar(error(ErrorCode.Value));
      }
      const column = referenceColumnNumber(argRefs[0], context);
      return stackScalar(column === undefined ? error(ErrorCode.Value) : numberValue(column));
    }
    case "ISOMITTED": {
      if (rawArgs.length !== 1) {
        return stackScalar(error(ErrorCode.Value));
      }
      return stackScalar({ tag: ValueTag.Boolean, value: rawArgs[0]?.kind === "omitted" });
    }
    case "FORMULATEXT": {
      if (rawArgs.length !== 1) {
        return stackScalar(error(ErrorCode.Value));
      }
      const address = referenceTopLeftAddress(argRefs[0]);
      const sheetName = referenceSheetName(argRefs[0], context);
      if (!address || !sheetName) {
        return stackScalar(error(ErrorCode.Ref));
      }
      const formula = context.resolveFormula?.(sheetName, address);
      return stackScalar(
        formula
          ? stringValue(formula.startsWith("=") ? formula : `=${formula}`)
          : error(ErrorCode.NA),
      );
    }
    case "FORMULA": {
      if (rawArgs.length !== 1) {
        return stackScalar(error(ErrorCode.Value));
      }
      const address = referenceTopLeftAddress(argRefs[0]);
      const sheetName = referenceSheetName(argRefs[0], context);
      if (!address || !sheetName) {
        return stackScalar(error(ErrorCode.Ref));
      }
      const formula = context.resolveFormula?.(sheetName, address);
      return stackScalar(
        formula
          ? stringValue(formula.startsWith("=") ? formula : `=${formula}`)
          : error(ErrorCode.NA),
      );
    }
    case "PHONETIC": {
      if (rawArgs.length !== 1) {
        return stackScalar(error(ErrorCode.Value));
      }
      const target = rawArgs[0]!;
      if (target.kind === "scalar") {
        return stackScalar(stringValue(toStringValue(target.value)));
      }
      if (target.kind === "range") {
        return stackScalar(stringValue(toStringValue(target.values[0] ?? emptyValue())));
      }
      return stackScalar(error(ErrorCode.Value));
    }
    case "GETPIVOTDATA": {
      if (rawArgs.length < 2 || (rawArgs.length - 2) % 2 !== 0) {
        return stackScalar(error(ErrorCode.Value));
      }
      const dataFieldValue = isSingleCellValue(rawArgs[0]!);
      const address = referenceTopLeftAddress(argRefs[1]);
      const sheetName = referenceSheetName(argRefs[1], context);
      if (!dataFieldValue) {
        return stackScalar(error(ErrorCode.Value));
      }
      if (!address || !sheetName) {
        return stackScalar(error(ErrorCode.Ref));
      }
      const filters: Array<{ field: string; item: CellValue }> = [];
      for (let index = 2; index < rawArgs.length; index += 2) {
        const fieldValue = isSingleCellValue(rawArgs[index]!);
        const itemValue = isSingleCellValue(rawArgs[index + 1]!);
        if (!fieldValue || !itemValue) {
          return stackScalar(error(ErrorCode.Value));
        }
        filters.push({ field: toStringValue(fieldValue), item: itemValue });
      }
      return stackScalar(
        context.resolvePivotData?.({
          dataField: toStringValue(dataFieldValue),
          sheetName,
          address,
          filters,
        }) ?? error(ErrorCode.Ref),
      );
    }
    case "GROUPBY": {
      if (rawArgs.length < 3 || rawArgs.length > 8) {
        return stackScalar(error(ErrorCode.Value));
      }
      const rowFields = matrixFromStackValue(rawArgs[0]!);
      const values = matrixFromStackValue(rawArgs[1]!);
      if (!rowFields || !values) {
        return stackScalar(error(ErrorCode.Value));
      }
      const sortOrder =
        vectorIntegerArgument(rawArgs[5]) ??
        (rawArgs[5] ? [scalarIntegerArgument(rawArgs[5]) ?? Number.NaN] : undefined);
      const fieldHeadersMode = scalarIntegerArgument(rawArgs[3]);
      const totalDepth = scalarIntegerArgument(rawArgs[4]);
      const filterArray = rawArgs[6] ? matrixFromStackValue(rawArgs[6]) : undefined;
      const fieldRelationship = scalarIntegerArgument(rawArgs[7]);
      const groupByOptions = {
        aggregate: (subset: readonly CellValue[], totalSet?: readonly CellValue[]) =>
          aggregateRangeSubset(rawArgs[2]!, subset, context, totalSet),
        ...(fieldHeadersMode !== undefined ? { fieldHeadersMode } : {}),
        ...(totalDepth !== undefined ? { totalDepth } : {}),
        ...(sortOrder?.every(Number.isFinite) ? { sortOrder } : {}),
        ...(filterArray !== undefined ? { filterArray } : {}),
        ...(fieldRelationship !== undefined ? { fieldRelationship } : {}),
      };
      const result = evaluateGroupBy(rowFields, values, groupByOptions);
      return isArrayValue(result) ? result : stackScalar(result);
    }
    case "PIVOTBY": {
      if (rawArgs.length < 4 || rawArgs.length > 11) {
        return stackScalar(error(ErrorCode.Value));
      }
      const rowFields = matrixFromStackValue(rawArgs[0]!);
      const colFields = matrixFromStackValue(rawArgs[1]!);
      const values = matrixFromStackValue(rawArgs[2]!);
      if (!rowFields || !colFields || !values) {
        return stackScalar(error(ErrorCode.Value));
      }
      const rowSortOrder =
        vectorIntegerArgument(rawArgs[6]) ??
        (rawArgs[6] ? [scalarIntegerArgument(rawArgs[6]) ?? Number.NaN] : undefined);
      const colSortOrder =
        vectorIntegerArgument(rawArgs[8]) ??
        (rawArgs[8] ? [scalarIntegerArgument(rawArgs[8]) ?? Number.NaN] : undefined);
      const fieldHeadersMode = scalarIntegerArgument(rawArgs[4]);
      const rowTotalDepth = scalarIntegerArgument(rawArgs[5]);
      const colTotalDepth = scalarIntegerArgument(rawArgs[7]);
      const filterArray = rawArgs[9] ? matrixFromStackValue(rawArgs[9]) : undefined;
      const relativeTo = scalarIntegerArgument(rawArgs[10]);
      const pivotByOptions = {
        aggregate: (subset: readonly CellValue[], totalSet?: readonly CellValue[]) =>
          aggregateRangeSubset(rawArgs[3]!, subset, context, totalSet),
        ...(fieldHeadersMode !== undefined ? { fieldHeadersMode } : {}),
        ...(rowTotalDepth !== undefined ? { rowTotalDepth } : {}),
        ...(rowSortOrder?.every(Number.isFinite) ? { rowSortOrder } : {}),
        ...(colTotalDepth !== undefined ? { colTotalDepth } : {}),
        ...(colSortOrder?.every(Number.isFinite) ? { colSortOrder } : {}),
        ...(filterArray !== undefined ? { filterArray } : {}),
        ...(relativeTo !== undefined ? { relativeTo } : {}),
      };
      const result = evaluatePivotBy(rowFields, colFields, values, pivotByOptions);
      return isArrayValue(result) ? result : stackScalar(result);
    }
    case "MULTIPLE.OPERATIONS": {
      if (rawArgs.length !== 3 && rawArgs.length !== 5) {
        return stackScalar(error(ErrorCode.Value));
      }
      const formulaAddress = referenceTopLeftAddress(argRefs[0]);
      const formulaSheetName = referenceSheetName(argRefs[0], context);
      const rowCellAddress = referenceTopLeftAddress(argRefs[1]);
      const rowCellSheetName = referenceSheetName(argRefs[1], context);
      const rowReplacementAddress = referenceTopLeftAddress(argRefs[2]);
      const rowReplacementSheetName = referenceSheetName(argRefs[2], context);
      if (
        !formulaAddress ||
        !formulaSheetName ||
        !rowCellAddress ||
        !rowCellSheetName ||
        !rowReplacementAddress ||
        !rowReplacementSheetName
      ) {
        return stackScalar(error(ErrorCode.Ref));
      }
      const columnCellAddress =
        rawArgs.length === 5 ? referenceTopLeftAddress(argRefs[3]) : undefined;
      const columnCellSheetName =
        rawArgs.length === 5 ? referenceSheetName(argRefs[3], context) : undefined;
      const columnReplacementAddress =
        rawArgs.length === 5 ? referenceTopLeftAddress(argRefs[4]) : undefined;
      const columnReplacementSheetName =
        rawArgs.length === 5 ? referenceSheetName(argRefs[4], context) : undefined;
      if (
        rawArgs.length === 5 &&
        (!columnCellAddress ||
          !columnCellSheetName ||
          !columnReplacementAddress ||
          !columnReplacementSheetName)
      ) {
        return stackScalar(error(ErrorCode.Ref));
      }
      const request = {
        formulaSheetName,
        formulaAddress,
        rowCellSheetName,
        rowCellAddress,
        rowReplacementSheetName,
        rowReplacementAddress,
        ...(columnCellSheetName ? { columnCellSheetName } : {}),
        ...(columnCellAddress ? { columnCellAddress } : {}),
        ...(columnReplacementSheetName ? { columnReplacementSheetName } : {}),
        ...(columnReplacementAddress ? { columnReplacementAddress } : {}),
      };
      return stackScalar(context.resolveMultipleOperations?.(request) ?? error(ErrorCode.Ref));
    }
    case "CHOOSE": {
      if (rawArgs.length < 2) {
        return stackScalar(error(ErrorCode.Value));
      }
      const indexValue = isSingleCellValue(rawArgs[0]!);
      const choice = indexValue ? toNumber(indexValue) : undefined;
      if (choice === undefined || !Number.isFinite(choice)) {
        return stackScalar(error(ErrorCode.Value));
      }
      const truncated = Math.trunc(choice);
      if (truncated < 1 || truncated >= rawArgs.length) {
        return stackScalar(error(ErrorCode.Value));
      }
      return cloneStackValue(rawArgs[truncated]!);
    }
    case "SHEET": {
      if (rawArgs.length > 1) {
        return stackScalar(error(ErrorCode.Value));
      }
      if (rawArgs.length === 0) {
        const index = sheetIndexByName(context.sheetName, context);
        return stackScalar(index === undefined ? error(ErrorCode.NA) : numberValue(index));
      }
      if (argRefs[0]) {
        const index = sheetIndexByName(
          referenceSheetName(argRefs[0], context) ?? context.sheetName,
          context,
        );
        return stackScalar(index === undefined ? error(ErrorCode.NA) : numberValue(index));
      }
      const scalar = isSingleCellValue(rawArgs[0]!);
      if (scalar?.tag !== ValueTag.String) {
        return stackScalar(error(ErrorCode.NA));
      }
      const index = sheetIndexByName(scalar.value, context);
      return stackScalar(index === undefined ? error(ErrorCode.NA) : numberValue(index));
    }
    case "SHEETS": {
      if (rawArgs.length > 1) {
        return stackScalar(error(ErrorCode.Value));
      }
      if (rawArgs.length === 0) {
        return stackScalar(numberValue(sheetNames(context).length));
      }
      if (argRefs[0]) {
        return stackScalar(numberValue(1));
      }
      const scalar = isSingleCellValue(rawArgs[0]!);
      if (scalar?.tag !== ValueTag.String) {
        return stackScalar(error(ErrorCode.NA));
      }
      return stackScalar(
        sheetIndexByName(scalar.value, context) === undefined
          ? error(ErrorCode.NA)
          : numberValue(1),
      );
    }
    case "CELL": {
      if (rawArgs.length < 1 || rawArgs.length > 2) {
        return stackScalar(error(ErrorCode.Value));
      }
      const infoType = isSingleCellValue(rawArgs[0]!);
      if (infoType?.tag !== ValueTag.String) {
        return stackScalar(error(ErrorCode.Value));
      }
      const ref = rawArgs.length === 2 ? argRefs[1] : currentCellReference(context);
      if (!ref) {
        return stackScalar(error(ErrorCode.Value));
      }
      const normalizedInfoType = infoType.value.trim().toLowerCase();
      switch (normalizedInfoType) {
        case "address": {
          const address = absoluteAddress(ref, context);
          return stackScalar(address ? stringValue(address) : error(ErrorCode.Value));
        }
        case "row": {
          const row = referenceRowNumber(ref, context);
          return stackScalar(row === undefined ? error(ErrorCode.Value) : numberValue(row));
        }
        case "col": {
          const column = referenceColumnNumber(ref, context);
          return stackScalar(column === undefined ? error(ErrorCode.Value) : numberValue(column));
        }
        case "contents": {
          const address = referenceTopLeftAddress(ref);
          const sheetName = referenceSheetName(ref, context);
          if (!address || !sheetName) {
            return stackScalar(error(ErrorCode.Value));
          }
          return stackScalar(context.resolveCell(sheetName, address));
        }
        case "type": {
          const address = referenceTopLeftAddress(ref);
          const sheetName = referenceSheetName(ref, context);
          if (!address || !sheetName) {
            return stackScalar(error(ErrorCode.Value));
          }
          return stackScalar(stringValue(cellTypeCode(context.resolveCell(sheetName, address))));
        }
        case "filename":
          return stackScalar(stringValue(""));
        default:
          return stackScalar(error(ErrorCode.Value));
      }
    }
    case "INDIRECT": {
      if (rawArgs.length < 1 || rawArgs.length > 2) {
        return stackScalar(error(ErrorCode.Value));
      }
      const refText = coerceScalarTextArgument(rawArgs[0]);
      if (isCellValueError(refText)) {
        return stackScalar(refText);
      }
      const a1Mode = coerceOptionalBooleanArgument(rawArgs[1], true);
      if (isCellValueError(a1Mode)) {
        return stackScalar(a1Mode);
      }
      if (!a1Mode) {
        return stackScalar(error(ErrorCode.Value));
      }
      const normalizedRefText = refText.trim();
      if (normalizedRefText === "") {
        return stackScalar(error(ErrorCode.Ref));
      }

      try {
        const cell = parseCellAddress(normalizedRefText, context.sheetName);
        return stackScalar(context.resolveCell(cell.sheetName ?? context.sheetName, cell.text));
      } catch {
        // fall through to range/name resolution
      }

      try {
        const range = parseRangeAddress(normalizedRefText, context.sheetName);
        if (range.kind !== "cells") {
          return stackScalar(error(ErrorCode.Ref));
        }
        const targetSheetName = range.sheetName ?? context.sheetName;
        const values = context.resolveRange(
          targetSheetName,
          range.start.text,
          range.end.text,
          "cells",
        );
        const rows = range.end.row - range.start.row + 1;
        const cols = range.end.col - range.start.col + 1;
        return {
          kind: "range",
          values,
          refKind: "cells",
          rows,
          cols,
        };
      } catch {
        // fall through to name resolution
      }

      const resolvedName = context.resolveName?.(normalizedRefText);
      return stackScalar(resolvedName ?? error(ErrorCode.Ref));
    }
    case "EXPAND": {
      if (rawArgs.length < 2 || rawArgs.length > 4) {
        return stackScalar(error(ErrorCode.Value));
      }
      const source = toRangeLike(rawArgs[0]!);
      const rows = coerceOptionalPositiveIntegerArgument(rawArgs[1], source.rows);
      const cols = coerceOptionalPositiveIntegerArgument(rawArgs[2], source.cols);
      if (isCellValueError(rows)) {
        return stackScalar(rows);
      }
      if (isCellValueError(cols)) {
        return stackScalar(cols);
      }
      const padArgument = rawArgs[3];
      const padValue =
        padArgument === undefined
          ? error(ErrorCode.NA)
          : ((): CellValue => {
              const scalar = isSingleCellValue(padArgument);
              return scalar ?? error(ErrorCode.Value);
            })();
      if (padValue.tag === ValueTag.Error && padArgument !== undefined) {
        const scalar = isSingleCellValue(padArgument);
        if (!scalar) {
          return stackScalar(error(ErrorCode.Value));
        }
      }
      if (rows < source.rows || cols < source.cols) {
        return stackScalar(error(ErrorCode.Value));
      }
      const values: CellValue[] = [];
      for (let row = 0; row < rows; row += 1) {
        for (let col = 0; col < cols; col += 1) {
          values.push(
            row < source.rows && col < source.cols ? getRangeCell(source, row, col) : padValue,
          );
        }
      }
      return makeArrayStack(rows, cols, values);
    }
    case "TEXTSPLIT": {
      if (rawArgs.length < 2 || rawArgs.length > 6) {
        return stackScalar(error(ErrorCode.Value));
      }
      const text = coerceScalarTextArgument(rawArgs[0]);
      const columnDelimiter = coerceScalarTextArgument(rawArgs[1]);
      const rowDelimiter =
        rawArgs[2] === undefined ? undefined : coerceScalarTextArgument(rawArgs[2]);
      const ignoreEmpty = coerceOptionalBooleanArgument(rawArgs[3], false);
      const matchMode = coerceOptionalMatchModeArgument(rawArgs[4], 0);
      if (isCellValueError(text)) {
        return stackScalar(text);
      }
      if (isCellValueError(columnDelimiter)) {
        return stackScalar(columnDelimiter);
      }
      if (rowDelimiter !== undefined && isCellValueError(rowDelimiter)) {
        return stackScalar(rowDelimiter);
      }
      if (isCellValueError(ignoreEmpty)) {
        return stackScalar(ignoreEmpty);
      }
      if (isCellValueError(matchMode)) {
        return stackScalar(matchMode);
      }
      if (columnDelimiter === "" && rowDelimiter === undefined) {
        return stackScalar(error(ErrorCode.Value));
      }
      const padArgument = rawArgs[5];
      const padValue =
        padArgument === undefined
          ? error(ErrorCode.NA)
          : ((): CellValue => {
              const scalar = isSingleCellValue(padArgument);
              return scalar ?? error(ErrorCode.Value);
            })();
      if (padArgument !== undefined && !isSingleCellValue(padArgument)) {
        return stackScalar(error(ErrorCode.Value));
      }

      const rowSlices =
        rowDelimiter === undefined || rowDelimiter === ""
          ? [text]
          : splitTextByDelimiter(text, rowDelimiter, matchMode);
      const matrix = rowSlices.map((rowSlice) => {
        const parts =
          columnDelimiter === ""
            ? [rowSlice]
            : splitTextByDelimiter(rowSlice, columnDelimiter, matchMode);
        const filtered = ignoreEmpty ? parts.filter((part) => part !== "") : parts;
        return filtered.length === 0 ? [] : filtered;
      });
      const rows = Math.max(matrix.length, 1);
      const cols = Math.max(1, ...matrix.map((row) => row.length));
      const values: CellValue[] = [];
      for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
        const row = matrix[rowIndex] ?? [];
        for (let colIndex = 0; colIndex < cols; colIndex += 1) {
          values.push(colIndex < row.length ? stringValue(row[colIndex]!) : padValue);
        }
      }
      return makeArrayStack(rows, cols, values);
    }
    case "TRIMRANGE": {
      if (rawArgs.length < 1 || rawArgs.length > 3) {
        return stackScalar(error(ErrorCode.Value));
      }
      const source = toRangeLike(rawArgs[0]!);
      const trimRows = coerceOptionalTrimModeArgument(rawArgs[1], 3);
      const trimCols = coerceOptionalTrimModeArgument(rawArgs[2], 3);
      if (isCellValueError(trimRows)) {
        return stackScalar(trimRows);
      }
      if (isCellValueError(trimCols)) {
        return stackScalar(trimCols);
      }

      let startRow = 0;
      let endRow = source.rows - 1;
      let startCol = 0;
      let endCol = source.cols - 1;

      const trimLeadingRows = trimRows === 1 || trimRows === 3;
      const trimTrailingRows = trimRows === 2 || trimRows === 3;
      const trimLeadingCols = trimCols === 1 || trimCols === 3;
      const trimTrailingCols = trimCols === 2 || trimCols === 3;

      if (trimLeadingRows) {
        while (startRow <= endRow) {
          let hasNonEmpty = false;
          for (let col = 0; col < source.cols; col += 1) {
            if (!isTrimRangeEmptyCell(getRangeCell(source, startRow, col))) {
              hasNonEmpty = true;
              break;
            }
          }
          if (hasNonEmpty) {
            break;
          }
          startRow += 1;
        }
      }

      if (trimTrailingRows) {
        while (endRow >= startRow) {
          let hasNonEmpty = false;
          for (let col = 0; col < source.cols; col += 1) {
            if (!isTrimRangeEmptyCell(getRangeCell(source, endRow, col))) {
              hasNonEmpty = true;
              break;
            }
          }
          if (hasNonEmpty) {
            break;
          }
          endRow -= 1;
        }
      }

      if (startRow > endRow) {
        return makeArrayStack(1, 1, [emptyValue()]);
      }

      if (trimLeadingCols) {
        while (startCol <= endCol) {
          let hasNonEmpty = false;
          for (let row = startRow; row <= endRow; row += 1) {
            if (!isTrimRangeEmptyCell(getRangeCell(source, row, startCol))) {
              hasNonEmpty = true;
              break;
            }
          }
          if (hasNonEmpty) {
            break;
          }
          startCol += 1;
        }
      }

      if (trimTrailingCols) {
        while (endCol >= startCol) {
          let hasNonEmpty = false;
          for (let row = startRow; row <= endRow; row += 1) {
            if (!isTrimRangeEmptyCell(getRangeCell(source, row, endCol))) {
              hasNonEmpty = true;
              break;
            }
          }
          if (hasNonEmpty) {
            break;
          }
          endCol -= 1;
        }
      }

      if (startCol > endCol) {
        return makeArrayStack(1, 1, [emptyValue()]);
      }

      const rows = endRow - startRow + 1;
      const cols = endCol - startCol + 1;
      const values: CellValue[] = [];
      for (let row = startRow; row <= endRow; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
          values.push(getRangeCell(source, row, col));
        }
      }
      return makeArrayStack(rows, cols, values);
    }
    case "MAKEARRAY": {
      if (rawArgs.length !== 3) {
        return stackScalar(error(ErrorCode.Value));
      }
      const rows = toPositiveInteger(rawArgs[0]);
      const cols = toPositiveInteger(rawArgs[1]);
      if (rows === undefined || cols === undefined) {
        return stackScalar(error(ErrorCode.Value));
      }
      const lambda = rawArgs[2]!;
      const values: CellValue[] = [];
      for (let row = 1; row <= rows; row += 1) {
        for (let col = 1; col <= cols; col += 1) {
          const result = applyLambda(
            lambda,
            [
              stackScalar({ tag: ValueTag.Number, value: row }),
              stackScalar({ tag: ValueTag.Number, value: col }),
            ],
            context,
          );
          const scalar = isSingleCellValue(result);
          if (!scalar) {
            return stackScalar(error(ErrorCode.Value));
          }
          values.push(scalar);
        }
      }
      return { kind: "array", rows, cols, values };
    }
    case "MAP": {
      if (rawArgs.length < 2) {
        return stackScalar(error(ErrorCode.Value));
      }
      const lambda = rawArgs[rawArgs.length - 1]!;
      const inputs = rawArgs.slice(0, -1);
      const shape = getBroadcastShape(inputs);
      if (!shape) {
        return stackScalar(error(ErrorCode.Value));
      }
      const ranges = inputs.map(toRangeLike);
      const values: CellValue[] = [];
      for (let row = 0; row < shape.rows; row += 1) {
        for (let col = 0; col < shape.cols; col += 1) {
          const lambdaArgs = ranges.map((range) =>
            stackScalar(
              getRangeCell(range, Math.min(row, range.rows - 1), Math.min(col, range.cols - 1)),
            ),
          );
          const result = applyLambda(lambda, lambdaArgs, context);
          const scalar = isSingleCellValue(result);
          if (!scalar) {
            return stackScalar(error(ErrorCode.Value));
          }
          values.push(scalar);
        }
      }
      return { kind: "array", rows: shape.rows, cols: shape.cols, values };
    }
    case "BYROW":
    case "BYCOL": {
      if (rawArgs.length !== 2) {
        return stackScalar(error(ErrorCode.Value));
      }
      const source = toRangeLike(rawArgs[0]!);
      const lambda = rawArgs[1]!;
      const values: CellValue[] = [];
      if (callee === "BYROW") {
        for (let row = 0; row < source.rows; row += 1) {
          const rowValues: CellValue[] = [];
          for (let col = 0; col < source.cols; col += 1) {
            rowValues.push(getRangeCell(source, row, col));
          }
          const result = applyLambda(
            lambda,
            [{ kind: "range", values: rowValues, rows: 1, cols: source.cols, refKind: "cells" }],
            context,
          );
          const scalar = isSingleCellValue(result);
          if (!scalar) {
            return stackScalar(error(ErrorCode.Value));
          }
          values.push(scalar);
        }
        return { kind: "array", rows: source.rows, cols: 1, values };
      }
      for (let col = 0; col < source.cols; col += 1) {
        const colValues: CellValue[] = [];
        for (let row = 0; row < source.rows; row += 1) {
          colValues.push(getRangeCell(source, row, col));
        }
        const result = applyLambda(
          lambda,
          [{ kind: "range", values: colValues, rows: source.rows, cols: 1, refKind: "cells" }],
          context,
        );
        const scalar = isSingleCellValue(result);
        if (!scalar) {
          return stackScalar(error(ErrorCode.Value));
        }
        values.push(scalar);
      }
      return { kind: "array", rows: 1, cols: source.cols, values };
    }
    case "REDUCE":
    case "SCAN": {
      if (rawArgs.length !== 2 && rawArgs.length !== 3) {
        return stackScalar(error(ErrorCode.Value));
      }
      const hasInitial = rawArgs.length === 3;
      let accumulator = hasInitial ? cloneStackValue(rawArgs[0]!) : stackScalar(emptyValue());
      const source = toRangeLike(rawArgs[hasInitial ? 1 : 0]!);
      const lambda = rawArgs[hasInitial ? 2 : 1]!;
      const scanValues: CellValue[] = [];
      for (const cell of source.values) {
        accumulator = applyLambda(lambda, [accumulator, stackScalar(cell)], context);
        if (callee === "SCAN") {
          const scalar = isSingleCellValue(accumulator);
          if (!scalar) {
            return stackScalar(error(ErrorCode.Value));
          }
          scanValues.push(scalar);
        }
      }
      return callee === "SCAN"
        ? { kind: "array", rows: source.rows, cols: source.cols, values: scanValues }
        : accumulator;
    }
    default:
      return undefined;
  }
}

function lowerNode(node: FormulaNode, plan: JsPlanInstruction[]): void {
  switch (node.kind) {
    case "NumberLiteral":
      plan.push({ opcode: "push-number", value: node.value });
      return;
    case "BooleanLiteral":
      plan.push({ opcode: "push-boolean", value: node.value });
      return;
    case "StringLiteral":
      plan.push({ opcode: "push-string", value: node.value });
      return;
    case "ErrorLiteral":
      plan.push({ opcode: "push-error", code: node.code as ErrorCode });
      return;
    case "NameRef":
      plan.push({ opcode: "push-name", name: node.name });
      return;
    case "StructuredRef":
    case "SpillRef":
      plan.push({ opcode: "push-error", code: ErrorCode.Ref });
      return;
    case "CellRef":
      plan.push(
        node.sheetName
          ? { opcode: "push-cell", sheetName: node.sheetName, address: node.ref }
          : { opcode: "push-cell", address: node.ref },
      );
      return;
    case "RangeRef":
      plan.push(
        node.sheetName
          ? {
              opcode: "push-range",
              sheetName: node.sheetName,
              start: node.start,
              end: node.end,
              refKind: node.refKind,
            }
          : {
              opcode: "push-range",
              start: node.start,
              end: node.end,
              refKind: node.refKind,
            },
      );
      return;
    case "RowRef":
    case "ColumnRef":
      plan.push({ opcode: "push-number", value: Number.NaN });
      return;
    case "UnaryExpr":
      lowerNode(node.argument, plan);
      plan.push({ opcode: "unary", operator: node.operator });
      return;
    case "BinaryExpr":
      lowerNode(node.left, plan);
      lowerNode(node.right, plan);
      plan.push({ opcode: "binary", operator: node.operator });
      return;
    case "InvokeExpr":
      lowerNode(node.callee, plan);
      node.args.forEach((arg) => lowerNode(arg, plan));
      plan.push({ opcode: "invoke", argc: node.args.length });
      return;
    case "CallExpr": {
      const rewritten = rewriteSpecialCall(node);
      if (rewritten) {
        lowerNode(rewritten, plan);
        return;
      }
      const callee = node.callee.toUpperCase();
      if (callee === "LAMBDA") {
        if (node.args.length < 1) {
          plan.push({ opcode: "push-error", code: ErrorCode.Value });
          return;
        }
        const params: string[] = [];
        for (let index = 0; index < node.args.length - 1; index += 1) {
          const paramNode = node.args[index]!;
          if (paramNode.kind !== "NameRef") {
            plan.push({ opcode: "push-error", code: ErrorCode.Value });
            return;
          }
          params.push(paramNode.name);
        }
        const body: JsPlanInstruction[] = [];
        lowerNode(node.args[node.args.length - 1]!, body);
        body.push({ opcode: "return" });
        plan.push({ opcode: "push-lambda", params, body });
        return;
      }
      if (callee === "LET") {
        if (node.args.length < 3 || node.args.length % 2 === 0) {
          plan.push({ opcode: "push-error", code: ErrorCode.Value });
          return;
        }
        plan.push({ opcode: "begin-scope" });
        for (let index = 0; index < node.args.length - 1; index += 2) {
          const nameNode = node.args[index]!;
          if (nameNode.kind !== "NameRef") {
            plan.push({ opcode: "push-error", code: ErrorCode.Value });
            plan.push({ opcode: "end-scope" });
            return;
          }
          lowerNode(node.args[index + 1]!, plan);
          plan.push({ opcode: "bind-name", name: nameNode.name });
        }
        lowerNode(node.args[node.args.length - 1]!, plan);
        plan.push({ opcode: "end-scope" });
        return;
      }
      if (callee === "IF" && node.args.length === 3) {
        lowerNode(node.args[0]!, plan);
        const jumpIfFalseIndex = plan.push({ opcode: "jump-if-false", target: -1 }) - 1;
        lowerNode(node.args[1]!, plan);
        const jumpIndex = plan.push({ opcode: "jump", target: -1 }) - 1;
        const falseTarget = plan.length;
        lowerNode(node.args[2]!, plan);
        const endTarget = plan.length;
        plan[jumpIfFalseIndex] = { opcode: "jump-if-false", target: falseTarget };
        plan[jumpIndex] = { opcode: "jump", target: endTarget };
        return;
      }
      if (!hasBuiltin(callee)) {
        lowerNode({ kind: "NameRef", name: node.callee }, plan);
        node.args.forEach((arg) => lowerNode(arg, plan));
        plan.push({ opcode: "invoke", argc: node.args.length });
        return;
      }

      const aggregateArgumentIndex = callee === "GROUPBY" ? 2 : callee === "PIVOTBY" ? 3 : -1;
      node.args.forEach((arg, index) => {
        if (index === aggregateArgumentIndex && arg.kind === "NameRef") {
          plan.push({ opcode: "push-string", value: arg.name });
          return;
        }
        lowerNode(arg, plan);
      });
      plan.push({
        opcode: "call",
        callee,
        argc: node.args.length,
        argRefs: node.args.map((arg) => referenceOperandFromNode(arg)),
      });
      return;
    }
  }
}

export function lowerToPlan(node: FormulaNode): JsPlanInstruction[] {
  const plan: JsPlanInstruction[] = [];
  lowerNode(node, plan);
  plan.push({ opcode: "return" });
  return plan;
}

function executePlan(
  plan: readonly JsPlanInstruction[],
  context: EvaluationContext,
  initialScopes: readonly Map<string, StackValue>[] = [],
): StackValue | undefined {
  const stack: StackValue[] = [];
  const scopes: Array<Map<string, StackValue>> = cloneScopes(initialScopes);
  let pc = 0;

  while (pc < plan.length) {
    const instruction = plan[pc]!;
    switch (instruction.opcode) {
      case "push-number":
        stack.push({ kind: "scalar", value: { tag: ValueTag.Number, value: instruction.value } });
        break;
      case "push-boolean":
        stack.push({ kind: "scalar", value: { tag: ValueTag.Boolean, value: instruction.value } });
        break;
      case "push-string":
        stack.push({
          kind: "scalar",
          value: { tag: ValueTag.String, value: instruction.value, stringId: 0 },
        });
        break;
      case "push-error":
        stack.push({ kind: "scalar", value: error(instruction.code) });
        break;
      case "push-name":
        {
          let scopedValue: StackValue | undefined;
          for (let index = scopes.length - 1; index >= 0; index -= 1) {
            const found = scopes[index]!.get(normalizeScopeName(instruction.name));
            if (found) {
              scopedValue = found;
              break;
            }
          }
          stack.push(
            scopedValue
              ? cloneStackValue(scopedValue)
              : {
                  kind: "scalar",
                  value: context.resolveName?.(instruction.name) ?? error(ErrorCode.Name),
                },
          );
        }
        break;
      case "push-cell":
        stack.push({
          kind: "scalar",
          value: context.resolveCell(
            instruction.sheetName ?? context.sheetName,
            instruction.address,
          ),
        });
        break;
      case "push-range":
        {
          const values = context.resolveRange(
            instruction.sheetName ?? context.sheetName,
            instruction.start,
            instruction.end,
            instruction.refKind,
          );
          let rows = values.length;
          let cols = 1;
          if (instruction.refKind === "cells") {
            try {
              const sheetPrefix = instruction.sheetName ? `${instruction.sheetName}!` : "";
              const range = parseRangeAddress(
                `${sheetPrefix}${instruction.start}:${instruction.end}`,
              );
              if (range.kind === "cells") {
                rows = range.end.row - range.start.row + 1;
                cols = range.end.col - range.start.col + 1;
              }
            } catch {
              rows = values.length;
              cols = 1;
            }
          }
          stack.push({
            kind: "range",
            values,
            refKind: instruction.refKind,
            rows,
            cols,
          });
        }
        break;
      case "push-lambda":
        stack.push({
          kind: "lambda",
          params: [...instruction.params],
          body: instruction.body,
          scopes: cloneScopes(scopes),
        });
        break;
      case "unary": {
        const value = popScalar(stack);
        const numeric = toNumber(value);
        stack.push({
          kind: "scalar",
          value:
            numeric === undefined
              ? error(ErrorCode.Value)
              : { tag: ValueTag.Number, value: instruction.operator === "-" ? -numeric : numeric },
        });
        break;
      }
      case "binary": {
        const right = popArgument(stack);
        const left = popArgument(stack);
        const result = evaluateBinary(instruction.operator, left, right);
        stack.push(isArrayValue(result) ? result : { kind: "scalar", value: result });
        break;
      }
      case "begin-scope":
        scopes.push(new Map());
        break;
      case "bind-name": {
        const scope = scopes[scopes.length - 1];
        if (!scope) {
          stack.push({ kind: "scalar", value: error(ErrorCode.Value) });
          break;
        }
        scope.set(normalizeScopeName(instruction.name), cloneStackValue(popArgument(stack)));
        break;
      }
      case "end-scope":
        scopes.pop();
        break;
      case "call": {
        const rawArgs: StackValue[] = [];
        for (let index = 0; index < instruction.argc; index += 1) {
          rawArgs.unshift(popArgument(stack));
        }
        const specialResult = evaluateSpecialCall(
          instruction.callee,
          rawArgs,
          context,
          instruction.argRefs,
        );
        if (specialResult) {
          stack.push(specialResult);
          break;
        }
        const lookupBuiltin = getLookupBuiltin(instruction.callee);
        if (lookupBuiltin) {
          const args: Array<CellValue | RangeBuiltinArgument> = [];
          for (const rawArg of rawArgs) {
            args.push(toRangeArgument(rawArg));
          }
          const result = lookupBuiltin(...args);
          stack.push(isArrayValue(result) ? result : { kind: "scalar", value: result });
          break;
        }

        const builtin =
          context.resolveBuiltin?.(instruction.callee) ?? getBuiltin(instruction.callee);
        if (!builtin) {
          stack.push({ kind: "scalar", value: error(ErrorCode.Name) });
          break;
        }
        const args: CellValue[] = [];
        for (const rawArg of rawArgs) {
          if (rawArg.kind === "scalar") {
            args.push(rawArg.value);
            continue;
          }
          if (rawArg.kind === "omitted") {
            args.push(error(ErrorCode.Value));
            continue;
          }
          if (rawArg.kind === "lambda") {
            args.push(error(ErrorCode.Value));
            continue;
          }
          args.push(...rawArg.values);
        }
        const result = builtin(...args);
        stack.push(isArrayValue(result) ? result : { kind: "scalar", value: result });
        break;
      }
      case "invoke": {
        const args: StackValue[] = [];
        for (let index = 0; index < instruction.argc; index += 1) {
          args.unshift(popArgument(stack));
        }
        const callee = popArgument(stack);
        stack.push(applyLambda(callee, args, context));
        break;
      }
      case "jump-if-false": {
        const value = popScalar(stack);
        if (value.tag === ValueTag.Error) {
          return stackScalar(value);
        }
        if (!truthy(value)) {
          pc = instruction.target;
          continue;
        }
        break;
      }
      case "jump":
        pc = instruction.target;
        continue;
      case "return":
        return stack.pop();
    }
    pc += 1;
  }

  return stack.pop();
}

export function evaluatePlanResult(
  plan: readonly JsPlanInstruction[],
  context: EvaluationContext,
): EvaluationResult {
  return toEvaluationResult(executePlan(plan, context));
}

export function evaluatePlan(
  plan: readonly JsPlanInstruction[],
  context: EvaluationContext,
): CellValue {
  return scalarFromEvaluationResult(evaluatePlanResult(plan, context));
}

export function evaluateAst(node: FormulaNode, context: EvaluationContext): CellValue {
  return evaluatePlan(lowerToPlan(node), context);
}

export function evaluateAstResult(node: FormulaNode, context: EvaluationContext): EvaluationResult {
  return evaluatePlanResult(lowerToPlan(node), context);
}
