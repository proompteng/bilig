import { ErrorCode, ValueTag, formatErrorCode, type CellValue } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { parseRangeAddress } from "./addressing.js";
import { getBuiltin, hasBuiltin } from "./builtins.js";
import { getLookupBuiltin, type RangeBuiltinArgument } from "./builtins/lookup.js";
import {
  isArrayValue,
  scalarFromEvaluationResult,
  type ArrayValue,
  type EvaluationResult,
  type RangeLikeValue
} from "./runtime-values.js";
import { rewriteSpecialCall } from "./special-call-rewrites.js";

export interface EvaluationContext {
  sheetName: string;
  resolveCell: (sheetName: string, address: string) => CellValue;
  resolveRange: (sheetName: string, start: string, end: string, refKind: "cells" | "rows" | "cols") => CellValue[];
  resolveName?: (name: string) => CellValue;
  resolveBuiltin?: (name: string) => ((...args: CellValue[]) => EvaluationResult) | undefined;
}

export type JsPlanInstruction =
  | { opcode: "push-number"; value: number }
  | { opcode: "push-boolean"; value: boolean }
  | { opcode: "push-string"; value: string }
  | { opcode: "push-error"; code: ErrorCode }
  | { opcode: "push-name"; name: string }
  | { opcode: "push-cell"; sheetName?: string; address: string }
  | { opcode: "push-range"; sheetName?: string; start: string; end: string; refKind: "cells" | "rows" | "cols" }
  | { opcode: "push-lambda"; params: string[]; body: JsPlanInstruction[] }
  | { opcode: "unary"; operator: "+" | "-" }
  | { opcode: "binary"; operator: "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | ">" | ">=" | "<" | "<=" }
  | { opcode: "call"; callee: string; argc: number }
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
  | { kind: "range"; values: CellValue[]; refKind: "cells" | "rows" | "cols"; rows: number; cols: number }
  | { kind: "lambda"; params: string[]; body: JsPlanInstruction[]; scopes: Array<Map<string, StackValue>> }
  | ArrayValue;

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

function error(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
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
  if (value.kind === "lambda") {
    return error(ErrorCode.Value);
  }
  if (value.kind === "range") {
    return {
      kind: "array",
      rows: value.rows,
      cols: value.cols,
      values: value.values
    };
  }
  return value;
}

function cloneStackValue(value: StackValue): StackValue {
  if (value.kind === "scalar") {
    return { kind: "scalar", value: value.value };
  }
  if (value.kind === "range") {
    return { kind: "range", values: value.values, refKind: value.refKind, rows: value.rows, cols: value.cols };
  }
  if (value.kind === "lambda") {
    return { kind: "lambda", params: [...value.params], body: value.body, scopes: cloneScopes(value.scopes) };
  }
  return { kind: "array", values: value.values, rows: value.rows, cols: value.cols };
}

function cloneScopes(scopes: readonly Map<string, StackValue>[]): Array<Map<string, StackValue>> {
  return scopes.map((scope) =>
    new Map([...scope.entries()].map(([name, value]) => [name, cloneStackValue(value)]))
  );
}

function toRangeLike(value: StackValue): RangeLikeValue {
  if (value.kind === "lambda") {
    return { kind: "range", values: [error(ErrorCode.Value)], rows: 1, cols: 1, refKind: "cells" };
  }
  if (value.kind === "range") {
    return value;
  }
  if (value.kind === "array") {
    return { kind: "range", values: value.values, rows: value.rows, cols: value.cols, refKind: "cells" };
  }
  return { kind: "range", values: [value.value], rows: 1, cols: 1, refKind: "cells" };
}

function scalarBinary(operator: BinaryOperator, leftValue: CellValue, rightValue: CellValue): CellValue {
  if (leftValue.tag === ValueTag.Error) {
    return leftValue;
  }
  if (rightValue.tag === ValueTag.Error) {
    return rightValue;
  }

  if (operator === "&") {
    return { tag: ValueTag.String, value: `${toStringValue(leftValue)}${toStringValue(rightValue)}`, stringId: 0 };
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
                : comparison <= 0
  };
}

function evaluateBinary(operator: BinaryOperator, leftValue: StackValue, rightValue: StackValue): EvaluationResult {
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
      const leftIndex = Math.min(row, leftRange.rows - 1) * leftRange.cols + Math.min(col, leftRange.cols - 1);
      const rightIndex = Math.min(row, rightRange.rows - 1) * rightRange.cols + Math.min(col, rightRange.cols - 1);
      values.push(scalarBinary(operator, leftRange.values[leftIndex] ?? emptyValue(), rightRange.values[rightIndex] ?? emptyValue()));
    }
  }
  return rows === 1 && cols === 1 ? values[0] ?? emptyValue() : { kind: "array", values, rows, cols };
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
  if (value.kind === "lambda") {
    return undefined;
  }
  return value.rows * value.cols === 1 ? value.values[0] ?? emptyValue() : undefined;
}

function toRangeArgument(value: StackValue): CellValue | RangeBuiltinArgument {
  if (value.kind === "scalar") {
    return value.value;
  }
  if (value.kind === "lambda") {
    return error(ErrorCode.Value);
  }
  return {
    kind: "range",
    values: value.values,
    refKind: value.kind === "range" ? value.refKind : "cells",
    rows: value.rows,
    cols: value.cols
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

function getRangeCell(range: RangeLikeValue, row: number, col: number): CellValue {
  return range.values[row * range.cols + col] ?? emptyValue();
}

function getBroadcastShape(values: readonly StackValue[]): { rows: number; cols: number } | undefined {
  const ranges = values.map(toRangeLike);
  const rows = Math.max(...ranges.map((range) => range.rows));
  const cols = Math.max(...ranges.map((range) => range.cols));
  const compatible = ranges.every((range) => (range.rows === rows || range.rows === 1) && (range.cols === cols || range.cols === 1));
  return compatible ? { rows, cols } : undefined;
}

function applyLambda(
  lambdaValue: StackValue,
  args: StackValue[],
  context: EvaluationContext
): StackValue {
  if (lambdaValue.kind !== "lambda") {
    return stackScalar(lambdaValue.kind === "scalar" && lambdaValue.value.tag === ValueTag.Error ? lambdaValue.value : error(ErrorCode.Value));
  }
  if (lambdaValue.params.length !== args.length) {
    return stackScalar(error(ErrorCode.Value));
  }
  const parameterScope = new Map<string, StackValue>();
  lambdaValue.params.forEach((name, index) => {
    parameterScope.set(normalizeScopeName(name), cloneStackValue(args[index]!));
  });
  return executePlan(lambdaValue.body, context, [...cloneScopes(lambdaValue.scopes), parameterScope]) ?? stackScalar(error(ErrorCode.Value));
}

function evaluateSpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext
): StackValue | undefined {
  switch (callee) {
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
          const result = applyLambda(lambda, [stackScalar({ tag: ValueTag.Number, value: row }), stackScalar({ tag: ValueTag.Number, value: col })], context);
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
              getRangeCell(range, Math.min(row, range.rows - 1), Math.min(col, range.cols - 1))
            )
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
          const result = applyLambda(lambda, [{ kind: "range", values: rowValues, rows: 1, cols: source.cols, refKind: "cells" }], context);
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
        const result = applyLambda(lambda, [{ kind: "range", values: colValues, rows: source.rows, cols: 1, refKind: "cells" }], context);
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
          : { opcode: "push-cell", address: node.ref }
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
              refKind: node.refKind
            }
          : {
              opcode: "push-range",
              start: node.start,
              end: node.end,
              refKind: node.refKind
            }
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

      node.args.forEach((arg) => lowerNode(arg, plan));
      plan.push({ opcode: "call", callee, argc: node.args.length });
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
  initialScopes: readonly Map<string, StackValue>[] = []
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
        stack.push({ kind: "scalar", value: { tag: ValueTag.String, value: instruction.value, stringId: 0 } });
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
          stack.push(scopedValue
            ? cloneStackValue(scopedValue)
            : {
                kind: "scalar",
                value: context.resolveName?.(instruction.name) ?? error(ErrorCode.Name)
              });
        }
        break;
      case "push-cell":
        stack.push({
          kind: "scalar",
          value: context.resolveCell(instruction.sheetName ?? context.sheetName, instruction.address)
        });
        break;
      case "push-range":
        {
          const values = context.resolveRange(
            instruction.sheetName ?? context.sheetName,
            instruction.start,
            instruction.end,
            instruction.refKind
          );
          let rows = values.length;
          let cols = 1;
          if (instruction.refKind === "cells") {
            try {
              const sheetPrefix = instruction.sheetName ? `${instruction.sheetName}!` : "";
              const range = parseRangeAddress(`${sheetPrefix}${instruction.start}:${instruction.end}`);
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
            cols
          });
        }
        break;
      case "push-lambda":
        stack.push({ kind: "lambda", params: [...instruction.params], body: instruction.body, scopes: cloneScopes(scopes) });
        break;
      case "unary": {
        const value = popScalar(stack);
        const numeric = toNumber(value);
        stack.push({
          kind: "scalar",
          value: numeric === undefined
            ? error(ErrorCode.Value)
            : { tag: ValueTag.Number, value: instruction.operator === "-" ? -numeric : numeric }
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
        const specialResult = evaluateSpecialCall(instruction.callee, rawArgs, context);
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

        const builtin = context.resolveBuiltin?.(instruction.callee) ?? getBuiltin(instruction.callee);
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
          if (rawArg.kind === "lambda") {
            args.push(error(ErrorCode.Value));
            continue;
          }
          args.push(...rawArg.values);
        }
        const result = builtin(...args);
        stack.push(
          isArrayValue(result)
            ? result
            : { kind: "scalar", value: result }
        );
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

export function evaluatePlanResult(plan: readonly JsPlanInstruction[], context: EvaluationContext): EvaluationResult {
  return toEvaluationResult(executePlan(plan, context));
}

export function evaluatePlan(plan: readonly JsPlanInstruction[], context: EvaluationContext): CellValue {
  return scalarFromEvaluationResult(evaluatePlanResult(plan, context));
}

export function evaluateAst(node: FormulaNode, context: EvaluationContext): CellValue {
  return evaluatePlan(lowerToPlan(node), context);
}

export function evaluateAstResult(node: FormulaNode, context: EvaluationContext): EvaluationResult {
  return evaluatePlanResult(lowerToPlan(node), context);
}
