import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { getBuiltin } from "./builtins.js";

export interface EvaluationContext {
  sheetName: string;
  resolveCell: (sheetName: string, address: string) => CellValue;
  resolveRange: (sheetName: string, start: string, end: string) => CellValue[];
}

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
      return `#${value.code}`;
  }
}

function truthy(value: CellValue): boolean {
  return (toNumber(value) ?? 0) !== 0;
}

export function evaluateAst(node: FormulaNode, context: EvaluationContext): CellValue {
  switch (node.kind) {
    case "NumberLiteral":
      return { tag: ValueTag.Number, value: node.value };
    case "BooleanLiteral":
      return { tag: ValueTag.Boolean, value: node.value };
    case "StringLiteral":
      return { tag: ValueTag.String, value: node.value, stringId: 0 };
    case "CellRef":
      return context.resolveCell(node.sheetName ?? context.sheetName, node.ref);
    case "RangeRef": {
      const values = context.resolveRange(node.sheetName ?? context.sheetName, node.start, node.end);
      return values[0] ?? emptyValue();
    }
    case "UnaryExpr": {
      const value = evaluateAst(node.argument, context);
      const numeric = toNumber(value);
      if (numeric === undefined) return error(ErrorCode.Value);
      return { tag: ValueTag.Number, value: node.operator === "-" ? -numeric : numeric };
    }
    case "BinaryExpr": {
      const left = evaluateAst(node.left, context);
      const right = evaluateAst(node.right, context);
      if (left.tag === ValueTag.Error) return left;
      if (right.tag === ValueTag.Error) return right;

      if (node.operator === "&") {
        return { tag: ValueTag.String, value: `${toStringValue(left)}${toStringValue(right)}`, stringId: 0 };
      }

      const leftNum = toNumber(left);
      const rightNum = toNumber(right);
      if (leftNum === undefined || rightNum === undefined) return error(ErrorCode.Value);

      switch (node.operator) {
        case "+":
          return { tag: ValueTag.Number, value: leftNum + rightNum };
        case "-":
          return { tag: ValueTag.Number, value: leftNum - rightNum };
        case "*":
          return { tag: ValueTag.Number, value: leftNum * rightNum };
        case "/":
          return rightNum === 0 ? error(ErrorCode.Div0) : { tag: ValueTag.Number, value: leftNum / rightNum };
        case "^":
          return { tag: ValueTag.Number, value: leftNum ** rightNum };
        case "=":
          return { tag: ValueTag.Boolean, value: leftNum === rightNum };
        case "<>":
          return { tag: ValueTag.Boolean, value: leftNum !== rightNum };
        case ">":
          return { tag: ValueTag.Boolean, value: leftNum > rightNum };
        case ">=":
          return { tag: ValueTag.Boolean, value: leftNum >= rightNum };
        case "<":
          return { tag: ValueTag.Boolean, value: leftNum < rightNum };
        case "<=":
          return { tag: ValueTag.Boolean, value: leftNum <= rightNum };
      }
      return error(ErrorCode.Value);
    }
    case "CallExpr": {
      const builtin = getBuiltin(node.callee);
      if (!builtin) return error(ErrorCode.Name);
      const args = node.args.flatMap((arg) => {
        if (arg.kind !== "RangeRef") {
          return [evaluateAst(arg, context)];
        }
        return context.resolveRange(arg.sheetName ?? context.sheetName, arg.start, arg.end);
      });
      return builtin(...args);
    }
  }
}
