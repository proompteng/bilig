import { ErrorCode, ValueTag, formatErrorCode, type CellValue } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { parseRangeAddress } from "./addressing.js";
import { getBuiltin } from "./builtins.js";
import { getLookupBuiltin, type RangeBuiltinArgument } from "./builtins/lookup.js";
import {
  isArrayValue,
  scalarFromEvaluationResult,
  type ArrayValue,
  type EvaluationResult
} from "./runtime-values.js";

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
  | { opcode: "unary"; operator: "+" | "-" }
  | { opcode: "binary"; operator: "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | ">" | ">=" | "<" | "<=" }
  | { opcode: "call"; callee: string; argc: number }
  | { opcode: "jump-if-false"; target: number }
  | { opcode: "jump"; target: number }
  | { opcode: "return" };

type StackValue =
  | { kind: "scalar"; value: CellValue }
  | { kind: "range"; values: CellValue[]; refKind: "cells" | "rows" | "cols"; rows: number; cols: number }
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
  return value.values[0] ?? emptyValue();
}

function popArgumentValues(stack: StackValue[]): CellValue[] {
  const value = stack.pop();
  if (!value) {
    return [error(ErrorCode.Value)];
  }
  if (value.kind === "scalar") {
    return [value.value];
  }
  return value.values;
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
    case "CallExpr": {
      const callee = node.callee.toUpperCase();
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

export function evaluatePlanResult(plan: readonly JsPlanInstruction[], context: EvaluationContext): EvaluationResult {
  const stack: StackValue[] = [];
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
        stack.push({
          kind: "scalar",
          value: context.resolveName?.(instruction.name) ?? error(ErrorCode.Name)
        });
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
        const right = popScalar(stack);
        const left = popScalar(stack);
        if (left.tag === ValueTag.Error) {
          stack.push({ kind: "scalar", value: left });
          break;
        }
        if (right.tag === ValueTag.Error) {
          stack.push({ kind: "scalar", value: right });
          break;
        }

        if (instruction.operator === "&") {
          stack.push({
            kind: "scalar",
            value: { tag: ValueTag.String, value: `${toStringValue(left)}${toStringValue(right)}`, stringId: 0 }
          });
          break;
        }

        const binaryValue: CellValue = (() => {
          switch (instruction.operator) {
            case "+":
            case "-":
            case "*":
            case "/":
            case "^": {
              const leftNum = toNumber(left);
              const rightNum = toNumber(right);
              if (leftNum === undefined || rightNum === undefined) {
                return error(ErrorCode.Value);
              }

              switch (instruction.operator) {
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
              }
            }
            case "=":
            case "<>":
            case ">":
            case ">=":
            case "<":
            case "<=": {
              const comparison = compareScalars(left, right);
              if (comparison === undefined) {
                return error(ErrorCode.Value);
              }

              switch (instruction.operator) {
                case "=":
                  return { tag: ValueTag.Boolean, value: comparison === 0 };
                case "<>":
                  return { tag: ValueTag.Boolean, value: comparison !== 0 };
                case ">":
                  return { tag: ValueTag.Boolean, value: comparison > 0 };
                case ">=":
                  return { tag: ValueTag.Boolean, value: comparison >= 0 };
                case "<":
                  return { tag: ValueTag.Boolean, value: comparison < 0 };
                case "<=":
                  return { tag: ValueTag.Boolean, value: comparison <= 0 };
              }
            }
          }
        })();
        stack.push({ kind: "scalar", value: binaryValue });
        break;
      }
      case "call": {
        const lookupBuiltin = getLookupBuiltin(instruction.callee);
        if (lookupBuiltin) {
          const args: Array<CellValue | RangeBuiltinArgument> = [];
          for (let index = 0; index < instruction.argc; index += 1) {
            const rawArg = popArgument(stack);
            args.unshift(
              rawArg.kind === "scalar"
                ? rawArg.value
                : {
                    kind: "range",
                    values: rawArg.values,
                    refKind: rawArg.kind === "range" ? rawArg.refKind : "cells",
                    rows: rawArg.rows,
                    cols: rawArg.cols
                  }
            );
          }
          stack.push({ kind: "scalar", value: lookupBuiltin(...args) });
          break;
        }

        const builtin = context.resolveBuiltin?.(instruction.callee) ?? getBuiltin(instruction.callee);
        if (!builtin) {
          stack.push({ kind: "scalar", value: error(ErrorCode.Name) });
          break;
        }
        const args: CellValue[] = [];
        for (let index = 0; index < instruction.argc; index += 1) {
          const values = popArgumentValues(stack);
          args.unshift(...values);
        }
        const result = builtin(...args);
        stack.push(
          isArrayValue(result)
            ? result
            : { kind: "scalar", value: result }
        );
        break;
      }
      case "jump-if-false": {
        const value = popScalar(stack);
        if (value.tag === ValueTag.Error) {
          return value;
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
        return toEvaluationResult(stack.pop());
    }
    pc += 1;
  }

  return toEvaluationResult(stack.pop());
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
