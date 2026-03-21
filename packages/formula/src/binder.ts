import { BuiltinId, FormulaMode, MAX_WASM_RANGE_CELLS } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { formatRangeAddress, parseRangeAddress } from "./addressing.js";
import { hasBuiltin } from "./builtins.js";

function assertNever(value: never): never {
  throw new Error(`Unexpected formula node: ${JSON.stringify(value)}`);
}

export interface BoundFormula {
  ast: FormulaNode;
  deps: string[];
  symbolicNames: string[];
  mode: FormulaMode;
}

const WASM_SAFE_BUILTINS = new Set([
  "SUM",
  "AVG",
  "MIN",
  "MAX",
  "COUNT",
  "COUNTA",
  "ABS",
  "MOD",
  "AND",
  "OR",
  "NOT",
  "ROUND",
  "FLOOR",
  "CEILING",
  "LEN",
  "ISBLANK",
  "ISNUMBER",
  "ISTEXT",
  "DATE",
  "YEAR",
  "MONTH",
  "DAY",
  "TIME",
  "HOUR",
  "MINUTE",
  "SECOND",
  "WEEKDAY",
  "EDATE",
  "EOMONTH",
  "EXACT",
  "INT",
  "ROUNDUP",
  "ROUNDDOWN"
]);
const RANGE_SAFE_BUILTINS = new Set(["SUM", "AVG", "MIN", "MAX", "COUNT", "COUNTA"]);

function isWasmSafeBuiltinArity(callee: string, argc: number): boolean {
  switch (callee) {
    case "NOT":
    case "LEN":
    case "YEAR":
    case "MONTH":
    case "DAY":
    case "HOUR":
    case "MINUTE":
    case "SECOND":
    case "INT":
      return argc === 1;
    case "WEEKDAY":
      return argc === 1 || argc === 2;
    case "EXACT":
      return argc === 2;
    case "ISBLANK":
    case "ISNUMBER":
    case "ISTEXT":
      return argc === 0 || argc === 1;
    case "ROUND":
    case "ROUNDUP":
    case "ROUNDDOWN":
    case "FLOOR":
    case "CEILING":
      return argc === 1 || argc === 2;
    case "DATE":
    case "TIME":
      return argc === 3;
    case "EDATE":
    case "EOMONTH":
      return argc === 2;
    case "AND":
    case "OR":
      return argc >= 1;
    default:
      return true;
  }
}

export function bindFormula(ast: FormulaNode): BoundFormula {
  const deps = new Set<string>();
  const symbolicNames = new Set<string>();

  function collectDeps(node: FormulaNode): void {
    switch (node.kind) {
      case "NumberLiteral":
      case "BooleanLiteral":
      case "StringLiteral":
        break;
      case "NameRef":
        symbolicNames.add(node.name);
        break;
      case "CellRef":
        deps.add(node.sheetName ? `${node.sheetName}!${node.ref}` : node.ref);
        break;
      case "RowRef":
      case "ColumnRef":
        throw new Error("Row and column references must appear inside a range");
      case "RangeRef":
        deps.add(formatRangeAddress(parseRangeAddress(node.sheetName ? `${node.sheetName}!${node.start}:${node.end}` : `${node.start}:${node.end}`)));
        break;
      case "UnaryExpr":
        collectDeps(node.argument);
        break;
      case "BinaryExpr":
        collectDeps(node.left);
        collectDeps(node.right);
        break;
      case "CallExpr":
        node.args.forEach(collectDeps);
        break;
      default:
        assertNever(node);
    }
  }

  function isWasmSafe(node: FormulaNode, allowRange = false): boolean {
    switch (node.kind) {
      case "NumberLiteral":
      case "BooleanLiteral":
        return true;
      case "StringLiteral":
      case "NameRef":
        return false;
      case "CellRef":
        return true;
      case "RowRef":
      case "ColumnRef":
        return false;
      case "RangeRef":
        if (!allowRange) return false;
        try {
          const sheetPrefix = node.sheetName ? `${node.sheetName}!` : "";
          const range = parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`);
          if (range.kind !== "cells") {
            return false;
          }
          const cellCount = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
          return cellCount <= MAX_WASM_RANGE_CELLS;
        } catch {
          return false;
        }
      case "UnaryExpr":
        return ["+", "-"].includes(node.operator) && isWasmSafe(node.argument);
      case "BinaryExpr":
        return node.operator !== "&" && isWasmSafe(node.left) && isWasmSafe(node.right);
      case "CallExpr": {
        const callee = node.callee.toUpperCase();
        if (!hasBuiltin(callee) || !WASM_SAFE_BUILTINS.has(callee)) {
          return false;
        }
        if (!isWasmSafeBuiltinArity(callee, node.args.length)) {
          return false;
        }
        const allowRangeArgs = RANGE_SAFE_BUILTINS.has(callee);
        return node.args.every((arg) => isWasmSafe(arg, allowRangeArgs));
      }
    }
  }

  collectDeps(ast);
  return {
    ast,
    deps: [...deps],
    symbolicNames: [...symbolicNames],
    mode:
      ast.kind === "CellRef" || ast.kind === "RangeRef" || ast.kind === "StringLiteral" || !isWasmSafe(ast)
        ? FormulaMode.JsOnly
        : FormulaMode.WasmFastPath
  };
}

export function isBuiltinAvailable(name: string): boolean {
  return hasBuiltin(name);
}

export function encodeBuiltin(name: string): BuiltinId {
  const builtins: Record<string, BuiltinId> = {
    SUM: BuiltinId.Sum,
    AVG: BuiltinId.Avg,
    MIN: BuiltinId.Min,
    MAX: BuiltinId.Max,
    COUNT: BuiltinId.Count,
    COUNTA: BuiltinId.CountA,
    ABS: BuiltinId.Abs,
    ROUND: BuiltinId.Round,
    FLOOR: BuiltinId.Floor,
    CEILING: BuiltinId.Ceiling,
    MOD: BuiltinId.Mod,
    IF: BuiltinId.If,
    AND: BuiltinId.And,
    OR: BuiltinId.Or,
    NOT: BuiltinId.Not,
    LEN: BuiltinId.Len,
    ISBLANK: BuiltinId.IsBlank,
    ISNUMBER: BuiltinId.IsNumber,
    ISTEXT: BuiltinId.IsText,
    DATE: BuiltinId.Date,
    YEAR: BuiltinId.Year,
    MONTH: BuiltinId.Month,
    DAY: BuiltinId.Day,
    TIME: BuiltinId.Time,
    HOUR: BuiltinId.Hour,
    MINUTE: BuiltinId.Minute,
    SECOND: BuiltinId.Second,
    WEEKDAY: BuiltinId.Weekday,
    EDATE: BuiltinId.Edate,
    EOMONTH: BuiltinId.Eomonth,
    EXACT: BuiltinId.Exact,
    INT: BuiltinId.Int,
    ROUNDUP: BuiltinId.RoundUp,
    ROUNDDOWN: BuiltinId.RoundDown
  };
  const id = builtins[name.toUpperCase()];
  if (!id) {
    throw new Error(`Unsupported builtin for wasm: ${name}`);
  }
  return id;
}
