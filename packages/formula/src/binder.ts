import { BuiltinId, FormulaMode, MAX_WASM_RANGE_CELLS, type FormulaRecord } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { formatRangeAddress, parseRangeAddress } from "./addressing.js";
import { getBuiltin, hasBuiltin } from "./builtins.js";

export interface BoundFormula {
  ast: FormulaNode;
  deps: string[];
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
  "ISBLANK",
  "ISNUMBER",
  "ISTEXT",
  "DATE",
  "YEAR",
  "MONTH",
  "DAY",
  "EDATE",
  "EOMONTH"
]);
const RANGE_SAFE_BUILTINS = new Set(["SUM", "AVG", "MIN", "MAX", "COUNT", "COUNTA"]);

function isWasmSafeBuiltinArity(callee: string, argc: number): boolean {
  switch (callee) {
    case "NOT":
    case "YEAR":
    case "MONTH":
    case "DAY":
      return argc === 1;
    case "ISBLANK":
    case "ISNUMBER":
    case "ISTEXT":
      return argc === 0 || argc === 1;
    case "ROUND":
    case "FLOOR":
    case "CEILING":
      return argc === 1 || argc === 2;
    case "DATE":
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

  function collectDeps(node: FormulaNode): void {
    switch (node.kind) {
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
        break;
    }
  }

  function isWasmSafe(node: FormulaNode, allowRange = false): boolean {
    switch (node.kind) {
      case "NumberLiteral":
      case "BooleanLiteral":
        return true;
      case "StringLiteral":
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
    ISBLANK: BuiltinId.IsBlank,
    ISNUMBER: BuiltinId.IsNumber,
    ISTEXT: BuiltinId.IsText,
    DATE: BuiltinId.Date,
    YEAR: BuiltinId.Year,
    MONTH: BuiltinId.Month,
    DAY: BuiltinId.Day,
    EDATE: BuiltinId.Edate,
    EOMONTH: BuiltinId.Eomonth
  };
  const id = builtins[name.toUpperCase()];
  if (!id) {
    throw new Error(`Unsupported builtin for wasm: ${name}`);
  }
  return id;
}
