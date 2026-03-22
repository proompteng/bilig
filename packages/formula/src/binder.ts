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
  symbolicTables: string[];
  symbolicSpills: string[];
  mode: FormulaMode;
}

const WASM_SAFE_BUILTINS = new Set([
  "SUM",
  "AVG",
  "MIN",
  "MAX",
  "COUNT",
  "COUNTA",
  "COUNTIF",
  "COUNTIFS",
  "SUMIF",
  "SUMIFS",
  "AVERAGEIF",
  "AVERAGEIFS",
  "SUMPRODUCT",
  "MATCH",
  "XMATCH",
  "XLOOKUP",
  "INDEX",
  "VLOOKUP",
  "HLOOKUP",
  "ABS",
  "MOD",
  "IF",
  "IFERROR",
  "IFNA",
  "NA",
  "AND",
  "OR",
  "NOT",
  "ROUND",
  "FLOOR",
  "CEILING",
  "LEN",
  "CONCAT",
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
  "TODAY",
  "NOW",
  "RAND",
  "EDATE",
  "EOMONTH",
  "EXACT",
  "INT",
  "ROUNDUP",
  "ROUNDDOWN",
  "LEFT",
  "RIGHT",
  "MID",
  "TRIM",
  "UPPER",
  "LOWER",
  "FIND",
  "SEARCH",
  "VALUE"
]);
const RANGE_SAFE_BUILTINS = new Set(["SUM", "AVG", "MIN", "MAX", "COUNT", "COUNTA"]);

function isCellRangeNode(node: FormulaNode): boolean {
  if (node.kind !== "RangeRef") {
    return false;
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : "";
    return parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`).kind === "cells";
  } catch {
    return false;
  }
}

function isCellVectorNode(node: FormulaNode): boolean {
  if (node.kind !== "RangeRef") {
    return false;
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : "";
    const range = parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`);
    return range.kind === "cells"
      && (range.start.row === range.end.row || range.start.col === range.end.col);
  } catch {
    return false;
  }
}

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
    case "TODAY":
    case "NOW":
    case "RAND":
      return argc === 0;
    case "NA":
      return argc === 0;
    case "IF":
      return argc === 3;
    case "IFERROR":
    case "IFNA":
      return argc === 2;
    case "WEEKDAY":
      return argc === 1 || argc === 2;
    case "EXACT":
      return argc === 2;
    case "UPPER":
    case "LOWER":
    case "TRIM":
    case "VALUE":
      return argc === 1;
    case "MATCH":
      return argc === 2 || argc === 3;
    case "XMATCH":
      return argc >= 2 && argc <= 4;
    case "XLOOKUP":
      return argc >= 3 && argc <= 6;
    case "INDEX":
      return argc === 2 || argc === 3;
    case "VLOOKUP":
    case "HLOOKUP":
      return argc === 3 || argc === 4;
    case "LEFT":
    case "RIGHT":
      return argc === 1 || argc === 2;
    case "MID":
      return argc === 3;
    case "FIND":
    case "SEARCH":
      return argc === 2 || argc === 3;
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
    case "SEQUENCE":
      return argc >= 1 && argc <= 4;
    default:
      return true;
  }
}

export function bindFormula(ast: FormulaNode): BoundFormula {
  const deps = new Set<string>();
  const symbolicNames = new Set<string>();
  const symbolicTables = new Set<string>();
  const symbolicSpills = new Set<string>();

  function collectDeps(node: FormulaNode): void {
    switch (node.kind) {
      case "NumberLiteral":
      case "BooleanLiteral":
      case "StringLiteral":
      case "ErrorLiteral":
        break;
      case "NameRef":
        symbolicNames.add(node.name);
        break;
      case "StructuredRef":
        symbolicTables.add(node.tableName);
        break;
      case "CellRef":
        deps.add(node.sheetName ? `${node.sheetName}!${node.ref}` : node.ref);
        break;
      case "SpillRef":
        symbolicSpills.add(node.sheetName ? `${node.sheetName}!${node.ref}` : node.ref);
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
      case "StringLiteral":
      case "ErrorLiteral":
        return true;
      case "NameRef":
      case "StructuredRef":
      case "SpillRef":
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
            return true;
          }
          const cellCount = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
          return cellCount <= MAX_WASM_RANGE_CELLS;
        } catch {
          return false;
        }
      case "UnaryExpr":
        return ["+", "-"].includes(node.operator) && isWasmSafe(node.argument);
      case "BinaryExpr":
        return ["+", "-", "*", "/", "^", "&", "=", "<>", ">", ">=", "<", "<="].includes(node.operator)
          && isWasmSafe(node.left)
          && isWasmSafe(node.right);
      case "CallExpr": {
        const callee = node.callee.toUpperCase();
        if (!hasBuiltin(callee) || !WASM_SAFE_BUILTINS.has(callee)) {
          return false;
        }
        if (!isWasmSafeBuiltinArity(callee, node.args.length)) {
          return false;
        }
        return isWasmSafeBuiltinArgs(callee, node.args);
      }
    }
  }

  function isWasmSafeBuiltinArgs(callee: string, args: readonly FormulaNode[]): boolean {
    const isScalarArg = (arg: FormulaNode): boolean => isWasmSafe(arg);
    const isCellRangeArg = (arg: FormulaNode): boolean => isWasmSafe(arg, true) && isCellRangeNode(arg);
    const isCellVectorArg = (arg: FormulaNode): boolean => isWasmSafe(arg, true) && isCellVectorNode(arg);
    const isNativeSequenceArg = (arg: FormulaNode): boolean =>
      arg.kind === "CallExpr"
      && arg.callee.toUpperCase() === "SEQUENCE"
      && isWasmSafeBuiltinArity("SEQUENCE", arg.args.length)
      && arg.args.every((child) => isWasmSafe(child));

    switch (callee) {
      case "SUM":
      case "AVG":
      case "MIN":
      case "MAX":
      case "COUNT":
      case "COUNTA":
        return args.every((arg) => isWasmSafe(arg, true) || isNativeSequenceArg(arg));
      case "COUNTIF":
        return args.length === 2 && isCellRangeArg(args[0]!) && isScalarArg(args[1]!);
      case "COUNTIFS":
        if (args.length === 0 || args.length % 2 !== 0) {
          return false;
        }
        return args.every((arg, index) => (index % 2 === 0 ? isCellRangeArg(arg) : isScalarArg(arg)));
      case "SUMIF":
      case "AVERAGEIF":
        if (args.length !== 2 && args.length !== 3) {
          return false;
        }
        return isCellRangeArg(args[0]!)
          && isScalarArg(args[1]!)
          && (args.length === 2 || isCellRangeArg(args[2]!));
      case "SUMIFS":
      case "AVERAGEIFS":
        if (args.length < 3 || args.length % 2 === 0) {
          return false;
        }
        if (!isCellRangeArg(args[0]!)) {
          return false;
        }
        return args.slice(1).every((arg, index) => (index % 2 === 0 ? isCellRangeArg(arg) : isScalarArg(arg)));
      case "SUMPRODUCT":
        return args.length >= 1 && args.every((arg) => isCellRangeArg(arg));
      case "MATCH":
        return (args.length === 2 || args.length === 3)
          && isScalarArg(args[0]!)
          && isCellVectorArg(args[1]!)
          && (args.length === 2 || isScalarArg(args[2]!));
      case "XMATCH":
        return (args.length >= 2 && args.length <= 4)
          && isScalarArg(args[0]!)
          && isCellVectorArg(args[1]!)
          && args.slice(2).every((arg) => isScalarArg(arg));
      case "XLOOKUP":
        return (args.length >= 3 && args.length <= 6)
          && isScalarArg(args[0]!)
          && isCellVectorArg(args[1]!)
          && isCellVectorArg(args[2]!)
          && args.slice(3).every((arg) => isScalarArg(arg));
      case "INDEX":
        return (args.length === 2 || args.length === 3)
          && isCellRangeArg(args[0]!)
          && isScalarArg(args[1]!)
          && (args.length === 2 || isScalarArg(args[2]!));
      case "VLOOKUP":
        return (args.length === 3 || args.length === 4)
          && isScalarArg(args[0]!)
          && isCellRangeArg(args[1]!)
          && isScalarArg(args[2]!)
          && (args.length === 3 || isScalarArg(args[3]!));
      case "HLOOKUP":
        return (args.length === 3 || args.length === 4)
          && isScalarArg(args[0]!)
          && isCellRangeArg(args[1]!)
          && isScalarArg(args[2]!)
          && (args.length === 3 || isScalarArg(args[3]!));
      default: {
        const allowRangeArgs = RANGE_SAFE_BUILTINS.has(callee);
        return args.every((arg) => isWasmSafe(arg, allowRangeArgs));
      }
    }
  }

  function isTopLevelWasmSafe(node: FormulaNode): boolean {
    if (node.kind !== "CallExpr") {
      return false;
    }
    const callee = node.callee.toUpperCase();
    if (callee !== "SEQUENCE") {
      return false;
    }
    if (!hasBuiltin(callee) || !isWasmSafeBuiltinArity(callee, node.args.length)) {
      return false;
    }
    return node.args.every((arg) => isWasmSafe(arg));
  }

  collectDeps(ast);
  return {
    ast,
    deps: [...deps],
    symbolicNames: [...symbolicNames],
    symbolicTables: [...symbolicTables],
    symbolicSpills: [...symbolicSpills],
    mode:
      ast.kind === "RangeRef" || (!isWasmSafe(ast) && !isTopLevelWasmSafe(ast))
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
    IFERROR: BuiltinId.Iferror,
    IFNA: BuiltinId.Ifna,
    NA: BuiltinId.Na,
    AND: BuiltinId.And,
    OR: BuiltinId.Or,
    NOT: BuiltinId.Not,
    LEN: BuiltinId.Len,
    CONCAT: BuiltinId.Concat,
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
    ROUNDDOWN: BuiltinId.RoundDown,
    LEFT: BuiltinId.Left,
    RIGHT: BuiltinId.Right,
    MID: BuiltinId.Mid,
    TRIM: BuiltinId.Trim,
    UPPER: BuiltinId.Upper,
    LOWER: BuiltinId.Lower,
    FIND: BuiltinId.Find,
    SEARCH: BuiltinId.Search,
    VALUE: BuiltinId.Value,
    TODAY: BuiltinId.Today,
    NOW: BuiltinId.Now,
    RAND: BuiltinId.Rand,
    MATCH: BuiltinId.Match,
    INDEX: BuiltinId.Index,
    VLOOKUP: BuiltinId.Vlookup,
    HLOOKUP: BuiltinId.Hlookup,
    XMATCH: BuiltinId.Xmatch,
    XLOOKUP: BuiltinId.Xlookup,
    COUNTIF: BuiltinId.Countif,
    COUNTIFS: BuiltinId.Countifs,
    SUMIF: BuiltinId.Sumif,
    SUMIFS: BuiltinId.Sumifs,
    AVERAGEIF: BuiltinId.Averageif,
    AVERAGEIFS: BuiltinId.Averageifs,
    SUMPRODUCT: BuiltinId.Sumproduct,
    SEQUENCE: BuiltinId.Sequence
  };
  const id = builtins[name.toUpperCase()];
  if (!id) {
    throw new Error(`Unsupported builtin for wasm: ${name}`);
  }
  return id;
}
