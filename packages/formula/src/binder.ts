import { BuiltinId, FormulaMode, MAX_WASM_RANGE_CELLS } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import { builtinWasmEnabledNames } from "./builtin-capabilities.js";
import { formatRangeAddress, parseRangeAddress } from "./addressing.js";
import { hasBuiltin } from "./builtins.js";
import { rewriteSpecialCall } from "./special-call-rewrites.js";

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

const RANGE_SAFE_BUILTINS = new Set(["SUM", "AVG", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA"]);

const AXIS_AGGREGATE_CODES = new Map<string, number>([
  ["SUM", 1],
  ["AVERAGE", 2],
  ["AVG", 2],
  ["MIN", 3],
  ["MAX", 4],
  ["COUNT", 5],
  ["COUNTA", 6],
]);

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

function getNativeAxisAggregateCode(node: FormulaNode): number | null {
  if (
    node.kind !== "CallExpr" ||
    node.callee.toUpperCase() !== "LAMBDA" ||
    node.args.length !== 2
  ) {
    return null;
  }
  const [param, body] = node.args;
  if (param?.kind !== "NameRef" || body?.kind !== "CallExpr" || body.args.length !== 1) {
    return null;
  }
  const aggregateCode = AXIS_AGGREGATE_CODES.get(body.callee.toUpperCase());
  if (aggregateCode === undefined) {
    return null;
  }
  return body.args[0]?.kind === "NameRef" &&
    body.args[0].name.trim().toUpperCase() === param.name.trim().toUpperCase()
    ? aggregateCode
    : null;
}

function getNativeRunningFoldCode(node: FormulaNode): number | null {
  if (
    node.kind !== "CallExpr" ||
    node.callee.toUpperCase() !== "LAMBDA" ||
    node.args.length !== 3
  ) {
    return null;
  }
  const [acc, value, body] = node.args;
  if (acc?.kind !== "NameRef" || value?.kind !== "NameRef" || body?.kind !== "BinaryExpr") {
    return null;
  }
  const foldCode = body.operator === "+" ? 1 : body.operator === "*" ? 2 : null;
  if (foldCode === null) {
    return null;
  }
  const left = body.left;
  const right = body.right;
  const accName = acc.name.trim().toUpperCase();
  const valueName = value.name.trim().toUpperCase();
  return left.kind === "NameRef" &&
    right.kind === "NameRef" &&
    ((left.name.trim().toUpperCase() === accName &&
      right.name.trim().toUpperCase() === valueName) ||
      (left.name.trim().toUpperCase() === valueName && right.name.trim().toUpperCase() === accName))
    ? foldCode
    : null;
}

function isNativeMakearraySumLambda(node: FormulaNode): boolean {
  if (
    node.kind !== "CallExpr" ||
    node.callee.toUpperCase() !== "LAMBDA" ||
    node.args.length !== 3
  ) {
    return false;
  }
  const [rowParam, colParam, body] = node.args;
  if (rowParam?.kind !== "NameRef" || colParam?.kind !== "NameRef" || body?.kind !== "BinaryExpr") {
    return false;
  }
  if (body.operator !== "+") {
    return false;
  }
  const left = body.left;
  const right = body.right;
  const rowName = rowParam.name.trim().toUpperCase();
  const colName = colParam.name.trim().toUpperCase();
  return (
    left.kind === "NameRef" &&
    right.kind === "NameRef" &&
    ((left.name.trim().toUpperCase() === rowName && right.name.trim().toUpperCase() === colName) ||
      (left.name.trim().toUpperCase() === colName && right.name.trim().toUpperCase() === rowName))
  );
}

function isCellVectorNode(node: FormulaNode): boolean {
  if (node.kind !== "RangeRef") {
    return false;
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : "";
    const range = parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`);
    return (
      range.kind === "cells" &&
      (range.start.row === range.end.row || range.start.col === range.end.col)
    );
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
    case "SIN":
    case "COS":
    case "TAN":
    case "ASIN":
    case "ACOS":
    case "ATAN":
    case "DEGREES":
    case "RADIANS":
    case "EXP":
    case "LN":
    case "LOG10":
    case "SQRT":
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
    case "DAYS":
      return argc === 2;
    case "WEEKNUM":
      return argc === 1 || argc === 2;
    case "WORKDAY":
    case "NETWORKDAYS":
      return argc === 2 || argc === 3;
    case "REPLACE":
      return argc === 4;
    case "SUBSTITUTE":
      return argc === 3 || argc === 4;
    case "REPT":
      return argc === 2;
    case "EXACT":
    case "ATAN2":
    case "POWER":
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
    case "LOG":
      return argc === 1 || argc === 2;
    case "T":
    case "N":
    case "TYPE":
    case "GAUSS":
    case "PHI":
    case "NORMSDIST":
    case "NORMSINV":
      return argc === 1;
    case "DELTA":
    case "GESTEP":
    case "LOGNORMDIST":
    case "EFFECT":
    case "NOMINAL":
    case "RRI":
    case "PERMUT":
    case "PERMUTATIONA":
      return argc === 2;
    case "STANDARDIZE":
    case "NORMINV":
    case "LOGINV":
    case "PDURATION":
    case "CONFIDENCE.NORM":
    case "CONFIDENCE":
    case "CRITBINOM":
    case "BINOM.INV":
      return argc === 3;
    case "ERF":
      return argc === 1 || argc === 2;
    case "ERF.PRECISE":
    case "ERFC":
    case "ERFC.PRECISE":
    case "FISHER":
    case "FISHERINV":
    case "GAMMALN":
    case "GAMMALN.PRECISE":
    case "GAMMA":
      return argc === 1;
    case "CHIDIST":
    case "CHISQ.DIST.RT":
      return argc === 2;
    case "CHISQ.DIST":
      return argc === 3;
    case "WEIBULL":
    case "WEIBULL.DIST":
    case "GAMMADIST":
    case "GAMMA.DIST":
    case "BINOMDIST":
    case "BINOM.DIST":
    case "NEGBINOM.DIST":
      return argc === 4;
    case "EXPONDIST":
    case "EXPON.DIST":
    case "POISSON":
    case "POISSON.DIST":
    case "NEGBINOMDIST":
      return argc === 3;
    case "BINOM.DIST.RANGE":
      return argc === 3 || argc === 4;
    case "HYPGEOMDIST":
      return argc === 4;
    case "HYPGEOM.DIST":
      return argc === 5;
    case "NORMDIST":
      return argc === 4;
    case "NORM.DIST":
      return argc === 4;
    case "NORM.INV":
      return argc === 3;
    case "NORM.S.DIST":
      return argc === 1 || argc === 2;
    case "NORM.S.INV":
      return argc === 1;
    case "LOGNORM.DIST":
      return argc === 3 || argc === 4;
    case "LOGNORM.INV":
      return argc === 3;
    case "MODE":
    case "MODE.SNGL":
    case "STDEV":
    case "STDEV.P":
    case "STDEV.S":
    case "STDEVA":
    case "STDEVP":
    case "STDEVPA":
    case "VAR":
    case "VAR.P":
    case "VAR.S":
    case "VARA":
    case "VARP":
    case "VARPA":
    case "SKEW":
    case "SKEW.P":
    case "KURT":
    case "NPV":
      return argc >= 1;
    case "FV":
    case "PV":
    case "PMT":
    case "NPER":
      return argc >= 3 && argc <= 5;
    case "IPMT":
    case "PPMT":
      return argc >= 4 && argc <= 6;
    case "ISPMT":
      return argc === 4;
    case "DATE":
    case "TIME":
    case "DATEDIF":
      return argc === 3;
    case "FVSCHEDULE":
      return argc >= 2;
    case "SLN":
      return argc === 3;
    case "DB":
    case "DDB":
      return argc === 4 || argc === 5;
    case "SYD":
      return argc === 4;
    case "VDB":
      return argc >= 5 && argc <= 7;
    case "EDATE":
    case "EOMONTH":
      return argc === 2;
    case "AND":
    case "OR":
      return argc >= 1;
    case "SEQUENCE":
      return argc >= 1 && argc <= 4;
    case "EXPAND":
      return argc >= 2 && argc <= 4;
    case "FILTER":
      return argc === 2 || argc === 3;
    case "UNIQUE":
      return argc >= 1 && argc <= 3;
    case "TRIMRANGE":
      return argc >= 1 && argc <= 3;
    case "OFFSET":
      return argc >= 3 && argc <= 5;
    case "TAKE":
    case "DROP":
      return argc >= 1 && argc <= 3;
    case "CHOOSECOLS":
    case "CHOOSEROWS":
      return argc >= 2;
    case "SORT":
      return argc >= 1 && argc <= 4;
    case "SORTBY":
      return argc >= 2;
    case "TOCOL":
    case "TOROW":
      return argc >= 1 && argc <= 3;
    case "WRAPROWS":
    case "WRAPCOLS":
      return argc >= 2 && argc <= 4;
    case "LOOKUP":
      return argc === 2 || argc === 3;
    case "AREAS":
    case "COLUMNS":
    case "ROWS":
    case "TRANSPOSE":
      return argc === 1;
    case "HSTACK":
    case "VSTACK":
      return argc >= 1;
    case "ARRAYTOTEXT":
      return argc === 1 || argc === 2;
    case "MINIFS":
    case "MAXIFS":
      return argc >= 3 && argc % 2 === 1;
    default:
      return true;
  }
}

export function bindFormula(ast: FormulaNode): BoundFormula {
  const deps = new Set<string>();
  const symbolicNames = new Set<string>();
  const symbolicTables = new Set<string>();
  const symbolicSpills = new Set<string>();

  function collectDeps(node: FormulaNode, localNames: ReadonlySet<string> = new Set()): void {
    switch (node.kind) {
      case "NumberLiteral":
      case "BooleanLiteral":
      case "StringLiteral":
      case "ErrorLiteral":
        break;
      case "NameRef":
        if (!localNames.has(node.name)) {
          symbolicNames.add(node.name);
        }
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
        deps.add(
          formatRangeAddress(
            parseRangeAddress(
              node.sheetName
                ? `${node.sheetName}!${node.start}:${node.end}`
                : `${node.start}:${node.end}`,
            ),
          ),
        );
        break;
      case "UnaryExpr":
        collectDeps(node.argument, localNames);
        break;
      case "BinaryExpr":
        collectDeps(node.left, localNames);
        collectDeps(node.right, localNames);
        break;
      case "CallExpr": {
        const rewritten = rewriteSpecialCall(node);
        if (rewritten) {
          collectDeps(rewritten, localNames);
          break;
        }
        const callee = node.callee.toUpperCase();
        if (callee === "LET" && node.args.length >= 3 && node.args.length % 2 === 1) {
          const scopedNames = new Set(localNames);
          for (let index = 0; index < node.args.length - 1; index += 2) {
            const nameNode = node.args[index]!;
            collectDeps(node.args[index + 1]!, scopedNames);
            if (nameNode.kind === "NameRef") {
              scopedNames.add(nameNode.name);
            }
          }
          collectDeps(node.args[node.args.length - 1]!, scopedNames);
          break;
        }
        if (callee === "LAMBDA" && node.args.length >= 1) {
          const scopedNames = new Set(localNames);
          for (let index = 0; index < node.args.length - 1; index += 1) {
            const paramNode = node.args[index]!;
            if (paramNode.kind === "NameRef") {
              scopedNames.add(paramNode.name);
            }
          }
          collectDeps(node.args[node.args.length - 1]!, scopedNames);
          break;
        }
        if (!hasBuiltin(callee) && !localNames.has(node.callee)) {
          symbolicNames.add(node.callee);
        }
        node.args.forEach((arg) => {
          collectDeps(arg, localNames);
        });
        break;
      }
      case "InvokeExpr":
        collectDeps(node.callee, localNames);
        node.args.forEach((arg) => {
          collectDeps(arg, localNames);
        });
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
          const cellCount =
            (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
          return cellCount <= MAX_WASM_RANGE_CELLS;
        } catch {
          return false;
        }
      case "UnaryExpr":
        return ["+", "-"].includes(node.operator) && isWasmSafe(node.argument, true);
      case "BinaryExpr":
        return (
          ["+", "-", "*", "/", "^", "&", "=", "<>", ">", ">=", "<", "<="].includes(node.operator) &&
          isWasmSafe(node.left, true) &&
          isWasmSafe(node.right, true)
        );
      case "CallExpr": {
        const rewritten = rewriteSpecialCall(node);
        if (rewritten) {
          return isWasmSafe(rewritten, allowRange);
        }
        const callee = node.callee.toUpperCase();
        if (
          (callee === "BYROW" || callee === "BYCOL") &&
          node.args.length === 2 &&
          isWasmSafe(node.args[0]!, true) &&
          getNativeAxisAggregateCode(node.args[1]!) !== null
        ) {
          return true;
        }
        if (callee === "REDUCE" || callee === "SCAN") {
          const sourceArg = node.args.length === 3 ? node.args[1] : node.args[0];
          const lambdaArg = node.args.length === 3 ? node.args[2] : node.args[1];
          const initialArg = node.args.length === 3 ? node.args[0] : undefined;
          const foldCode = lambdaArg ? getNativeRunningFoldCode(lambdaArg) : null;
          if (
            (node.args.length === 2 || node.args.length === 3) &&
            sourceArg !== undefined &&
            lambdaArg !== undefined &&
            isWasmSafe(sourceArg, true) &&
            (initialArg === undefined || isWasmSafe(initialArg)) &&
            foldCode !== null
          ) {
            return true;
          }
        }
        if (
          callee === "MAKEARRAY" &&
          node.args.length === 3 &&
          isWasmSafe(node.args[0]!) &&
          isWasmSafe(node.args[1]!) &&
          isNativeMakearraySumLambda(node.args[2]!)
        ) {
          return true;
        }
        if (
          callee === "LET" ||
          callee === "LAMBDA" ||
          callee === "MAKEARRAY" ||
          callee === "MAP" ||
          callee === "REDUCE" ||
          callee === "SCAN" ||
          callee === "BYROW" ||
          callee === "BYCOL"
        ) {
          return false;
        }
        if (!hasBuiltin(callee) || !builtinWasmEnabledNames.has(callee)) {
          return false;
        }
        if (!isWasmSafeBuiltinArity(callee, node.args.length)) {
          return false;
        }
        return isWasmSafeBuiltinArgs(callee, node.args);
      }
      case "InvokeExpr":
        return false;
    }
  }

  function isWasmSafeBuiltinArgs(callee: string, args: readonly FormulaNode[]): boolean {
    const argc = args.length;
    const isScalarArg = (arg: FormulaNode): boolean => isWasmSafe(arg);
    const isCellRangeArg = (arg: FormulaNode): boolean =>
      isWasmSafe(arg, true) && isCellRangeNode(arg);
    const isCellVectorArg = (arg: FormulaNode): boolean =>
      isWasmSafe(arg, true) && isCellVectorNode(arg);
    const isCellOrScalarArg = (arg: FormulaNode): boolean =>
      isCellVectorArg(arg) || isScalarArg(arg);
    const isNativeSequenceArg = (arg: FormulaNode): boolean =>
      arg.kind === "CallExpr" &&
      arg.callee.toUpperCase() === "SEQUENCE" &&
      isWasmSafeBuiltinArity("SEQUENCE", arg.args.length) &&
      arg.args.every((child) => isWasmSafe(child));

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
        return args.every((arg, index) =>
          index % 2 === 0 ? isCellRangeArg(arg) : isScalarArg(arg),
        );
      case "SUMIF":
      case "AVERAGEIF":
        if (args.length !== 2 && args.length !== 3) {
          return false;
        }
        return (
          isCellRangeArg(args[0]!) &&
          isScalarArg(args[1]!) &&
          (args.length === 2 || isCellRangeArg(args[2]!))
        );
      case "SUMIFS":
      case "AVERAGEIFS":
        if (args.length < 3 || args.length % 2 === 0) {
          return false;
        }
        if (!isCellRangeArg(args[0]!)) {
          return false;
        }
        return args
          .slice(1)
          .every((arg, index) => (index % 2 === 0 ? isCellRangeArg(arg) : isScalarArg(arg)));
      case "SUMPRODUCT":
        return args.length >= 1 && args.every((arg) => isCellRangeArg(arg));
      case "MATCH":
        return (
          (args.length === 2 || args.length === 3) &&
          isScalarArg(args[0]!) &&
          isCellVectorArg(args[1]!) &&
          (args.length === 2 || isScalarArg(args[2]!))
        );
      case "XMATCH":
        return (
          args.length >= 2 &&
          args.length <= 4 &&
          isScalarArg(args[0]!) &&
          isCellVectorArg(args[1]!) &&
          args.slice(2).every((arg) => isScalarArg(arg))
        );
      case "XLOOKUP":
        return (
          args.length >= 3 &&
          args.length <= 6 &&
          isScalarArg(args[0]!) &&
          isCellVectorArg(args[1]!) &&
          isCellVectorArg(args[2]!) &&
          args.slice(3).every((arg) => isScalarArg(arg))
        );
      case "INDEX":
        return (
          (args.length === 2 || args.length === 3) &&
          isCellRangeArg(args[0]!) &&
          isScalarArg(args[1]!) &&
          (args.length === 2 || isScalarArg(args[2]!))
        );
      case "VLOOKUP":
        return (
          (args.length === 3 || args.length === 4) &&
          isScalarArg(args[0]!) &&
          isCellRangeArg(args[1]!) &&
          isScalarArg(args[2]!) &&
          (args.length === 3 || isScalarArg(args[3]!))
        );
      case "HLOOKUP":
        return (
          (args.length === 3 || args.length === 4) &&
          isScalarArg(args[0]!) &&
          isCellRangeArg(args[1]!) &&
          isScalarArg(args[2]!) &&
          (args.length === 3 || isScalarArg(args[3]!))
        );
      case "DAYS":
      case "WEEKNUM":
        return args.every((arg) => isScalarArg(arg));
      case "EXPAND":
        return (
          argc >= 2 &&
          argc <= 4 &&
          isWasmSafe(args[0]!, true) &&
          isScalarArg(args[1]!) &&
          (argc < 3 || isScalarArg(args[2]!)) &&
          (argc < 4 || isScalarArg(args[3]!))
        );
      case "WORKDAY":
        return args.length === 2
          ? args.every((arg) => isScalarArg(arg))
          : args.every((arg) => isScalarArg(arg));
      case "NETWORKDAYS":
        return args.length === 2
          ? args.every((arg) => isScalarArg(arg))
          : args.every((arg) => isScalarArg(arg));
      case "REPLACE":
      case "SUBSTITUTE":
      case "REPT":
        return args.every((arg) => isScalarArg(arg));
      case "OFFSET":
      case "TAKE":
      case "DROP":
      case "CHOOSECOLS":
      case "CHOOSEROWS":
      case "SORT":
      case "TOCOL":
      case "TOROW":
      case "WRAPROWS":
      case "WRAPCOLS":
        if (args.length === 0) {
          return false;
        }
        return isCellRangeArg(args[0]!) && args.slice(1).every((arg) => isScalarArg(arg));
      case "FILTER":
        return (
          (argc === 2 || argc === 3) &&
          isCellRangeArg(args[0]!) &&
          isWasmSafe(args[1]!, true) &&
          (argc === 2 || isScalarArg(args[2]!))
        );
      case "UNIQUE":
        return (
          argc >= 1 &&
          argc <= 3 &&
          isCellRangeArg(args[0]!) &&
          args.slice(1).every((arg) => isScalarArg(arg))
        );
      case "TRIMRANGE":
        return (
          argc >= 1 &&
          argc <= 3 &&
          isWasmSafe(args[0]!, true) &&
          args.slice(1).every((arg) => isScalarArg(arg))
        );
      case "LOOKUP":
        if (argc < 2 || argc > 3) {
          return false;
        }
        return (
          isScalarArg(args[0]!) &&
          isCellOrScalarArg(args[1]!) &&
          (argc === 2 || isCellVectorArg(args[2]!) || isScalarArg(args[2]!))
        );
      case "TRANSPOSE":
        return args.length === 1 && isWasmSafe(args[0]!, true);
      case "HSTACK":
      case "VSTACK":
        return args.length >= 1 && args.every((arg) => isWasmSafe(arg, true));
      case "AREAS":
      case "COLUMNS":
      case "ROWS":
        return args.length === 1 && isCellRangeArg(args[0]!);
      case "ARRAYTOTEXT":
        return (
          (argc === 1 || argc === 2) &&
          (isCellRangeArg(args[0]!) || isScalarArg(args[0]!)) &&
          (argc === 1 || isScalarArg(args[1]!))
        );
      case "MINIFS":
      case "MAXIFS":
        if (args.length < 3 || args.length % 2 === 0 || !isCellRangeArg(args[0]!)) {
          return false;
        }
        return args
          .slice(1)
          .every((arg, index) => (index % 2 === 0 ? isCellRangeArg(arg) : isScalarArg(arg)));
      case "SORTBY":
        if (args.length < 2) {
          return false;
        }
        return (
          isCellRangeArg(args[0]!) &&
          args
            .slice(1)
            .every((arg, index) =>
              index % 2 === 0 ? isScalarArg(arg) || isWasmSafe(arg, true) : isScalarArg(arg),
            )
        );
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
    const rewritten = rewriteSpecialCall(node);
    if (rewritten) {
      return isWasmSafe(rewritten);
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
        : FormulaMode.WasmFastPath,
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
    SIN: BuiltinId.Sin,
    COS: BuiltinId.Cos,
    TAN: BuiltinId.Tan,
    ASIN: BuiltinId.Asin,
    ACOS: BuiltinId.Acos,
    ATAN: BuiltinId.Atan,
    ATAN2: BuiltinId.Atan2,
    DEGREES: BuiltinId.Degrees,
    RADIANS: BuiltinId.Radians,
    EXP: BuiltinId.Exp,
    LN: BuiltinId.Ln,
    LOG: BuiltinId.Log,
    LOG10: BuiltinId.Log10,
    POWER: BuiltinId.Power,
    SQRT: BuiltinId.Sqrt,
    PI: BuiltinId.Pi,
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
    DATEDIF: BuiltinId.Datedif,
    TIME: BuiltinId.Time,
    HOUR: BuiltinId.Hour,
    MINUTE: BuiltinId.Minute,
    SECOND: BuiltinId.Second,
    WEEKDAY: BuiltinId.Weekday,
    DAYS: BuiltinId.Days,
    WEEKNUM: BuiltinId.Weeknum,
    WORKDAY: BuiltinId.Workday,
    NETWORKDAYS: BuiltinId.Networkdays,
    EDATE: BuiltinId.Edate,
    EOMONTH: BuiltinId.Eomonth,
    REPLACE: BuiltinId.Replace,
    SUBSTITUTE: BuiltinId.Substitute,
    REPT: BuiltinId.Rept,
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
    SEQUENCE: BuiltinId.Sequence,
    FILTER: BuiltinId.Filter,
    UNIQUE: BuiltinId.Unique,
    EXPAND: BuiltinId.Expand,
    TRIMRANGE: BuiltinId.Trimrange,
    OFFSET: BuiltinId.Offset,
    TAKE: BuiltinId.Take,
    DROP: BuiltinId.Drop,
    CHOOSECOLS: BuiltinId.Choosecols,
    CHOOSEROWS: BuiltinId.Chooserows,
    SORT: BuiltinId.Sort,
    SORTBY: BuiltinId.Sortby,
    TOCOL: BuiltinId.Tocol,
    TOROW: BuiltinId.Torow,
    WRAPROWS: BuiltinId.Wraprows,
    WRAPCOLS: BuiltinId.Wrapcols,
    LOOKUP: BuiltinId.Lookup,
    AREAS: BuiltinId.Areas,
    ARRAYTOTEXT: BuiltinId.Arraytotext,
    COLUMNS: BuiltinId.Columns,
    ROWS: BuiltinId.Rows,
    TRANSPOSE: BuiltinId.Transpose,
    HSTACK: BuiltinId.Hstack,
    VSTACK: BuiltinId.Vstack,
    MINIFS: BuiltinId.Minifs,
    MAXIFS: BuiltinId.Maxifs,
    T: BuiltinId.T,
    N: BuiltinId.N,
    TYPE: BuiltinId.Type,
    DELTA: BuiltinId.Delta,
    GESTEP: BuiltinId.Gestep,
    GAUSS: BuiltinId.Gauss,
    PHI: BuiltinId.Phi,
    STANDARDIZE: BuiltinId.Standardize,
    MODE: BuiltinId.Mode,
    "MODE.SNGL": BuiltinId.ModeSngl,
    STDEV: BuiltinId.Stdev,
    "STDEV.P": BuiltinId.StdevP,
    "STDEV.S": BuiltinId.StdevS,
    STDEVA: BuiltinId.Stdeva,
    STDEVP: BuiltinId.Stdevp,
    STDEVPA: BuiltinId.Stdevpa,
    VAR: BuiltinId.Var,
    "VAR.P": BuiltinId.VarP,
    "VAR.S": BuiltinId.VarS,
    VARA: BuiltinId.Vara,
    VARP: BuiltinId.Varp,
    VARPA: BuiltinId.Varpa,
    SKEW: BuiltinId.Skew,
    "SKEW.P": BuiltinId.SkewP,
    KURT: BuiltinId.Kurt,
    NORMDIST: BuiltinId.Normdist,
    "NORM.DIST": BuiltinId.NormDist,
    NORMINV: BuiltinId.Norminv,
    "NORM.INV": BuiltinId.NormInv,
    NORMSDIST: BuiltinId.Normsdist,
    "NORM.S.DIST": BuiltinId.NormSDist,
    NORMSINV: BuiltinId.Normsinv,
    "NORM.S.INV": BuiltinId.NormSInv,
    LOGINV: BuiltinId.Loginv,
    LOGNORMDIST: BuiltinId.Lognormdist,
    "LOGNORM.DIST": BuiltinId.LognormDist,
    "LOGNORM.INV": BuiltinId.LognormInv,
    "CONFIDENCE.NORM": BuiltinId.ConfidenceNorm,
    CONFIDENCE: BuiltinId.Confidence,
    EFFECT: BuiltinId.Effect,
    NOMINAL: BuiltinId.Nominal,
    PDURATION: BuiltinId.Pduration,
    RRI: BuiltinId.Rri,
    FV: BuiltinId.Fv,
    FVSCHEDULE: BuiltinId.Fvschedule,
    DB: BuiltinId.Db,
    DDB: BuiltinId.Ddb,
    VDB: BuiltinId.Vdb,
    PV: BuiltinId.Pv,
    PMT: BuiltinId.Pmt,
    NPER: BuiltinId.Nper,
    NPV: BuiltinId.Npv,
    IPMT: BuiltinId.Ipmt,
    PPMT: BuiltinId.Ppmt,
    ISPMT: BuiltinId.Ispmt,
    SLN: BuiltinId.Sln,
    SYD: BuiltinId.Syd,
    PERMUT: BuiltinId.Permut,
    PERMUTATIONA: BuiltinId.Permutationa,
    ERF: BuiltinId.Erf,
    "ERF.PRECISE": BuiltinId.ErfPrecise,
    ERFC: BuiltinId.Erfc,
    "ERFC.PRECISE": BuiltinId.ErfcPrecise,
    FISHER: BuiltinId.Fisher,
    FISHERINV: BuiltinId.Fisherinv,
    GAMMALN: BuiltinId.Gammaln,
    "GAMMALN.PRECISE": BuiltinId.GammalnPrecise,
    GAMMA: BuiltinId.Gamma,
    EXPONDIST: BuiltinId.Expondist,
    "EXPON.DIST": BuiltinId.ExponDist,
    POISSON: BuiltinId.Poisson,
    "POISSON.DIST": BuiltinId.PoissonDist,
    WEIBULL: BuiltinId.Weibull,
    "WEIBULL.DIST": BuiltinId.WeibullDist,
    GAMMADIST: BuiltinId.Gammadist,
    "GAMMA.DIST": BuiltinId.GammaDist,
    CHIDIST: BuiltinId.Chidist,
    "CHISQ.DIST.RT": BuiltinId.ChisqDistRt,
    "CHISQ.DIST": BuiltinId.ChisqDist,
    BINOMDIST: BuiltinId.Binomdist,
    "BINOM.DIST": BuiltinId.BinomDist,
    "BINOM.DIST.RANGE": BuiltinId.BinomDistRange,
    CRITBINOM: BuiltinId.Critbinom,
    "BINOM.INV": BuiltinId.BinomInv,
    HYPGEOMDIST: BuiltinId.Hypgeomdist,
    "HYPGEOM.DIST": BuiltinId.HypgeomDist,
    NEGBINOMDIST: BuiltinId.Negbinomdist,
    "NEGBINOM.DIST": BuiltinId.NegbinomDist,
  };
  const id = builtins[name.toUpperCase()];
  if (!id) {
    throw new Error(`Unsupported builtin for wasm: ${name}`);
  }
  return id;
}
