import { BuiltinId, FormulaMode, MAX_WASM_RANGE_CELLS } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import {
  getNativeAxisAggregateCode,
  getNativeRunningFoldCode,
  isNativeMakearraySumLambda,
  isWasmSafeBuiltinArgs as checkWasmSafeBuiltinArgs,
} from "./binder-wasm-rules.js";
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
    case "IFS":
      return argc >= 2 && argc % 2 === 0;
    case "IFERROR":
    case "IFNA":
      return argc === 2;
    case "WEEKDAY":
      return argc === 1 || argc === 2;
    case "DAYS":
      return argc === 2;
    case "COUNTBLANK":
      return argc >= 1;
    case "CHOOSE":
      return argc >= 2;
    case "DAYS360":
    case "YEARFRAC":
      return argc === 2 || argc === 3;
    case "DISC":
    case "INTRATE":
    case "RECEIVED":
    case "PRICEDISC":
    case "YIELDDISC":
      return argc === 4 || argc === 5;
    case "COUPDAYBS":
    case "COUPDAYS":
    case "COUPDAYSNC":
    case "COUPNCD":
    case "COUPNUM":
    case "COUPPCD":
      return argc === 3 || argc === 4;
    case "PRICEMAT":
    case "YIELDMAT":
    case "DURATION":
    case "MDURATION":
      return argc === 5 || argc === 6;
    case "ODDFPRICE":
    case "ODDFYIELD":
    case "ODDLPRICE":
    case "ODDLYIELD":
      return argc === 7 || argc === 8;
    case "TBILLPRICE":
    case "TBILLYIELD":
    case "TBILLEQ":
      return argc === 3;
    case "IRR":
      return argc === 1 || argc === 2;
    case "MIRR":
      return argc === 3;
    case "XNPV":
      return argc === 3;
    case "XIRR":
      return argc === 2 || argc === 3;
    case "PRICE":
    case "YIELD":
      return argc === 6 || argc === 7;
    case "ISOWEEKNUM":
    case "TIMEVALUE":
      return argc === 1;
    case "WEEKNUM":
      return argc === 1 || argc === 2;
    case "WORKDAY":
    case "NETWORKDAYS":
      return argc === 2 || argc === 3;
    case "WORKDAY.INTL":
    case "NETWORKDAYS.INTL":
      return argc >= 2 && argc <= 4;
    case "COUNTIF":
    case "USE.THE.COUNTIF":
      return argc === 2;
    case "COUNTIFS":
      return argc >= 2 && argc % 2 === 0;
    case "DAVERAGE":
    case "DCOUNT":
    case "DCOUNTA":
    case "DGET":
    case "DMAX":
    case "DMIN":
    case "DPRODUCT":
    case "DSTDEV":
    case "DSTDEVP":
    case "DSUM":
    case "DVAR":
    case "DVARP":
      return argc === 3;
    case "ADDRESS":
      return argc >= 2 && argc <= 5;
    case "SUMIF":
    case "AVERAGEIF":
      return argc === 2 || argc === 3;
    case "SUMIFS":
    case "AVERAGEIFS":
      return argc >= 3 && argc % 2 === 1;
    case "REPLACE":
      return argc === 4;
    case "SUBSTITUTE":
      return argc === 3 || argc === 4;
    case "REPT":
      return argc === 2;
    case "TEXT":
      return argc === 2;
    case "PHONETIC":
      return argc === 1;
    case "TEXTBEFORE":
    case "TEXTAFTER":
      return argc >= 2 && argc <= 6;
    case "TEXTSPLIT":
      return argc >= 2 && argc <= 6;
    case "TEXTJOIN":
      return argc >= 3;
    case "POWER":
    case "CONVERT":
      return argc === 3;
    case "EXACT":
    case "ATAN2":
      return argc === 2;
    case "BESSELI":
    case "BESSELJ":
    case "BESSELK":
    case "BESSELY":
      return argc === 2;
    case "EUROCONVERT":
      return argc >= 3 && argc <= 5;
    case "UPPER":
    case "LOWER":
    case "TRIM":
    case "VALUE":
    case "CHAR":
    case "CODE":
    case "UNICODE":
    case "UNICHAR":
    case "CLEAN":
    case "ASC":
    case "JIS":
    case "DBCS":
    case "BAHTTEXT":
    case "LENB":
    case "SINH":
    case "COSH":
    case "TANH":
    case "ASINH":
    case "ACOSH":
    case "ATANH":
    case "ACOT":
    case "ACOTH":
    case "COT":
    case "COTH":
    case "CSC":
    case "CSCH":
    case "SEC":
    case "SECH":
    case "SIGN":
    case "EVEN":
    case "ODD":
    case "FACT":
    case "FACTDOUBLE":
      return argc === 1;
    case "NUMBERVALUE":
      return argc >= 1 && argc <= 3;
    case "VALUETOTEXT":
      return argc === 1 || argc === 2;
    case "DOLLAR":
      return argc >= 1 && argc <= 3;
    case "DOLLARDE":
    case "DOLLARFR":
    case "COMBIN":
    case "COMBINA":
    case "QUOTIENT":
      return argc === 2;
    case "BASE":
      return argc === 2 || argc === 3;
    case "DECIMAL":
      return argc === 2;
    case "BIN2DEC":
    case "HEX2DEC":
    case "OCT2DEC":
      return argc === 1;
    case "BIN2HEX":
    case "BIN2OCT":
    case "DEC2BIN":
    case "DEC2HEX":
    case "DEC2OCT":
    case "HEX2BIN":
    case "HEX2OCT":
    case "OCT2BIN":
    case "OCT2HEX":
      return argc === 1 || argc === 2;
    case "BITAND":
    case "BITOR":
    case "BITXOR":
      return argc >= 2;
    case "BITLSHIFT":
    case "BITRSHIFT":
      return argc === 2;
    case "MATCH":
      return argc === 2 || argc === 3;
    case "CORREL":
    case "COVAR":
    case "PEARSON":
    case "COVARIANCE.P":
    case "COVARIANCE.S":
    case "PERCENTRANK":
    case "PERCENTRANK.INC":
    case "PERCENTRANK.EXC":
    case "SMALL":
    case "LARGE":
    case "PERCENTILE":
    case "PERCENTILE.INC":
    case "PERCENTILE.EXC":
    case "QUARTILE":
    case "QUARTILE.INC":
    case "QUARTILE.EXC":
    case "RANK":
    case "RANK.EQ":
    case "RANK.AVG":
    case "INTERCEPT":
    case "RSQ":
    case "SLOPE":
    case "STEYX":
      return argc === 2 || argc === 3;
    case "MEDIAN":
    case "MODE.MULT":
    case "GCD":
    case "LCM":
    case "PRODUCT":
    case "GEOMEAN":
    case "HARMEAN":
    case "SUMSQ":
      return argc >= 1;
    case "FREQUENCY":
      return argc === 2;
    case "PROB":
      return argc === 3 || argc === 4;
    case "TRIMMEAN":
      return argc === 2;
    case "FORECAST":
    case "FORECAST.LINEAR":
      return argc === 3;
    case "TREND":
    case "GROWTH":
    case "LINEST":
    case "LOGEST":
      return argc >= 1 && argc <= 4;
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
    case "LEFTB":
    case "RIGHTB":
      return argc === 1 || argc === 2;
    case "MID":
    case "MIDB":
      return argc === 3;
    case "FIND":
    case "SEARCH":
    case "FINDB":
    case "SEARCHB":
      return argc === 2 || argc === 3;
    case "REPLACEB":
      return argc === 4;
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
      return argc === 1 || (argc === 0 && (callee === "T" || callee === "N" || callee === "TYPE"));
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
    case "CONFIDENCE.T":
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
    case "GAMMA.INV":
    case "GAMMAINV":
      return argc === 3;
    case "CHIDIST":
    case "LEGACY.CHIDIST":
    case "CHIINV":
    case "CHISQ.DIST.RT":
    case "CHISQ.INV.RT":
    case "CHISQDIST":
    case "CHISQINV":
    case "LEGACY.CHIINV":
    case "CHISQ.TEST":
    case "CHITEST":
    case "LEGACY.CHITEST":
    case "F.TEST":
    case "FTEST":
      return argc === 2;
    case "Z.TEST":
    case "ZTEST":
      return argc === 2 || argc === 3;
    case "F.DIST.RT":
    case "FDIST":
    case "LEGACY.FDIST":
      return argc === 3;
    case "CHISQ.INV":
      return argc === 2;
    case "CHISQ.DIST":
      return argc === 3;
    case "BETA.INV":
    case "BETAINV":
      return argc >= 3 && argc <= 5;
    case "BETA.DIST":
      return argc >= 4 && argc <= 6;
    case "BETADIST":
      return argc >= 3 && argc <= 5;
    case "F.DIST":
      return argc === 4;
    case "T.DIST":
      return argc === 3;
    case "T.DIST.RT":
    case "T.DIST.2T":
    case "T.INV":
    case "T.INV.2T":
    case "TINV":
      return argc === 2;
    case "TDIST":
      return argc === 3;
    case "T.TEST":
    case "TTEST":
      return argc === 4;
    case "F.INV":
    case "F.INV.RT":
    case "FINV":
    case "LEGACY.FINV":
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
    case "RATE":
      return argc >= 3 && argc <= 6;
    case "IPMT":
    case "PPMT":
      return argc >= 4 && argc <= 6;
    case "ISPMT":
      return argc === 4;
    case "CUMIPMT":
    case "CUMPRINC":
      return argc === 6;
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
    case "XOR":
      return argc >= 1;
    case "SWITCH":
      return argc >= 3;
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
        const aggregateArgumentIndex = callee === "GROUPBY" ? 2 : callee === "PIVOTBY" ? 3 : -1;
        node.args.forEach((arg, index) => {
          if (index === aggregateArgumentIndex && arg.kind === "NameRef") {
            return;
          }
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
    return checkWasmSafeBuiltinArgs(callee, args, { isWasmSafe });
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
    CHOOSE: BuiltinId.Choose,
    MIN: BuiltinId.Min,
    MAX: BuiltinId.Max,
    COUNT: BuiltinId.Count,
    COUNTA: BuiltinId.CountA,
    COUNTBLANK: BuiltinId.Countblank,
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
    IFS: BuiltinId.Ifs,
    IFERROR: BuiltinId.Iferror,
    IFNA: BuiltinId.Ifna,
    NA: BuiltinId.Na,
    AND: BuiltinId.And,
    OR: BuiltinId.Or,
    NOT: BuiltinId.Not,
    SWITCH: BuiltinId.Switch,
    XOR: BuiltinId.Xor,
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
    DAYS360: BuiltinId.Days360,
    YEARFRAC: BuiltinId.Yearfrac,
    ISOWEEKNUM: BuiltinId.Isoweeknum,
    TIMEVALUE: BuiltinId.Timevalue,
    WEEKNUM: BuiltinId.Weeknum,
    WORKDAY: BuiltinId.Workday,
    NETWORKDAYS: BuiltinId.Networkdays,
    "WORKDAY.INTL": BuiltinId.WorkdayIntl,
    "NETWORKDAYS.INTL": BuiltinId.NetworkdaysIntl,
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
    LEFTB: BuiltinId.Leftb,
    RIGHTB: BuiltinId.Rightb,
    MIDB: BuiltinId.Midb,
    TRIM: BuiltinId.Trim,
    UPPER: BuiltinId.Upper,
    LOWER: BuiltinId.Lower,
    FIND: BuiltinId.Find,
    SEARCH: BuiltinId.Search,
    FINDB: BuiltinId.Findb,
    LENB: BuiltinId.Lenb,
    SEARCHB: BuiltinId.Searchb,
    REPLACEB: BuiltinId.Replaceb,
    ADDRESS: BuiltinId.Address,
    DOLLAR: BuiltinId.Dollar,
    DOLLARDE: BuiltinId.Dollarde,
    DOLLARFR: BuiltinId.Dollarfr,
    BASE: BuiltinId.Base,
    DECIMAL: BuiltinId.Decimal,
    BIN2DEC: BuiltinId.Bin2dec,
    BIN2HEX: BuiltinId.Bin2hex,
    BIN2OCT: BuiltinId.Bin2oct,
    DEC2BIN: BuiltinId.Dec2bin,
    DEC2HEX: BuiltinId.Dec2hex,
    DEC2OCT: BuiltinId.Dec2oct,
    HEX2BIN: BuiltinId.Hex2bin,
    HEX2DEC: BuiltinId.Hex2dec,
    HEX2OCT: BuiltinId.Hex2oct,
    OCT2BIN: BuiltinId.Oct2bin,
    OCT2DEC: BuiltinId.Oct2dec,
    OCT2HEX: BuiltinId.Oct2hex,
    BITAND: BuiltinId.Bitand,
    BITOR: BuiltinId.Bitor,
    BITXOR: BuiltinId.Bitxor,
    BITLSHIFT: BuiltinId.Bitlshift,
    BITRSHIFT: BuiltinId.Bitrshift,
    CONVERT: BuiltinId.Convert,
    EUROCONVERT: BuiltinId.Euroconvert,
    BESSELI: BuiltinId.Besseli,
    BESSELJ: BuiltinId.Besselj,
    BESSELK: BuiltinId.Besselk,
    BESSELY: BuiltinId.Bessely,
    VALUE: BuiltinId.Value,
    CHAR: BuiltinId.Char,
    CODE: BuiltinId.Code,
    UNICODE: BuiltinId.Unicode,
    UNICHAR: BuiltinId.Unichar,
    CLEAN: BuiltinId.Clean,
    ASC: BuiltinId.Asc,
    JIS: BuiltinId.Jis,
    DBCS: BuiltinId.Dbcs,
    BAHTTEXT: BuiltinId.Bahttext,
    SINH: BuiltinId.Sinh,
    COSH: BuiltinId.Cosh,
    TANH: BuiltinId.Tanh,
    ASINH: BuiltinId.Asinh,
    ACOSH: BuiltinId.Acosh,
    ATANH: BuiltinId.Atanh,
    ACOT: BuiltinId.Acot,
    ACOTH: BuiltinId.Acoth,
    COT: BuiltinId.Cot,
    COTH: BuiltinId.Coth,
    CSC: BuiltinId.Csc,
    CSCH: BuiltinId.Csch,
    SEC: BuiltinId.Sec,
    SECH: BuiltinId.Sech,
    SIGN: BuiltinId.Sign,
    EVEN: BuiltinId.Even,
    ODD: BuiltinId.Odd,
    FACT: BuiltinId.Fact,
    FACTDOUBLE: BuiltinId.Factdouble,
    COMBIN: BuiltinId.Combin,
    COMBINA: BuiltinId.Combina,
    GCD: BuiltinId.Gcd,
    LCM: BuiltinId.Lcm,
    PRODUCT: BuiltinId.Product,
    QUOTIENT: BuiltinId.Quotient,
    GEOMEAN: BuiltinId.Geomean,
    HARMEAN: BuiltinId.Harmean,
    SUMSQ: BuiltinId.Sumsq,
    "FLOOR.MATH": BuiltinId.FloorMath,
    "FLOOR.PRECISE": BuiltinId.FloorPrecise,
    "CEILING.MATH": BuiltinId.CeilingMath,
    "CEILING.PRECISE": BuiltinId.CeilingPrecise,
    "ISO.CEILING": BuiltinId.IsoCeiling,
    TRUNC: BuiltinId.Trunc,
    MROUND: BuiltinId.Mround,
    SQRTPI: BuiltinId.Sqrtpi,
    SERIESSUM: BuiltinId.Seriessum,
    TEXTBEFORE: BuiltinId.Textbefore,
    TEXTAFTER: BuiltinId.Textafter,
    TEXTJOIN: BuiltinId.Textjoin,
    TEXTSPLIT: BuiltinId.Textsplit,
    TEXT: BuiltinId.Text,
    PHONETIC: BuiltinId.Phonetic,
    NUMBERVALUE: BuiltinId.Numbervalue,
    VALUETOTEXT: BuiltinId.Valuetotext,
    TODAY: BuiltinId.Today,
    NOW: BuiltinId.Now,
    RAND: BuiltinId.Rand,
    MATCH: BuiltinId.Match,
    CORREL: BuiltinId.Correl,
    COVAR: BuiltinId.Covar,
    PEARSON: BuiltinId.Pearson,
    "COVARIANCE.P": BuiltinId.CovarianceP,
    "COVARIANCE.S": BuiltinId.CovarianceS,
    MEDIAN: BuiltinId.Median,
    "MODE.MULT": BuiltinId.ModeMult,
    FREQUENCY: BuiltinId.Frequency,
    PROB: BuiltinId.Prob,
    TRIMMEAN: BuiltinId.Trimmean,
    SMALL: BuiltinId.Small,
    LARGE: BuiltinId.Large,
    PERCENTILE: BuiltinId.Percentile,
    "PERCENTILE.INC": BuiltinId.PercentileInc,
    "PERCENTILE.EXC": BuiltinId.PercentileExc,
    PERCENTRANK: BuiltinId.Percentrank,
    "PERCENTRANK.INC": BuiltinId.PercentrankInc,
    "PERCENTRANK.EXC": BuiltinId.PercentrankExc,
    QUARTILE: BuiltinId.Quartile,
    "QUARTILE.INC": BuiltinId.QuartileInc,
    "QUARTILE.EXC": BuiltinId.QuartileExc,
    RANK: BuiltinId.Rank,
    "RANK.EQ": BuiltinId.RankEq,
    "RANK.AVG": BuiltinId.RankAvg,
    FORECAST: BuiltinId.Forecast,
    "FORECAST.LINEAR": BuiltinId.Forecast,
    LINEST: BuiltinId.Linest,
    LOGEST: BuiltinId.Logest,
    INTERCEPT: BuiltinId.Intercept,
    RSQ: BuiltinId.Rsq,
    SLOPE: BuiltinId.Slope,
    STEYX: BuiltinId.Steyx,
    TREND: BuiltinId.Trend,
    GROWTH: BuiltinId.Growth,
    INDEX: BuiltinId.Index,
    VLOOKUP: BuiltinId.Vlookup,
    HLOOKUP: BuiltinId.Hlookup,
    XMATCH: BuiltinId.Xmatch,
    XLOOKUP: BuiltinId.Xlookup,
    COUNTIF: BuiltinId.Countif,
    "USE.THE.COUNTIF": BuiltinId.Countif,
    COUNTIFS: BuiltinId.Countifs,
    DAVERAGE: BuiltinId.Daverage,
    DCOUNT: BuiltinId.Dcount,
    DCOUNTA: BuiltinId.Dcounta,
    DGET: BuiltinId.Dget,
    DMAX: BuiltinId.Dmax,
    DMIN: BuiltinId.Dmin,
    DPRODUCT: BuiltinId.Dproduct,
    DSTDEV: BuiltinId.Dstdev,
    DSTDEVP: BuiltinId.Dstdevp,
    DSUM: BuiltinId.Dsum,
    DVAR: BuiltinId.Dvar,
    DVARP: BuiltinId.Dvarp,
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
    "CONFIDENCE.T": BuiltinId.ConfidenceT,
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
    RATE: BuiltinId.Rate,
    NPV: BuiltinId.Npv,
    IRR: BuiltinId.Irr,
    MIRR: BuiltinId.Mirr,
    XNPV: BuiltinId.Xnpv,
    XIRR: BuiltinId.Xirr,
    IPMT: BuiltinId.Ipmt,
    PPMT: BuiltinId.Ppmt,
    ISPMT: BuiltinId.Ispmt,
    CUMIPMT: BuiltinId.Cumipmt,
    CUMPRINC: BuiltinId.Cumprinc,
    SLN: BuiltinId.Sln,
    SYD: BuiltinId.Syd,
    DISC: BuiltinId.Disc,
    INTRATE: BuiltinId.Intrate,
    RECEIVED: BuiltinId.Received,
    COUPDAYBS: BuiltinId.Coupdaybs,
    COUPDAYS: BuiltinId.Coupdays,
    COUPDAYSNC: BuiltinId.Coupdaysnc,
    COUPNCD: BuiltinId.Coupncd,
    COUPNUM: BuiltinId.Coupnum,
    COUPPCD: BuiltinId.Couppcd,
    PRICEDISC: BuiltinId.Pricedisc,
    YIELDDISC: BuiltinId.Yielddisc,
    PRICEMAT: BuiltinId.Pricemat,
    YIELDMAT: BuiltinId.Yieldmat,
    ODDFPRICE: BuiltinId.Oddfprice,
    ODDFYIELD: BuiltinId.Oddfyield,
    ODDLPRICE: BuiltinId.Oddlprice,
    ODDLYIELD: BuiltinId.Oddlyield,
    PRICE: BuiltinId.Price,
    YIELD: BuiltinId.Yield,
    DURATION: BuiltinId.Duration,
    MDURATION: BuiltinId.Mduration,
    TBILLPRICE: BuiltinId.Tbillprice,
    TBILLYIELD: BuiltinId.Tbillyield,
    TBILLEQ: BuiltinId.Tbilleq,
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
    "GAMMA.INV": BuiltinId.GammaInv,
    GAMMAINV: BuiltinId.Gammainv,
    EXPONDIST: BuiltinId.Expondist,
    "EXPON.DIST": BuiltinId.ExponDist,
    POISSON: BuiltinId.Poisson,
    "POISSON.DIST": BuiltinId.PoissonDist,
    WEIBULL: BuiltinId.Weibull,
    "WEIBULL.DIST": BuiltinId.WeibullDist,
    GAMMADIST: BuiltinId.Gammadist,
    "GAMMA.DIST": BuiltinId.GammaDist,
    CHIDIST: BuiltinId.Chidist,
    "LEGACY.CHIDIST": BuiltinId.LegacyChidist,
    CHIINV: BuiltinId.Chiinv,
    "CHISQ.DIST.RT": BuiltinId.ChisqDistRt,
    "CHISQ.DIST": BuiltinId.ChisqDist,
    "CHISQ.INV.RT": BuiltinId.ChisqInvRt,
    "CHISQ.INV": BuiltinId.ChisqInv,
    CHISQDIST: BuiltinId.Chisqdist,
    CHISQINV: BuiltinId.Chisqinv,
    "LEGACY.CHIINV": BuiltinId.LegacyChiinv,
    "BETA.DIST": BuiltinId.BetaDist,
    "BETA.INV": BuiltinId.BetaInv,
    BETADIST: BuiltinId.Betadist,
    BETAINV: BuiltinId.Betainv,
    "F.DIST": BuiltinId.FDist,
    "F.DIST.RT": BuiltinId.FDistRt,
    "F.INV": BuiltinId.FInv,
    "F.INV.RT": BuiltinId.FInvRt,
    FDIST: BuiltinId.Fdist,
    FINV: BuiltinId.Finv,
    "LEGACY.FDIST": BuiltinId.LegacyFdist,
    "LEGACY.FINV": BuiltinId.LegacyFinv,
    "T.DIST": BuiltinId.TDist,
    "T.DIST.RT": BuiltinId.TDistRt,
    "T.DIST.2T": BuiltinId.TDist2T,
    "T.INV": BuiltinId.TInv,
    "T.INV.2T": BuiltinId.TInv2T,
    TDIST: BuiltinId.Tdist,
    TINV: BuiltinId.Tinv,
    "CHISQ.TEST": BuiltinId.ChisqTest,
    CHITEST: BuiltinId.Chitest,
    "LEGACY.CHITEST": BuiltinId.LegacyChitest,
    "F.TEST": BuiltinId.FTest,
    FTEST: BuiltinId.Ftest,
    "T.TEST": BuiltinId.TTest,
    TTEST: BuiltinId.Ttest,
    "Z.TEST": BuiltinId.ZTest,
    ZTEST: BuiltinId.Ztest,
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
