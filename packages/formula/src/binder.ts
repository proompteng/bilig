import { BuiltinId, FormulaMode, MAX_WASM_RANGE_CELLS } from "@bilig/protocol";
import type { FormulaNode } from "./ast.js";
import {
  getNativeAxisAggregateCode,
  getNativeRunningFoldCode,
  isNativeMakearraySumLambda,
  isWasmSafeBuiltinArity,
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
