import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag, type CellValue } from '@bilig/protocol'
import { createKernel } from '../index.js'

const BUILTIN = {
  CONCAT: BuiltinId.Concat,
  LEN: BuiltinId.Len,
  ISBLANK: BuiltinId.IsBlank,
  ISNUMBER: BuiltinId.IsNumber,
  ISTEXT: BuiltinId.IsText,
  DATE: BuiltinId.Date,
  YEAR: BuiltinId.Year,
  MONTH: BuiltinId.Month,
  DAY: BuiltinId.Day,
  EDATE: BuiltinId.Edate,
  EOMONTH: BuiltinId.Eomonth,
  EXACT: BuiltinId.Exact,
  INT: BuiltinId.Int,
  ROUNDUP: BuiltinId.RoundUp,
  ROUNDDOWN: BuiltinId.RoundDown,
  TIME: BuiltinId.Time,
  HOUR: BuiltinId.Hour,
  MINUTE: BuiltinId.Minute,
  SECOND: BuiltinId.Second,
  WEEKDAY: BuiltinId.Weekday,
  DAYS: BuiltinId.Days,
  DAYS360: BuiltinId.Days360,
  WORKDAY: BuiltinId.Workday,
  NETWORKDAYS: BuiltinId.Networkdays,
  WEEKNUM: BuiltinId.Weeknum,
  ISOWEEKNUM: BuiltinId.Isoweeknum,
  TIMEVALUE: BuiltinId.Timevalue,
  TODAY: BuiltinId.Today,
  NOW: BuiltinId.Now,
  RAND: BuiltinId.Rand,
  WORKDAY_INTL: BuiltinId.WorkdayIntl,
  NETWORKDAYS_INTL: BuiltinId.NetworkdaysIntl,
  FILTER: BuiltinId.Filter,
  UNIQUE: BuiltinId.Unique,
  BYROW_SUM: BuiltinId.ByrowSum,
  BYCOL_SUM: BuiltinId.BycolSum,
  REDUCE_SUM: BuiltinId.ReduceSum,
  SCAN_SUM: BuiltinId.ScanSum,
  REDUCE_PRODUCT: BuiltinId.ReduceProduct,
  SCAN_PRODUCT: BuiltinId.ScanProduct,
  MAKEARRAY_SUM: BuiltinId.MakearraySum,
  BYROW_AGGREGATE: BuiltinId.ByrowAggregate,
  BYCOL_AGGREGATE: BuiltinId.BycolAggregate,
  LEFT: BuiltinId.Left,
  RIGHT: BuiltinId.Right,
  MID: BuiltinId.Mid,
  TRIM: BuiltinId.Trim,
  UPPER: BuiltinId.Upper,
  LOWER: BuiltinId.Lower,
  REPLACE: BuiltinId.Replace,
  SUBSTITUTE: BuiltinId.Substitute,
  REPT: BuiltinId.Rept,
  FIND: BuiltinId.Find,
  SEARCH: BuiltinId.Search,
  VALUE: BuiltinId.Value,
  NUMBERVALUE: BuiltinId.Numbervalue,
  VALUETOTEXT: BuiltinId.Valuetotext,
  LEFTB: BuiltinId.Leftb,
  MIDB: BuiltinId.Midb,
  RIGHTB: BuiltinId.Rightb,
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
  CHAR: BuiltinId.Char,
  CODE: BuiltinId.Code,
  UNICODE: BuiltinId.Unicode,
  UNICHAR: BuiltinId.Unichar,
  CLEAN: BuiltinId.Clean,
  ASC: BuiltinId.Asc,
  JIS: BuiltinId.Jis,
  DBCS: BuiltinId.Dbcs,
  BAHTTEXT: BuiltinId.Bahttext,
  TEXT: BuiltinId.Text,
  PHONETIC: BuiltinId.Phonetic,
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
  CHOOSE: BuiltinId.Choose,
  TEXTBEFORE: BuiltinId.Textbefore,
  TEXTAFTER: BuiltinId.Textafter,
  TEXTJOIN: BuiltinId.Textjoin,
  TEXTSPLIT: BuiltinId.Textsplit,
  T: BuiltinId.T,
  N: BuiltinId.N,
  TYPE: BuiltinId.Type,
  DELTA: BuiltinId.Delta,
  GESTEP: BuiltinId.Gestep,
  GAUSS: BuiltinId.Gauss,
  PHI: BuiltinId.Phi,
  STANDARDIZE: BuiltinId.Standardize,
  STDEV: BuiltinId.Stdev,
  STDEV_P: BuiltinId.StdevP,
  STDEV_S: BuiltinId.StdevS,
  STDEVA: BuiltinId.Stdeva,
  STDEVP: BuiltinId.Stdevp,
  STDEVPA: BuiltinId.Stdevpa,
  VAR: BuiltinId.Var,
  VAR_P: BuiltinId.VarP,
  VAR_S: BuiltinId.VarS,
  VARA: BuiltinId.Vara,
  VARP: BuiltinId.Varp,
  VARPA: BuiltinId.Varpa,
  SKEW: BuiltinId.Skew,
  SKEW_P: BuiltinId.SkewP,
  KURT: BuiltinId.Kurt,
  NORMDIST: BuiltinId.Normdist,
  NORM_DIST: BuiltinId.NormDist,
  NORMINV: BuiltinId.Norminv,
  NORM_INV: BuiltinId.NormInv,
  NORMSDIST: BuiltinId.Normsdist,
  NORM_S_DIST: BuiltinId.NormSDist,
  NORMSINV: BuiltinId.Normsinv,
  NORM_S_INV: BuiltinId.NormSInv,
  LOGINV: BuiltinId.Loginv,
  LOGNORMDIST: BuiltinId.Lognormdist,
  LOGNORM_DIST: BuiltinId.LognormDist,
  LOGNORM_INV: BuiltinId.LognormInv,
  BITAND: BuiltinId.Bitand,
  BITOR: BuiltinId.Bitor,
  BITXOR: BuiltinId.Bitxor,
  BITLSHIFT: BuiltinId.Bitlshift,
  BITRSHIFT: BuiltinId.Bitrshift,
  CONVERT: BuiltinId.Convert,
  EUROCONVERT: BuiltinId.Euroconvert,
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
  PERMUT: BuiltinId.Permut,
  PERMUTATIONA: BuiltinId.Permutationa,
  GCD: BuiltinId.Gcd,
  LCM: BuiltinId.Lcm,
  PRODUCT: BuiltinId.Product,
  QUOTIENT: BuiltinId.Quotient,
  MROUND: BuiltinId.Mround,
  GEOMEAN: BuiltinId.Geomean,
  HARMEAN: BuiltinId.Harmean,
  SUMSQ: BuiltinId.Sumsq,
  FLOOR_MATH: BuiltinId.FloorMath,
  FLOOR_PRECISE: BuiltinId.FloorPrecise,
  CEILING_MATH: BuiltinId.CeilingMath,
  CEILING_PRECISE: BuiltinId.CeilingPrecise,
  ISO_CEILING: BuiltinId.IsoCeiling,
  TRUNC: BuiltinId.Trunc,
  SQRTPI: BuiltinId.Sqrtpi,
  SERIESSUM: BuiltinId.Seriessum,
  BESSELI: BuiltinId.Besseli,
  BESSELJ: BuiltinId.Besselj,
  BESSELK: BuiltinId.Besselk,
  BESSELY: BuiltinId.Bessely,
  NA: BuiltinId.Na,
  IFS: BuiltinId.Ifs,
  IFERROR: BuiltinId.Iferror,
  IFNA: BuiltinId.Ifna,
  SWITCH: BuiltinId.Switch,
  XOR: BuiltinId.Xor,
  COUNTIF: BuiltinId.Countif,
  COUNTIFS: BuiltinId.Countifs,
  SUMIF: BuiltinId.Sumif,
  SUMIFS: BuiltinId.Sumifs,
  AVERAGEIF: BuiltinId.Averageif,
  AVERAGEIFS: BuiltinId.Averageifs,
  SUMPRODUCT: BuiltinId.Sumproduct,
  MATCH: BuiltinId.Match,
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
  INDEX: BuiltinId.Index,
  VLOOKUP: BuiltinId.Vlookup,
  HLOOKUP: BuiltinId.Hlookup,
  XMATCH: BuiltinId.Xmatch,
  XLOOKUP: BuiltinId.Xlookup,
  OFFSET: BuiltinId.Offset,
  TAKE: BuiltinId.Take,
  DROP: BuiltinId.Drop,
  EXPAND: BuiltinId.Expand,
  TRIMRANGE: BuiltinId.Trimrange,
  DATEDIF: BuiltinId.Datedif,
  FV: BuiltinId.Fv,
  FVSCHEDULE: BuiltinId.Fvschedule,
  PV: BuiltinId.Pv,
  PMT: BuiltinId.Pmt,
  NPER: BuiltinId.Nper,
  NPV: BuiltinId.Npv,
  RATE: BuiltinId.Rate,
  IPMT: BuiltinId.Ipmt,
  PPMT: BuiltinId.Ppmt,
  ISPMT: BuiltinId.Ispmt,
  CUMIPMT: BuiltinId.Cumipmt,
  CUMPRINC: BuiltinId.Cumprinc,
  DB: BuiltinId.Db,
  DDB: BuiltinId.Ddb,
  VDB: BuiltinId.Vdb,
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
  EFFECT: BuiltinId.Effect,
  NOMINAL: BuiltinId.Nominal,
  PDURATION: BuiltinId.Pduration,
  RRI: BuiltinId.Rri,
  IRR: BuiltinId.Irr,
  MIRR: BuiltinId.Mirr,
  XNPV: BuiltinId.Xnpv,
  XIRR: BuiltinId.Xirr,
  YEARFRAC: BuiltinId.Yearfrac,
  COUNTBLANK: BuiltinId.Countblank,
  CHOOSECOLS: BuiltinId.Choosecols,
  CHOOSEROWS: BuiltinId.Chooserows,
  CORREL: BuiltinId.Correl,
  COVAR: BuiltinId.Covar,
  PEARSON: BuiltinId.Pearson,
  COVARIANCE_P: BuiltinId.CovarianceP,
  COVARIANCE_S: BuiltinId.CovarianceS,
  FORECAST: BuiltinId.Forecast,
  GROWTH: BuiltinId.Growth,
  INTERCEPT: BuiltinId.Intercept,
  LINEST: BuiltinId.Linest,
  LOGEST: BuiltinId.Logest,
  MEDIAN: BuiltinId.Median,
  MODE: BuiltinId.Mode,
  MODE_SNGL: BuiltinId.ModeSngl,
  MODE_MULT: BuiltinId.ModeMult,
  FREQUENCY: BuiltinId.Frequency,
  SMALL: BuiltinId.Small,
  LARGE: BuiltinId.Large,
  PERCENTILE: BuiltinId.Percentile,
  PERCENTILE_INC: BuiltinId.PercentileInc,
  PERCENTILE_EXC: BuiltinId.PercentileExc,
  PERCENTRANK: BuiltinId.Percentrank,
  PERCENTRANK_INC: BuiltinId.PercentrankInc,
  PERCENTRANK_EXC: BuiltinId.PercentrankExc,
  PROB: BuiltinId.Prob,
  QUARTILE: BuiltinId.Quartile,
  QUARTILE_INC: BuiltinId.QuartileInc,
  QUARTILE_EXC: BuiltinId.QuartileExc,
  RANK: BuiltinId.Rank,
  RANK_EQ: BuiltinId.RankEq,
  RANK_AVG: BuiltinId.RankAvg,
  RSQ: BuiltinId.Rsq,
  SLOPE: BuiltinId.Slope,
  STEYX: BuiltinId.Steyx,
  TREND: BuiltinId.Trend,
  TRIMMEAN: BuiltinId.Trimmean,
  SORT: BuiltinId.Sort,
  SORTBY: BuiltinId.Sortby,
  TOCOL: BuiltinId.Tocol,
  TOROW: BuiltinId.Torow,
  WRAPROWS: BuiltinId.Wraprows,
  WRAPCOLS: BuiltinId.Wrapcols,
  ERF: BuiltinId.Erf,
  ERFC: BuiltinId.Erfc,
  FISHER: BuiltinId.Fisher,
  FISHERINV: BuiltinId.Fisherinv,
  GAMMALN: BuiltinId.Gammaln,
  GAMMA: BuiltinId.Gamma,
  GAMMA_INV: BuiltinId.GammaInv,
  GAMMAINV: BuiltinId.Gammainv,
  CONFIDENCE_NORM: BuiltinId.ConfidenceNorm,
  CONFIDENCE: BuiltinId.Confidence,
  CONFIDENCE_T: BuiltinId.ConfidenceT,
  EXPONDIST: BuiltinId.Expondist,
  POISSON: BuiltinId.Poisson,
  WEIBULL: BuiltinId.Weibull,
  GAMMADIST: BuiltinId.Gammadist,
  CHIDIST: BuiltinId.Chidist,
  LEGACY_CHIDIST: BuiltinId.LegacyChidist,
  CHIINV: BuiltinId.Chiinv,
  CHISQ_INV_RT: BuiltinId.ChisqInvRt,
  CHISQ_INV: BuiltinId.ChisqInv,
  CHISQ_DIST: BuiltinId.ChisqDist,
  CHISQDIST: BuiltinId.Chisqdist,
  CHISQINV: BuiltinId.Chisqinv,
  LEGACY_CHIINV: BuiltinId.LegacyChiinv,
  CHISQ_TEST: BuiltinId.ChisqTest,
  CHITEST: BuiltinId.Chitest,
  LEGACY_CHITEST: BuiltinId.LegacyChitest,
  F_TEST: BuiltinId.FTest,
  FTEST: BuiltinId.Ftest,
  Z_TEST: BuiltinId.ZTest,
  ZTEST: BuiltinId.Ztest,
  BETA_DIST: BuiltinId.BetaDist,
  BETA_INV: BuiltinId.BetaInv,
  BETADIST: BuiltinId.Betadist,
  BETAINV: BuiltinId.Betainv,
  F_DIST: BuiltinId.FDist,
  F_DIST_RT: BuiltinId.FDistRt,
  F_INV: BuiltinId.FInv,
  F_INV_RT: BuiltinId.FInvRt,
  FDIST: BuiltinId.Fdist,
  FINV: BuiltinId.Finv,
  LEGACY_FDIST: BuiltinId.LegacyFdist,
  LEGACY_FINV: BuiltinId.LegacyFinv,
  T_DIST: BuiltinId.TDist,
  T_DIST_RT: BuiltinId.TDistRt,
  T_DIST_2T: BuiltinId.TDist2T,
  T_INV: BuiltinId.TInv,
  T_INV_2T: BuiltinId.TInv2T,
  TDIST: BuiltinId.Tdist,
  TINV: BuiltinId.Tinv,
  T_TEST: BuiltinId.TTest,
  TTEST: BuiltinId.Ttest,
  BINOMDIST: BuiltinId.Binomdist,
  BINOM_DIST_RANGE: BuiltinId.BinomDistRange,
  CRITBINOM: BuiltinId.Critbinom,
  HYPGEOMDIST: BuiltinId.Hypgeomdist,
  NEGBINOMDIST: BuiltinId.Negbinomdist,
} as const

const OUTPUT_STRING_BASE = 2147483648

function asciiCodes(text: string): Uint16Array {
  return Uint16Array.from(Array.from(text, (char) => char.charCodeAt(0)))
}

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushCell(cellOffset: number): number {
  return (Opcode.PushCell << 24) | cellOffset
}

function encodePushRange(rangeIndex: number): number {
  return (Opcode.PushRange << 24) | rangeIndex
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId
}

function encodePushBoolean(value: boolean): number {
  return (Opcode.PushBoolean << 24) | (value ? 1 : 0)
}

function encodePushError(code: ErrorCode): number {
  return (Opcode.PushError << 24) | code
}

function encodeBinary(opcode: Opcode): number {
  return opcode << 24
}

function encodeRet(): number {
  return Opcode.Ret << 24
}

function packPrograms(programs: number[][]): {
  programs: Uint32Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0

  for (const program of programs) {
    offsets.push(offset)
    lengths.push(program.length)
    flat.push(...program)
    offset += program.length
  }

  return {
    programs: Uint32Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function packConstants(constantsByProgram: number[][]): {
  constants: Float64Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0

  for (const constants of constantsByProgram) {
    offsets.push(offset)
    lengths.push(constants.length)
    flat.push(...constants)
    offset += constants.length
  }

  return {
    constants: Float64Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

type KernelInstance = Awaited<ReturnType<typeof createKernel>>

function decodeValueTag(rawTag: number): ValueTag {
  switch (rawTag) {
    case 0:
      return ValueTag.Empty
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    default:
      throw new Error(`Unexpected spill tag: ${rawTag}`)
  }
}

function decodeErrorCode(rawCode: number): ErrorCode {
  switch (rawCode) {
    case 0:
      return ErrorCode.None
    case 1:
      return ErrorCode.Div0
    case 2:
      return ErrorCode.Ref
    case 3:
      return ErrorCode.Value
    case 4:
      return ErrorCode.Name
    case 5:
      return ErrorCode.NA
    case 6:
      return ErrorCode.Cycle
    case 7:
      return ErrorCode.Spill
    case 8:
      return ErrorCode.Blocked
    default:
      throw new Error(`Unexpected error code: ${rawCode}`)
  }
}

function readSpillValues(kernel: KernelInstance, ownerCellIndex: number, pooledStrings: readonly string[]): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0
  const tags = kernel.readSpillTags()
  const values = kernel.readSpillNumbers()
  const outputStrings = kernel.readOutputStrings()
  return Array.from({ length }, (_, index) => {
    const tag = decodeValueTag(tags[offset + index] ?? ValueTag.Empty)
    const rawValue = values[offset + index] ?? 0
    switch (tag) {
      case ValueTag.Number:
        return { tag, value: rawValue }
      case ValueTag.Boolean:
        return { tag, value: rawValue !== 0 }
      case ValueTag.Empty:
        return { tag }
      case ValueTag.Error:
        return { tag, code: decodeErrorCode(rawValue) }
      case ValueTag.String: {
        const outputIndex = rawValue >= OUTPUT_STRING_BASE ? rawValue - OUTPUT_STRING_BASE : -1
        return {
          tag,
          value: outputIndex >= 0 ? (outputStrings[outputIndex] ?? '') : (pooledStrings[rawValue] ?? ''),
          stringId: 0,
        }
      }
    }
    throw new Error('Unexpected decoded spill tag')
  })
}

function expectNumberCell(kernel: KernelInstance, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

function expectErrorCell(kernel: KernelInstance, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}

function expectEmptyCell(kernel: KernelInstance, index: number): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Empty)
}

describe('wasm kernel', () => {
  it('evaluates a simple program batch', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 4, 4, 4)
    kernel.writeCells(new Uint8Array([1, 0, 0, 0]), new Float64Array([10, 0, 0, 0]), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([(3 << 24) | 0, (1 << 24) | 0, 7 << 24, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([4]),
      new Uint32Array([1]),
    )
    kernel.uploadConstants(new Float64Array([2]), new Uint32Array([0]), new Uint32Array([1]))
    kernel.evalBatch(new Uint32Array([1]))
    expect(kernel.readNumbers()[1]).toBe(20)
    expect(kernel.readConstantOffsets()[0]).toBe(0)
    expect(kernel.readConstantLengths()[0]).toBe(1)
    expect(kernel.readConstants()[0]).toBe(2)
  })

  it('evaluates aggregate and numeric builtins', async () => {
    const kernel = await createKernel()
    kernel.init(6, 6, 2, 6, 6)
    kernel.writeCells(new Uint8Array([1, 1, 0, 0, 0, 0]), new Float64Array([2, 3, 0, 0, 0, 0]), new Uint32Array(6), new Uint16Array(6))
    kernel.uploadPrograms(
      new Uint32Array([(3 << 24) | 0, (3 << 24) | 1, (20 << 24) | (1 << 8) | 2, (1 << 24) | 0, 5 << 24, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([2]),
    )
    kernel.uploadConstants(new Float64Array([4]), new Uint32Array([0]), new Uint32Array([1]))

    kernel.evalBatch(new Uint32Array([2]))
    expect(kernel.readNumbers()[2]).toBe(9)
  })

  it('evaluates branch programs with jump opcodes', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 4, 4, 4)
    kernel.writeCells(new Uint8Array([2, 0, 0, 0]), new Float64Array([1, 0, 0, 0]), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([(3 << 24) | 0, (19 << 24) | 4, (1 << 24) | 0, (18 << 24) | 5, (1 << 24) | 1, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([1]),
    )
    kernel.uploadConstants(new Float64Array([10, 20]), new Uint32Array([0]), new Uint32Array([2]))

    kernel.evalBatch(new Uint32Array([1]))
    expect(kernel.readNumbers()[1]).toBe(10)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    const errors = kernel.readErrors()
    kernel.writeCells(
      new Uint8Array([2, tags[1], 0, 0]),
      new Float64Array([0, numbers[1], 0, 0]),
      new Uint32Array(4),
      new Uint16Array([0, errors[1], 0, 0]),
    )
    kernel.evalBatch(new Uint32Array([1]))
    expect(kernel.readNumbers()[1]).toBe(20)
  })

  it('evaluates aggregate builtins through uploaded range members', async () => {
    const kernel = await createKernel()
    kernel.init(6, 6, 1, 4, 4)
    kernel.writeCells(new Uint8Array([1, 1, 0, 0, 0, 0]), new Float64Array([2, 3, 0, 0, 0, 0]), new Uint32Array(6), new Uint16Array(6))
    kernel.uploadPrograms(
      new Uint32Array([(4 << 24) | 0, (20 << 24) | (1 << 8) | 1, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([3]),
      new Uint32Array([2]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))
    kernel.uploadRangeShapes(new Uint32Array([2]), new Uint32Array([1]))

    kernel.evalBatch(new Uint32Array([2]))

    expect(kernel.readNumbers()[2]).toBe(5)
    expect(kernel.readRangeLengths()[0]).toBe(2)
    expect(kernel.readRangeMembers()[1]).toBe(1)
  })

  it('evaluates exact-safe logical info builtins with zero-arg, scalar, and range cases', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 2, 2, 2)
    kernel.writeCells(
      new Uint8Array([0, 1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 42, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(16),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))

    const packed = packPrograms([
      [encodeCall(BUILTIN.ISBLANK, 0), encodeRet()],
      [encodeCall(BUILTIN.ISNUMBER, 0), encodeRet()],
      [encodeCall(BUILTIN.ISTEXT, 0), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.ISBLANK, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.ISNUMBER, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.ISTEXT, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.ISNUMBER, 1), encodeRet()],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 7, width)]).toBe(ErrorCode.Value)
  })

  it('evaluates LEN with scalar coercion and range rejection', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 1, 1, 2)
    kernel.uploadStringLengths(Uint32Array.from([0, 0, 0, 0, 0, 0, 0, 0, 0, 5]))
    kernel.writeCells(
      new Uint8Array([0, 2, 1, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 1, 123.45, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 9, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array([0, 0, 0, 0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))

    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.LEN, 1), encodeRet()],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(4)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(6)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(5)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 5, width)]).toBe(ErrorCode.Ref)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 6, width)]).toBe(ErrorCode.Value)
  })

  it('evaluates EXACT and numeric rounding builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 2, 1, 2)
    kernel.uploadStrings(Uint32Array.from([0, 0, 5, 10]), Uint32Array.from([0, 5, 5, 5]), asciiCodes('AlphaAlphaalpha'))
    kernel.writeCells(
      new Uint8Array([3, 3, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 0, -3.145, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([1, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushCell(1), encodeCall(BUILTIN.EXACT, 2), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.INT, 1), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.ROUNDUP, 2), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.ROUNDDOWN, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([2]), new Uint32Array([0, 0, 0, 0]), new Uint32Array([0, 0, 1, 1]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(-3.15)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(-3.14)
  })

  it('evaluates string literals and CONCAT through the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 2]), Uint32Array.from([0, 2, 3]), asciiCodes('xyfoo'))
    kernel.writeCells(
      new Uint8Array([ValueTag.String, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array(16),
      new Uint32Array([2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(16),
    )
    const packed = packPrograms([
      [encodePushString(1), encodeRet()],
      [encodePushString(1), encodePushCell(0), encodeCall(BUILTIN.CONCAT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['xyfoo'])
  })

  it('evaluates binary text comparison and concat operators on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 5, 10, 11]), Uint32Array.from([0, 5, 5, 1, 1]), asciiCodes('helloHELLObA'))
    kernel.writeCells(
      new Uint8Array([ValueTag.String, ValueTag.String, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array(24),
      new Uint32Array([1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushCell(1), encodeBinary(Opcode.Eq), encodeRet()],
      [encodePushString(3), encodePushString(4), encodeBinary(Opcode.Gt), encodeRet()],
      [encodePushCell(0), encodePushString(4), encodeBinary(Opcode.Concat), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0]), new Uint32Array([0, 0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['helloA'])
  })

  it('evaluates text slicing, casing, and search builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(40, 8, 2, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 5, 21, 26, 30, 38, 40]),
      Uint32Array.from([0, 5, 16, 5, 4, 8, 2, 2]),
      asciiCodes('Alpha  alpha   beta  alphaBETAalphabetphP*'),
    )
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array(40),
      new Uint32Array([
        1, 2, 3, 4, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint16Array(40),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushNumber(0), encodeCall(BUILTIN.LEFT, 2), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.RIGHT, 1), encodeRet()],
      [encodePushCell(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MID, 3), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.TRIM, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.UPPER, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.LOWER, 1), encodeRet()],
      [encodePushString(6), encodePushCell(4), encodeCall(BUILTIN.FIND, 2), encodeRet()],
      [encodePushString(7), encodePushCell(4), encodeCall(BUILTIN.SEARCH, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array([2, 2, 3]), new Uint32Array([0, 1, 1, 3, 3, 3, 3]), new Uint32Array([1, 0, 2, 0, 0, 0, 0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Number)
    expect(kernel.readTags()[cellIndex(1, 8, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(3)
    expect(kernel.readOutputStrings()).toEqual(['Al', 'a', 'lph', 'alpha beta', 'ALPHA', 'beta'])
  })

  it('evaluates REPLACE, SUBSTITUTE, and REPT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 8, 9, 15, 17, 19]),
      Uint32Array.from([0, 8, 1, 6, 2, 2, 2]),
      asciiCodes('alphabetZbananaanooxo'),
    )
    kernel.writeCells(
      new Uint8Array([ValueTag.String, ValueTag.String, ValueTag.String, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array(24),
      new Uint32Array([1, 3, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushNumber(0), encodePushNumber(1), encodePushString(2), encodeCall(BUILTIN.REPLACE, 4), encodeRet()],
      [encodePushCell(1), encodePushString(4), encodePushString(5), encodeCall(BUILTIN.SUBSTITUTE, 3), encodeRet()],
      [encodePushCell(1), encodePushString(4), encodePushString(5), encodePushNumber(0), encodeCall(BUILTIN.SUBSTITUTE, 4), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.REPT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([3, 2, 2, 3]), new Uint32Array([0, 0, 2, 3]), new Uint32Array([2, 0, 1, 1]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['alZabet', 'booooa', 'banooa', 'xoxoxo'])
  })

  it('evaluates CHOOSE, TEXTBEFORE, TEXTAFTER, and TEXTJOIN on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(32, 7, 2, 1, 1)
    const pooledStrings = ['alpha-beta', '-', 'alpha', 'beta'] as const
    kernel.uploadStrings(Uint32Array.from([0, 0, 10, 11, 16]), Uint32Array.from([0, 10, 1, 5, 4]), asciiCodes('alpha-beta-alphabeta'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.Empty,
        ValueTag.String,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([10, 20, 30, 40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 3, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6]), Uint32Array.from([0, 4]), Uint32Array.from([4, 3]))
    kernel.uploadRangeShapes(Uint32Array.from([2, 3]), Uint32Array.from([2, 1]))

    const packed = packPrograms([
      [encodePushString(1), encodePushString(2), encodeCall(BUILTIN.TEXTBEFORE, 2), encodeRet()],
      [encodePushString(1), encodePushString(2), encodeCall(BUILTIN.TEXTAFTER, 2), encodeRet()],
      [encodePushString(2), encodePushBoolean(true), encodePushRange(1), encodeCall(BUILTIN.TEXTJOIN, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.CHOOSE, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    const constants = packConstants([[], [], [], [1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readOutputStrings()).toEqual(['alpha', 'beta', 'alpha-beta'])
    expect(kernel.readSpillRows()[cellIndex(1, 4, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(1, 4, width)]).toBe(2)
    expect(readSpillValues(kernel, cellIndex(1, 4, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.Number, value: 30 },
      { tag: ValueTag.Number, value: 40 },
    ])
  })

  it('evaluates TEXTSPLIT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 14, 15]), Uint32Array.from([0, 14, 1, 1]), asciiCodes('red,blue|green,|'))
    const packed = packPrograms([
      [encodePushString(1), encodePushString(2), encodePushString(3), encodeCall(BUILTIN.TEXTSPLIT, 3), encodeRet()],
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([cellIndex(1, 1, width)]))
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width)]))

    expect(kernel.readSpillRows()[cellIndex(1, 1, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(1, 1, width)]).toBe(2)
    expect(readSpillValues(kernel, cellIndex(1, 1, width), [])).toEqual([
      { tag: ValueTag.String, value: 'red', stringId: 0 },
      { tag: ValueTag.String, value: 'blue', stringId: 0 },
      { tag: ValueTag.String, value: 'green', stringId: 0 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
    ])
  })

  it('evaluates VALUE for dynamic scalar inputs on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 4, 16]), Uint32Array.from([0, 4, 12, 3]), asciiCodes('42.5  -17.25e1  not'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Boolean,
        ValueTag.Empty,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0, 0, 0]), new Uint32Array([0, 0, 0, 0, 0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(42.5)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-172.5)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(0)
  })

  it('evaluates NUMBERVALUE and VALUETOTEXT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 8, 9, 10]), Uint32Array.from([8, 1, 1, 5]), asciiCodes('2.500,27,.alpha'))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))
    const packed = packPrograms([
      [encodePushString(0), encodePushString(1), encodePushString(2), encodeCall(BUILTIN.NUMBERVALUE, 3), encodeRet()],
      [encodePushString(3), encodePushNumber(0), encodeCall(BUILTIN.VALUETOTEXT, 2), encodeRet()],
      [encodePushError(ErrorCode.Ref), encodeCall(BUILTIN.VALUETOTEXT, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    const constants = packConstants([[], [1], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(2500.27, 12)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['"alpha"', '#REF!'])
  })

  it('evaluates TEXT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 6, 1, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 8, 18, 28, 36]),
      Uint32Array.from([8, 10, 10, 8, 5]),
      asciiCodes('#,##0.00yyyy-mm-ddh:mm AM/PMprefix @alpha'),
    )
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))
    const packed = packPrograms([
      [encodePushNumber(0), encodePushString(0), encodeCall(BUILTIN.TEXT, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(1), encodeCall(BUILTIN.TEXT, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(2), encodeCall(BUILTIN.TEXT, 2), encodeRet()],
      [encodePushString(4), encodePushString(3), encodeCall(BUILTIN.TEXT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    const constants = packConstants([[1234.567], [45356], [0.5], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['1,234.57', '2024-03-05', '12:00 PM', 'prefix alpha'])
  })

  it('evaluates scalar text conversion builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 6, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 1, 6]), Uint32Array.from([1, 5, 2]), Uint16Array.from([65, 97, 1, 98, 127, 99, 54, 54]))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BUILTIN.CHAR, 1), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.CODE, 1), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.UNICODE, 1), encodeRet()],
      [encodePushString(2), encodeCall(BUILTIN.UNICHAR, 1), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.CLEAN, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    )
    const constants = packConstants([[65], [], [], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(1, 1, width), 65)
    expectNumberCell(kernel, cellIndex(1, 2, width), 65)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['A', 'B', 'abc'])
  })

  it('evaluates ASC, JIS, and DBCS on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    const strings = ['ＡＢＣ　１２３', 'ABC 123', 'ｶﾞｷﾞｸﾞｹﾞｺﾞ']
    const offsets = []
    const lengths = []
    const data = []
    let cursor = 0
    for (const value of strings) {
      const codes = Array.from(value, (char) => char.charCodeAt(0))
      offsets.push(cursor)
      lengths.push(codes.length)
      data.push(...codes)
      cursor += codes.length
    }
    kernel.uploadStrings(Uint32Array.from(offsets), Uint32Array.from(lengths), Uint16Array.from(data))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BUILTIN.ASC, 1), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.JIS, 1), encodeRet()],
      [encodePushString(2), encodeCall(BUILTIN.DBCS, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    const constants = packConstants([[], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]))
    const outputStrings = kernel.readOutputStrings()
    const readStringCell = (index: number): string => {
      expect(kernel.readTags()[index]).toBe(ValueTag.String)
      const raw = kernel.readStringIds()[index] ?? 0
      const outputIndex = raw >= OUTPUT_STRING_BASE ? raw - OUTPUT_STRING_BASE : -1
      return outputIndex >= 0 ? (outputStrings[outputIndex] ?? '') : (strings[raw] ?? '')
    }

    expect(readStringCell(cellIndex(1, 0, width))).toBe('ABC 123')
    expect(readStringCell(cellIndex(1, 1, width))).toBe('ＡＢＣ　１２３')
    expect(readStringCell(cellIndex(1, 2, width))).toBe('ガギグゲゴ')
  })

  it('evaluates BAHTTEXT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 2, 0, 0, 0)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BUILTIN.BAHTTEXT, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.BAHTTEXT, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    )
    const constants = packConstants([[1234], [21.25]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน', 'ยี่สิบเอ็ดบาทยี่สิบห้าสตางค์'])
  })

  it('evaluates PHONETIC on the wasm path for scalar text values', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 2, 0, 0, 0)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([4]), Uint16Array.from([65, 66, 67, 68]))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
      ]),
      new Float64Array(8),
      new Uint32Array([0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(8),
    )

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BUILTIN.PHONETIC, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.PHONETIC, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    )
    const constants = packConstants([[], [42]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['ABCD', '42'])
  })

  it('evaluates byte-oriented text builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 8, 4, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 6, 14, 16, 17, 18]),
      Uint32Array.from([6, 8, 2, 1, 1, 1]),
      Uint16Array.from(Array.from('abcdefalphabetphdZé', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(20), new Float64Array(20), new Uint32Array(20), new Uint16Array(20))

    const packed = packPrograms([
      [encodePushString(5), encodeCall(BUILTIN.LENB, 1), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.LEFTB, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MIDB, 3), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.RIGHTB, 2), encodeRet()],
      [encodePushString(3), encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.FINDB, 3), encodeRet()],
      [encodePushString(2), encodePushString(1), encodeCall(BUILTIN.SEARCHB, 2), encodeRet()],
      [encodePushString(1), encodePushNumber(0), encodePushNumber(1), encodePushString(4), encodeCall(BUILTIN.REPLACEB, 4), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    )
    const constants = packConstants([[], [2], [3, 2], [3], [3], [], [3, 2]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(4)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['ab', 'cd', 'def', 'alZabet'])
  })

  it('evaluates ADDRESS and dollar-format helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0]),
      Uint32Array.from([7]),
      Uint16Array.from(Array.from("O'Brien", (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.ADDRESS, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushString(0),
        encodeCall(BUILTIN.ADDRESS, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DOLLAR, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DOLLARDE, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DOLLARFR, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )
    const constants = packConstants([
      [12, 3],
      [2, 28, 3, 1],
      [-1234.5, 1],
      [1.08, 16],
      [1.5, 16],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1.5)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1.08)
    expect(kernel.readOutputStrings()).toEqual(['$C$12', "'O''Brien'!$AB2", '-$1,234.5'])
  })

  it('evaluates bitwise and base-conversion helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITAND, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITOR, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITXOR, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITLSHIFT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITRSHIFT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.BASE, 3), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.DECIMAL, 2), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.BIN2DEC, 1), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.BIN2HEX, 1), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.BIN2OCT, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DEC2BIN, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DEC2HEX, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DEC2OCT, 2), encodeRet()],
      [encodePushString(2), encodePushNumber(0), encodeCall(BUILTIN.HEX2BIN, 2), encodeRet()],
      [encodePushString(3), encodeCall(BUILTIN.HEX2DEC, 1), encodeRet()],
      [encodePushString(4), encodePushNumber(0), encodeCall(BUILTIN.HEX2OCT, 2), encodeRet()],
      [encodePushString(5), encodePushNumber(0), encodeCall(BUILTIN.OCT2BIN, 2), encodeRet()],
      [encodePushString(6), encodeCall(BUILTIN.OCT2DEC, 1), encodeRet()],
      [encodePushString(6), encodePushNumber(0), encodeCall(BUILTIN.OCT2HEX, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
        cellIndex(2, 3, width),
        cellIndex(2, 4, width),
        cellIndex(2, 5, width),
        cellIndex(2, 6, width),
        cellIndex(2, 7, width),
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
      ]),
    )
    const constants = packConstants([
      [6, 3],
      [6, 3],
      [6, 3],
      [1, 4],
      [16, 4],
      [255, 16, 4],
      [16],
      [],
      [],
      [],
      [10, 8],
      [255, 4],
      [15, 4],
      [8],
      [],
      [4],
      [8],
      [],
      [4],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.uploadStrings(
      Uint32Array.from([0, 4, 14, 15, 25, 26, 28]),
      Uint32Array.from([4, 10, 1, 10, 1, 2, 2]),
      asciiCodes('00FF1111111111AFFFFFFFFFFF1217'),
    )
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
        cellIndex(2, 3, width),
        cellIndex(2, 4, width),
        cellIndex(2, 5, width),
        cellIndex(2, 6, width),
        cellIndex(2, 7, width),
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
      ]),
    )

    expectNumberCell(kernel, cellIndex(1, 0, width), 2)
    expectNumberCell(kernel, cellIndex(1, 1, width), 7)
    expectNumberCell(kernel, cellIndex(1, 2, width), 5)
    expectNumberCell(kernel, cellIndex(1, 3, width), 16)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual([
      '00FF',
      'FFFFFFFFFF',
      '7777777777',
      '00001010',
      '00FF',
      '0017',
      '00001010',
      '0017',
      '00001010',
      '000F',
    ])
    expectNumberCell(kernel, cellIndex(1, 6, width), 255)
    expectNumberCell(kernel, cellIndex(1, 7, width), -1)
    expect(kernel.readTags()[cellIndex(2, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 5, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(2, 6, width), -1)
    expect(kernel.readTags()[cellIndex(2, 7, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(3, 1, width), 15)
    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.String)
  })

  it('evaluates Bessel engineering helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 4, 0, 1, 1)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BESSELI, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BESSELJ, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BESSELK, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BESSELY, 2), encodeRet()],
    ])
    const constants = packConstants([
      [1.5, 1],
      [1.9, 2],
      [1.5, 1],
      [2.5, 1],
    ])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.981666428, 7)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.329925728, 7)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.277387804, 7)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.145918138, 7)
  })

  it('evaluates CONVERT and EUROCONVERT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 4, 5, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 2, 4, 5, 6, 9, 12]),
      Uint32Array.from([0, 2, 2, 1, 1, 3, 3, 3]),
      asciiCodes('mikmFCDEMEURFRF'),
    )

    const packed = packPrograms([
      [encodePushNumber(0), encodePushString(1), encodePushString(2), encodeCall(BUILTIN.CONVERT, 3), encodeRet()],
      [encodePushNumber(0), encodePushString(3), encodePushString(4), encodeCall(BUILTIN.CONVERT, 3), encodeRet()],
      [encodePushNumber(0), encodePushString(5), encodePushString(6), encodeCall(BUILTIN.EUROCONVERT, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushString(7),
        encodePushString(5),
        encodePushBoolean(true),
        encodePushNumber(1),
        encodeCall(BUILTIN.EUROCONVERT, 5),
        encodeRet(),
      ],
    ])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    const constants = packConstants([[6], [68], [1.2], [1, 3]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(1, 0, width), 9.656064, 12)
    expectNumberCell(kernel, cellIndex(1, 1, width), 20, 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.61, 12)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.29728616, 12)
  })

  it('evaluates promoted scalar and reducer math helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(40, 26, 0, 4, 4)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([
        2, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint32Array(40),
      new Uint16Array(40),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1, 2]), new Uint32Array([0]), new Uint32Array([3]))
    kernel.uploadRangeShapes(new Uint32Array([3]), new Uint32Array([1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BUILTIN.ACOSH, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.COT, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.SECH, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.FACT, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.COMBIN, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.GCD, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.LCM, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.PRODUCT, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.QUOTIENT, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.GEOMEAN, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.HARMEAN, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.SUMSQ, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.EVEN, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.ODD, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.SIGN, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TRUNC, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.FLOOR_MATH, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.FLOOR_PRECISE, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CEILING_MATH, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CEILING_PRECISE, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.ISO_CEILING, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MROUND, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.SQRTPI, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.PERMUT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.PERMUTATIONA, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.SERIESSUM, 5),
        encodeRet(),
      ],
    ])
    const outputCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
      cellIndex(1, 7, width),
      cellIndex(2, 0, width),
      cellIndex(2, 1, width),
      cellIndex(2, 2, width),
      cellIndex(2, 3, width),
      cellIndex(2, 4, width),
      cellIndex(2, 5, width),
      cellIndex(2, 6, width),
      cellIndex(2, 7, width),
      cellIndex(3, 0, width),
      cellIndex(3, 1, width),
      cellIndex(3, 2, width),
      cellIndex(3, 3, width),
      cellIndex(3, 4, width),
      cellIndex(3, 5, width),
      cellIndex(3, 6, width),
      cellIndex(3, 7, width),
      cellIndex(4, 0, width),
      cellIndex(4, 1, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    const constants = packConstants([
      [1],
      [1],
      [0],
      [5],
      [8, 3],
      [],
      [],
      [],
      [7, 3],
      [],
      [],
      [],
      [-3],
      [-2],
      [-42],
      [-3.98, 1],
      [-5.5, 2],
      [-5.5, 2],
      [-5.5, 2],
      [-5.5, 2],
      [-5.5, 2],
      [10, 4],
      [2],
      [5, 3],
      [2, 3],
      [2, 1, 2, 1, 2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(1, 0, width), 0, 12)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.6420926159343306, 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 3, width), 120, 12)
    expectNumberCell(kernel, cellIndex(1, 4, width), 56, 12)
    expectNumberCell(kernel, cellIndex(1, 5, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), 12, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), 24, 12)
    expectNumberCell(kernel, cellIndex(2, 0, width), 2, 12)
    expectNumberCell(kernel, cellIndex(2, 1, width), 2.8844991406148166, 12)
    expectNumberCell(kernel, cellIndex(2, 2, width), 2.769230769230769, 12)
    expectNumberCell(kernel, cellIndex(2, 3, width), 29, 12)
    expectNumberCell(kernel, cellIndex(2, 4, width), -4, 12)
    expectNumberCell(kernel, cellIndex(2, 5, width), -3, 12)
    expectNumberCell(kernel, cellIndex(2, 6, width), -1, 12)
    expectNumberCell(kernel, cellIndex(2, 7, width), -3.9, 12)
    expectNumberCell(kernel, cellIndex(3, 0, width), -6, 12)
    expectNumberCell(kernel, cellIndex(3, 1, width), -6, 12)
    expectNumberCell(kernel, cellIndex(3, 2, width), -4, 12)
    expectNumberCell(kernel, cellIndex(3, 3, width), -4, 12)
    expectNumberCell(kernel, cellIndex(3, 4, width), -4, 12)
    expectNumberCell(kernel, cellIndex(3, 5, width), 12, 12)
    expectNumberCell(kernel, cellIndex(3, 6, width), 2.5066282746310002, 12)
    expectNumberCell(kernel, cellIndex(3, 7, width), 60, 12)
    expectNumberCell(kernel, cellIndex(4, 0, width), 8, 12)
    expectNumberCell(kernel, cellIndex(4, 1, width), 18, 12)
  })

  it('evaluates exact-parity information and threshold helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 0, 6, 6)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Boolean, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([42, 1, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(0),
      new Uint16Array(0),
    )

    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.T, 1), encodeRet()],
      [encodeCall(BUILTIN.N, 0), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.N, 1), encodeRet()],
      [encodeCall(BUILTIN.TYPE, 0), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.TYPE, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(0), encodeCall(BUILTIN.DELTA, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.GESTEP, 2), encodeRet()],
      [encodePushNumber(2), encodeCall(BUILTIN.GAUSS, 1), encodeRet()],
      [encodePushNumber(2), encodeCall(BUILTIN.PHI, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.T, 1), encodeRet()],
    ])
    const outputCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
      cellIndex(1, 7, width),
      cellIndex(2, 0, width),
      cellIndex(2, 1, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    const constants = packConstants([[4], [2], [0]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectEmptyCell(kernel, cellIndex(1, 0, width))
    expectNumberCell(kernel, cellIndex(1, 1, width), 0, 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 3, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 5, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), 0, 8)
    expectNumberCell(kernel, cellIndex(2, 0, width), 0.3989422804014327, 12)
    expectEmptyCell(kernel, cellIndex(2, 1, width))
  })

  it('evaluates IFERROR, IFNA, and NA on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 8]), Uint32Array.from([0, 8, 7]), asciiCodes('fallbackmissing'))
    kernel.writeCells(
      new Uint8Array([ValueTag.Error, ValueTag.Error, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([ErrorCode.Div0, ErrorCode.Ref, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array([ErrorCode.Div0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushString(1), encodeCall(BUILTIN.IFERROR, 2), encodeRet()],
      [encodeCall(BUILTIN.NA, 0), encodePushString(2), encodeCall(BUILTIN.IFNA, 2), encodeRet()],
      [encodePushCell(1), encodePushString(2), encodeCall(BUILTIN.IFNA, 2), encodeRet()],
      [encodePushCell(2), encodePushString(1), encodeCall(BUILTIN.IFERROR, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0, 0]), new Uint32Array([0, 0, 0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 2, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Ref)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(7)
  })

  it('evaluates conditional aggregates and SUMPRODUCT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(32, 8, 5, 1, 2)
    kernel.uploadStrings(Uint32Array.from([0, 0, 2, 3]), Uint32Array.from([0, 2, 1, 1]), asciiCodes('>0xy'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([2, 4, -1, 6, 0, 0, 0, 0, 10, 20, 30, 40, 1, 2, 3, 4, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 2, 2, 3, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(
      new Uint32Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]),
      Uint32Array.from([0, 4, 8, 12, 15]),
      Uint32Array.from([4, 4, 4, 3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 4, 3, 3]), Uint32Array.from([1, 1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushString(1), encodeCall(BUILTIN.COUNTIF, 2), encodeRet()],
      [encodePushRange(0), encodePushString(1), encodePushRange(1), encodePushString(2), encodeCall(BUILTIN.COUNTIFS, 4), encodeRet()],
      [encodePushRange(0), encodePushString(1), encodePushRange(2), encodeCall(BUILTIN.SUMIF, 3), encodeRet()],
      [
        encodePushRange(2),
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodePushString(2),
        encodeCall(BUILTIN.SUMIFS, 5),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushString(1), encodeCall(BUILTIN.AVERAGEIF, 2), encodeRet()],
      [
        encodePushRange(2),
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodePushString(2),
        encodeCall(BUILTIN.AVERAGEIFS, 5),
        encodeRet(),
      ],
      [encodePushRange(3), encodePushRange(4), encodeCall(BUILTIN.SUMPRODUCT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
        cellIndex(3, 7, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
        cellIndex(3, 7, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 1, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 2, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(3, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 3, width)]).toBe(70)
    expect(kernel.readTags()[cellIndex(3, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 4, width)]).toBe(70)
    expect(kernel.readTags()[cellIndex(3, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 5, width)]).toBe(4)
    expect(kernel.readTags()[cellIndex(3, 6, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 6, width)]).toBeCloseTo((10 + 20 + 40) / 3)
    expect(kernel.readTags()[cellIndex(3, 7, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 7, width)]).toBe(32)
  })

  it('evaluates INDEX, VLOOKUP, and HLOOKUP on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(32, 6, 3, 3, 12)
    kernel.uploadStrings(Uint32Array.from([0, 4, 9, 11, 13]), Uint32Array.from([4, 5, 2, 2, 2]), asciiCodes('pearappleQ1Q2Q3'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.Number,
        0,
        0,
        ValueTag.String,
        ValueTag.Number,
        0,
        0,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([0, 10, 0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 100, 200, 300, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 1, 0, 0, 0, 2, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 4, 5, 0, 1, 4, 5, 8, 9, 10, 12, 13, 14]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 6]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([2, 2, 2]), Uint32Array.from([2, 2, 3]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.INDEX, 3), encodeRet()],
      [encodePushString(1), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.VLOOKUP, 4), encodeRet()],
      [encodePushString(4), encodePushRange(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.HLOOKUP, 4), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]),
    )
    kernel.uploadConstants(new Float64Array([2, 2, 2, 0, 2, 0]), new Uint32Array([0, 2, 4]), new Uint32Array([2, 2, 2]))

    kernel.evalBatch(Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]))

    expect(kernel.readNumbers()[cellIndex(2, 0, width)]).toBe(20)
    expect(kernel.readNumbers()[cellIndex(2, 1, width)]).toBe(20)
    expect(kernel.readNumbers()[cellIndex(2, 2, width)]).toBe(300)
  })

  it('evaluates database aggregation builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    const height = 12
    const cellCount = width * height
    kernel.init(cellCount, 3, 0, 12, 3)
    kernel.uploadStrings(Uint32Array.from([0, 3, 9]), Uint32Array.from([3, 6, 5]), asciiCodes('AgeHeightYield'))

    const tags = new Uint8Array(cellCount)
    const numbers = new Float64Array(cellCount)
    const stringIds = new Uint32Array(cellCount)
    const errors = new Uint16Array(cellCount)
    const setNumber = (row: number, col: number, value: number) => {
      const index = cellIndex(row, col, width)
      tags[index] = ValueTag.Number
      numbers[index] = value
    }
    const setString = (row: number, col: number, stringId: number) => {
      const index = cellIndex(row, col, width)
      tags[index] = ValueTag.String
      stringIds[index] = stringId
    }

    setString(0, 0, 0)
    setString(0, 1, 1)
    setString(0, 2, 2)
    setNumber(1, 0, 10)
    setNumber(1, 1, 100)
    setNumber(1, 2, 5)
    setNumber(2, 0, 12)
    setNumber(2, 1, 110)
    setNumber(2, 2, 7)
    setNumber(3, 0, 12)
    setNumber(3, 1, 120)
    setNumber(3, 2, 9)
    setNumber(4, 0, 15)
    setNumber(4, 1, 130)
    setNumber(4, 2, 11)
    setString(0, 4, 0)
    setNumber(1, 4, 12)
    setString(0, 5, 0)
    setNumber(1, 5, 15)
    kernel.writeCells(tags, numbers, stringIds, errors)

    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(4, 0, width),
        cellIndex(4, 1, width),
        cellIndex(4, 2, width),
        cellIndex(0, 4, width),
        cellIndex(1, 4, width),
        cellIndex(0, 5, width),
        cellIndex(1, 5, width),
      ]),
      Uint32Array.from([0, 15, 17]),
      Uint32Array.from([15, 2, 2]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([5, 2, 2]), Uint32Array.from([3, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DAVERAGE, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DCOUNT, 3), encodeRet()],
      [encodePushRange(0), encodePushString(1), encodePushRange(1), encodeCall(BUILTIN.DCOUNTA, 3), encodeRet()],
      [encodePushRange(0), encodePushString(1), encodePushRange(2), encodeCall(BUILTIN.DGET, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DMAX, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DMIN, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DPRODUCT, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DSTDEV, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DSTDEVP, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DSUM, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DVAR, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DVARP, 3), encodeRet()],
    ])
    const outputCells = Uint32Array.from(Array.from({ length: 12 }, (_, index) => cellIndex(index, 7, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(outputCells)

    const outputTags = kernel.readTags()
    outputCells.forEach((index) => expect(outputTags[index]).toBe(ValueTag.Number))
    expectNumberCell(kernel, outputCells[0], 8)
    expectNumberCell(kernel, outputCells[1], 2)
    expectNumberCell(kernel, outputCells[2], 2)
    expectNumberCell(kernel, outputCells[3], 130)
    expectNumberCell(kernel, outputCells[4], 9)
    expectNumberCell(kernel, outputCells[5], 7)
    expectNumberCell(kernel, outputCells[6], 63)
    expectNumberCell(kernel, outputCells[7], Math.SQRT2, 12)
    expectNumberCell(kernel, outputCells[8], 1, 12)
    expectNumberCell(kernel, outputCells[9], 16)
    expectNumberCell(kernel, outputCells[10], 2)
    expectNumberCell(kernel, outputCells[11], 1)
  })

  it('evaluates vector lookup builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(40, 8, 6, 1, 2)
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 5, 9, 14, 18, 26]),
      Uint32Array.from([0, 5, 4, 5, 4, 8, 8]),
      asciiCodes('applepearpearplumfallbacknotfound'),
    )
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([
        0, 0, 0, 0, 10, 20, 30, 40, 1, 3, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint32Array([
        1, 2, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint16Array(40),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]), Uint32Array.from([0, 4, 8]), Uint32Array.from([4, 4, 3]))
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 3]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushString(2), encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.MATCH, 3), encodeRet()],
      [encodePushNumber(1), encodePushRange(2), encodePushNumber(2), encodeCall(BUILTIN.MATCH, 3), encodeRet()],
      [encodePushString(2), encodePushRange(0), encodePushNumber(0), encodePushNumber(3), encodeCall(BUILTIN.XMATCH, 4), encodeRet()],
      [encodePushString(2), encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.XLOOKUP, 3), encodeRet()],
      [encodePushString(6), encodePushRange(0), encodePushRange(1), encodePushString(5), encodeCall(BUILTIN.XLOOKUP, 4), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(4, 1, width),
        cellIndex(4, 2, width),
        cellIndex(4, 3, width),
        cellIndex(4, 4, width),
        cellIndex(4, 5, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array([0, 4, 1, -1]), new Uint32Array([0, 0, 0, 0, 0]), new Uint32Array([1, 2, 2, 0, 0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(4, 1, width),
        cellIndex(4, 2, width),
        cellIndex(4, 3, width),
        cellIndex(4, 4, width),
        cellIndex(4, 5, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(4, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 1, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(4, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 2, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(4, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 3, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(4, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 4, width)]).toBe(20)
    expect(kernel.readTags()[cellIndex(4, 5, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(4, 5, width)]).toBe(5)
  })

  it('evaluates LOOKUP, AREAS, ARRAYTOTEXT, COLUMNS, and ROWS on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(32, 8, 2, 3, 8)
    kernel.uploadStrings(Uint32Array.from([0, 0]), Uint32Array.from([0, 1]), asciiCodes('z'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 0, 3, 4, 10, 20, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(32),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 2, 3, 4, 5, 6, 4, 5]), Uint32Array.from([0, 3, 6]), Uint32Array.from([3, 3, 2]))
    kernel.uploadRangeShapes(Uint32Array.from([3, 3, 1]), Uint32Array.from([1, 1, 2]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.LOOKUP, 3), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.AREAS, 1), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.COLUMNS, 1), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.ROWS, 1), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.ARRAYTOTEXT, 1), encodeRet()],
      [encodePushRange(2), encodePushNumber(1), encodeCall(BUILTIN.ARRAYTOTEXT, 2), encodeRet()],
      [encodePushString(1), encodePushString(1), encodeCall(BUILTIN.LOOKUP, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array([3.5, 1]), new Uint32Array([0, 0, 0, 0, 0, 0, 0]), new Uint32Array([2, 2, 2, 2, 2, 2, 2]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(20)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 6, width)]).toBe(1)
    expect(kernel.readOutputStrings()).toEqual(['10\t20', '{10, 20}'])
  })

  it('evaluates TRANSPOSE, HSTACK, VSTACK, MINIFS, and MAXIFS on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    const pooledStrings = ['', 'x', 'a', 'b', 'c', '>0']
    kernel.init(40, 8, 1, 8, 24)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.String,
        ValueTag.Boolean,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Boolean,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ...Array.from({ length: 16 }, () => ValueTag.Empty),
      ]),
      new Float64Array([
        1,
        0,
        1,
        4,
        10,
        20,
        0,
        0,
        30,
        0,
        40,
        50,
        10,
        0,
        30,
        5,
        2,
        4,
        -1,
        6,
        0,
        0,
        0,
        0,
        ...Array.from({ length: 16 }, () => 0),
      ]),
      new Uint32Array([0, 1, 0, 0, 0, 0, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 2, 3, 2, ...Array.from({ length: 16 }, () => 0)]),
      new Uint16Array(40),
    )
    kernel.uploadStringLengths(Uint32Array.from(pooledStrings.map((value) => value.length)))
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 1, 2, 3, 4]),
      Uint32Array.from(pooledStrings.map((value) => value.length)),
      asciiCodes(pooledStrings.join('')),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]),
      Uint32Array.from([0, 4, 6, 8, 12, 16, 20]),
      Uint32Array.from([4, 2, 2, 4, 4, 4, 4]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([2, 2, 1, 2, 4, 4, 4]), Uint32Array.from([2, 1, 2, 2, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.TRANSPOSE, 1), encodeRet()],
      [encodePushRange(1), encodePushRange(2), encodePushString(4), encodeCall(BUILTIN.HSTACK, 3), encodeRet()],
      [encodePushRange(2), encodePushRange(3), encodePushString(4), encodeCall(BUILTIN.VSTACK, 3), encodeRet()],
      [
        encodePushRange(4),
        encodePushRange(5),
        encodePushString(5),
        encodePushRange(6),
        encodePushString(2),
        encodeCall(BUILTIN.MINIFS, 5),
        encodeRet(),
      ],
      [
        encodePushRange(4),
        encodePushRange(5),
        encodePushString(5),
        encodePushRange(6),
        encodePushString(2),
        encodeCall(BUILTIN.MAXIFS, 5),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0, 0, 0]), new Uint32Array([0, 0, 0, 0, 0]))

    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.Number)
    expect(readSpillValues(kernel, cellIndex(3, 0, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Number, value: 4 },
    ])

    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number)
    expect(readSpillValues(kernel, cellIndex(3, 1, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.String, value: 'a', stringId: 0 },
      { tag: ValueTag.String, value: 'b', stringId: 0 },
      { tag: ValueTag.String, value: 'c', stringId: 0 },
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.String, value: 'a', stringId: 0 },
      { tag: ValueTag.String, value: 'b', stringId: 0 },
      { tag: ValueTag.String, value: 'c', stringId: 0 },
    ])

    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.String)
    expect(readSpillValues(kernel, cellIndex(3, 2, width), pooledStrings)).toEqual([
      { tag: ValueTag.String, value: 'a', stringId: 0 },
      { tag: ValueTag.String, value: 'b', stringId: 0 },
      { tag: ValueTag.Number, value: 30 },
      { tag: ValueTag.Boolean, value: false },
      { tag: ValueTag.Number, value: 40 },
      { tag: ValueTag.Number, value: 50 },
      { tag: ValueTag.String, value: 'c', stringId: 0 },
      { tag: ValueTag.String, value: 'c', stringId: 0 },
    ])

    expect(kernel.readTags()[cellIndex(3, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 3, width)]).toBe(5)
    expect(kernel.readTags()[cellIndex(3, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 4, width)]).toBe(10)
  })

  it('evaluates exact-safe date builtins with Excel coercion and errors', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 10, 5, 2, 2)
    kernel.writeCells(
      new Uint8Array([3, 2, 4, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 1, 0, 45351, 45351.75, 60, 45322, 45337, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array([0, 0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    )

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.DATE, 3), encodeRet()],
      [encodePushCell(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.DATE, 3), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.YEAR, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.MONTH, 1), encodeRet()],
      [encodePushCell(5), encodeCall(BUILTIN.DAY, 1), encodeRet()],
      [encodePushCell(6), encodePushNumber(3), encodeCall(BUILTIN.EDATE, 2), encodeRet()],
      [encodePushCell(0), encodePushNumber(4), encodeCall(BUILTIN.EDATE, 2), encodeRet()],
      [encodePushCell(7), encodePushCell(1), encodeCall(BUILTIN.EOMONTH, 2), encodeRet()],
      [encodePushCell(2), encodePushNumber(4), encodeCall(BUILTIN.EOMONTH, 2), encodeRet()],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
        cellIndex(1, 9, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array([2024, 2, 29, 1.9, 1]), new Uint32Array([0]), new Uint32Array([5]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
        cellIndex(1, 9, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(45351)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 2, width)]).toBe(ErrorCode.Value)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(2024)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(29)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(45351)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 7, width)]).toBe(ErrorCode.Value)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(45382)
    expect(kernel.readTags()[cellIndex(1, 9, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 9, width)]).toBe(ErrorCode.Ref)
  })

  it('evaluates numeric-only dynamic-array builtins on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(24, 11, 1, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 3, 4, 5, 6, 4, 3, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]), Uint32Array.from([0, 6]), Uint32Array.from([6, 4]))
    kernel.uploadRangeShapes(Uint32Array.from([2, 4]), Uint32Array.from([3, 1]))

    const packed = packPrograms([
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(0),
        encodePushNumber(2),
        encodePushNumber(1),
        encodeCall(BUILTIN.OFFSET, 5),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.TAKE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(1), encodeCall(BUILTIN.DROP, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.CHOOSECOLS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.CHOOSEROWS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodePushNumber(4), encodeCall(BUILTIN.SORT, 3), encodeRet()],
      [encodePushRange(1), encodePushRange(1), encodeCall(BUILTIN.SORTBY, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.TOCOL, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.TOROW, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.WRAPROWS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.WRAPCOLS, 2), encodeRet()],
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]))
    kernel.uploadConstants(new Float64Array([0, 1, 2, 3, -1]), new Uint32Array([0, 0, 0, 0, 0]), new Uint32Array([1, 1, 1, 1, 1]))
    kernel.evalBatch(Uint32Array.from([12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]))

    expect(kernel.readTags()[12]).toBe(ValueTag.Number)
    expect(kernel.readTags()[13]).toBe(ValueTag.Number)
    expect(kernel.readTags()[14]).toBe(ValueTag.Number)
    expect(kernel.readTags()[15]).toBe(ValueTag.Number)
    expect(kernel.readTags()[16]).toBe(ValueTag.Number)
    expect(kernel.readTags()[17]).toBe(ValueTag.Number)
    expect(kernel.readTags()[18]).toBe(ValueTag.Number)
    expect(kernel.readTags()[19]).toBe(ValueTag.Number)
    expect(kernel.readTags()[20]).toBe(ValueTag.Number)
    expect(kernel.readTags()[21]).toBe(ValueTag.Number)
    expect(kernel.readTags()[22]).toBe(ValueTag.Number)

    expect(kernel.readSpillRows()[12]).toBe(2)
    expect(kernel.readSpillCols()[12]).toBe(1)
    expect(kernel.readSpillLengths()[12]).toBe(2)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[12], kernel.readSpillOffsets()[12] + 2))).toEqual([1, 4])

    expect(kernel.readSpillRows()[13]).toBe(2)
    expect(kernel.readSpillCols()[13]).toBe(3)
    expect(kernel.readSpillLengths()[13]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[13], kernel.readSpillOffsets()[13] + 6))).toEqual([
      1, 2, 3, 4, 5, 6,
    ])
    expect(kernel.readSpillRows()[14]).toBe(1)
    expect(kernel.readSpillCols()[14]).toBe(3)
    expect(kernel.readSpillLengths()[14]).toBe(3)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[14], kernel.readSpillOffsets()[14] + 3))).toEqual([4, 5, 6])
    expect(kernel.readSpillRows()[15]).toBe(2)
    expect(kernel.readSpillCols()[15]).toBe(1)
    expect(kernel.readSpillLengths()[15]).toBe(2)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[15], kernel.readSpillOffsets()[15] + 2))).toEqual([2, 5])
    expect(kernel.readSpillRows()[16]).toBe(1)
    expect(kernel.readSpillCols()[16]).toBe(3)
    expect(kernel.readSpillLengths()[16]).toBe(3)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[16], kernel.readSpillOffsets()[16] + 3))).toEqual([4, 5, 6])
    expect(kernel.readSpillRows()[17]).toBe(2)
    expect(kernel.readSpillCols()[17]).toBe(3)
    expect(kernel.readSpillLengths()[17]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[17], kernel.readSpillOffsets()[17] + 6))).toEqual([
      1, 1, 1, 4, 4, 4,
    ])
    expect(kernel.readSpillRows()[18]).toBe(4)
    expect(kernel.readSpillCols()[18]).toBe(1)
    expect(kernel.readSpillLengths()[18]).toBe(4)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[18], kernel.readSpillOffsets()[18] + 4))).toEqual([
      4, 4, 4, 4,
    ])
    expect(kernel.readSpillRows()[19]).toBe(6)
    expect(kernel.readSpillCols()[19]).toBe(1)
    expect(kernel.readSpillLengths()[19]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[19], kernel.readSpillOffsets()[19] + 6))).toEqual([
      1, 4, 2, 5, 3, 6,
    ])
    expect(kernel.readSpillRows()[20]).toBe(1)
    expect(kernel.readSpillCols()[20]).toBe(6)
    expect(kernel.readSpillLengths()[20]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[20], kernel.readSpillOffsets()[20] + 6))).toEqual([
      1, 2, 3, 4, 5, 6,
    ])
    expect(kernel.readSpillRows()[21]).toBe(3)
    expect(kernel.readSpillCols()[21]).toBe(2)
    expect(kernel.readSpillLengths()[21]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[21], kernel.readSpillOffsets()[21] + 6))).toEqual([
      1, 2, 3, 4, 5, 6,
    ])
    expect(kernel.readSpillRows()[22]).toBe(2)
    expect(kernel.readSpillCols()[22]).toBe(3)
    expect(kernel.readSpillLengths()[22]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[22], kernel.readSpillOffsets()[22] + 6))).toEqual([
      1, 2, 3, 4, 5, 6,
    ])
  })

  it('evaluates RAND from the uploaded recalc random sequence on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 1, 1, 1)
    kernel.writeCells(new Uint8Array(4), new Float64Array(4), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([
        encodeCall(BUILTIN.RAND, 0),
        encodeRet(),
        encodeCall(BUILTIN.RAND, 0),
        encodeCall(BUILTIN.RAND, 0),
        encodeBinary(Opcode.Add),
        encodeRet(),
      ]),
      new Uint32Array([0, 2]),
      new Uint32Array([2, 4]),
      new Uint32Array([0, 1]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]))
    kernel.uploadVolatileRandomValues(new Float64Array([0.625, 0.125, 0.875]))

    kernel.evalBatch(new Uint32Array([0, 1]))

    expect(kernel.readTags()[0]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[0]).toBe(0.625)
    expect(kernel.readTags()[1]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[1]).toBeCloseTo(1, 12)
  })

  it('evaluates EXPAND and TRIMRANGE on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(24, 4, 0, 4, 32)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
      ]),
      new Float64Array([0, 0, 0, 0, 10, 20, 0, 1, 2, 0, 30, 40, 0, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 4, width),
        cellIndex(0, 5, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(0, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(0, 3, width),
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
        cellIndex(2, 3, width),
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
      ]),
      Uint32Array.from([0, 4]),
      Uint32Array.from([4, 16]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([2, 4]), Uint32Array.from([2, 4]))

    const packedPrograms = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.EXPAND, 4), encodeRet()],
      [encodePushRange(1), encodeCall(BUILTIN.TRIMRANGE, 1), encodeRet()],
    ])
    const packedConstants = packConstants([[3, 3, 0], []])

    kernel.uploadPrograms(
      packedPrograms.programs,
      packedPrograms.offsets,
      packedPrograms.lengths,
      Uint32Array.from([cellIndex(3, 4, width), cellIndex(3, 5, width)]),
    )
    kernel.uploadConstants(packedConstants.constants, packedConstants.offsets, packedConstants.lengths)

    kernel.evalBatch(Uint32Array.from([cellIndex(3, 4, width), cellIndex(3, 5, width)]))

    expect(readSpillValues(kernel, cellIndex(3, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 30 },
      { tag: ValueTag.Number, value: 40 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 0 },
    ])
    expect(readSpillValues(kernel, cellIndex(3, 5, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Empty },
    ])
  })

  it('evaluates DATEDIF and financial scalar helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 25
    kernel.init(50, 1, 0, 25, 50)
    kernel.writeCells(new Uint8Array(50), new Float64Array(50), new Uint32Array(50), new Uint16Array(50))
    kernel.uploadStrings(Uint32Array.from([0, 0]), Uint32Array.from([0, 2]), asciiCodes('YM'))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.DATE, 3),
        encodePushString(1),
        encodeCall(BUILTIN.DATEDIF, 3),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.FVSCHEDULE, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.DB, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.DDB, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.VDB, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.SLN, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.SYD, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.DISC, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.INTRATE, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.RECEIVED, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.PRICEDISC, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.YIELDDISC, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.PRICEMAT, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.YIELDMAT, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodePushNumber(8),
        encodeCall(BUILTIN.ODDFPRICE, 9),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodePushNumber(8),
        encodeCall(BUILTIN.ODDFYIELD, 9),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodeCall(BUILTIN.ODDLPRICE, 8),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodeCall(BUILTIN.ODDLYIELD, 8),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TBILLPRICE, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TBILLYIELD, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TBILLEQ, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.EFFECT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.NOMINAL, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.PDURATION, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.RRI, 3), encodeRet()],
    ])
    const constants = packConstants([
      [2020, 1, 15, 2021, 3, 20],
      [1000, 0.09, 0.11, 0.1],
      [10000, 1000, 5, 1],
      [2400, 300, 10, 2],
      [2400, 300, 10, 1, 3],
      [10000, 1000, 9],
      [10000, 1000, 9, 1],
      [44927, 45017, 97, 100, 2],
      [44927, 45017, 1000, 1030, 2],
      [44927, 45017, 1000, 0.12, 2],
      [39494, 39508, 0.0525, 100, 2],
      [39494, 39508, 99.795, 100, 2],
      [39493, 39551, 39397, 0.061, 0.061, 0],
      [39522, 39755, 39394, 0.0625, 100.0123, 0],
      [39763, 44256, 39736, 39873, 0.0785, 0.0625, 100, 2, 1],
      [39763, 44256, 39736, 39873, 0.0575, 84.5, 100, 2, 0],
      [39485, 39614, 39370, 0.0375, 0.0405, 100, 2, 0],
      [39558, 39614, 39440, 0.0375, 99.875, 100, 2, 0],
      [39538, 39600, 0.09],
      [39538, 39600, 98.45],
      [39538, 39600, 0.0914],
      [0.12, 12],
      [0.12682503013196977, 12],
      [0.1, 100, 121],
      [2, 100, 121],
    ])

    const outputCells = Uint32Array.from([
      cellIndex(0, 0, width),
      cellIndex(0, 1, width),
      cellIndex(0, 2, width),
      cellIndex(0, 3, width),
      cellIndex(0, 4, width),
      cellIndex(0, 5, width),
      cellIndex(0, 6, width),
      cellIndex(0, 7, width),
      cellIndex(0, 8, width),
      cellIndex(0, 9, width),
      cellIndex(0, 10, width),
      cellIndex(0, 11, width),
      cellIndex(0, 12, width),
      cellIndex(0, 13, width),
      cellIndex(0, 14, width),
      cellIndex(0, 15, width),
      cellIndex(0, 16, width),
      cellIndex(0, 17, width),
      cellIndex(0, 18, width),
      cellIndex(0, 19, width),
      cellIndex(0, 20, width),
      cellIndex(0, 21, width),
      cellIndex(0, 22, width),
      cellIndex(0, 23, width),
      cellIndex(0, 24, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBe(2)
    expect(numbers[cellIndex(0, 1, width)]).toBeCloseTo(1330.89, 12)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(3690, 12)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(384, 12)
    expect(numbers[cellIndex(0, 4, width)]).toBeCloseTo(691.2, 12)
    expect(numbers[cellIndex(0, 5, width)]).toBe(1000)
    expect(numbers[cellIndex(0, 6, width)]).toBe(1800)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(0.12, 12)
    expect(numbers[cellIndex(0, 8, width)]).toBeCloseTo(0.12, 12)
    expect(numbers[cellIndex(0, 9, width)]).toBeCloseTo(1030.9278350515465, 12)
    expect(numbers[cellIndex(0, 10, width)]).toBeCloseTo(99.79583333333333, 12)
    expect(numbers[cellIndex(0, 11, width)]).toBeCloseTo(0.05282257198685834, 12)
    expect(numbers[cellIndex(0, 12, width)]).toBeCloseTo(99.98449887555694, 12)
    expect(numbers[cellIndex(0, 13, width)]).toBeCloseTo(0.060954333691538576, 12)
    expect(numbers[cellIndex(0, 14, width)]).toBeCloseTo(113.597717474079, 12)
    expect(numbers[cellIndex(0, 15, width)]).toBeCloseTo(0.0772455415972989, 11)
    expect(numbers[cellIndex(0, 16, width)]).toBeCloseTo(99.8782860147213, 12)
    expect(numbers[cellIndex(0, 17, width)]).toBeCloseTo(0.0451922356291692, 12)
    expect(numbers[cellIndex(0, 18, width)]).toBeCloseTo(98.45, 12)
    expect(numbers[cellIndex(0, 19, width)]).toBeCloseTo(0.09141696292534264, 12)
    expect(numbers[cellIndex(0, 20, width)]).toBeCloseTo(0.09415149356594302, 12)
    expect(numbers[cellIndex(0, 21, width)]).toBeCloseTo(0.12682503013196977, 12)
    expect(numbers[cellIndex(0, 22, width)]).toBeCloseTo(0.12, 12)
    expect(numbers[cellIndex(0, 23, width)]).toBeCloseTo(2, 12)
    expect(numbers[cellIndex(0, 24, width)]).toBeCloseTo(0.1, 12)
  })

  it('evaluates annuity and cumulative loan helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 11
    kernel.init(22, 1, 0, 11, 28)
    kernel.writeCells(new Uint8Array(20), new Float64Array(20), new Uint32Array(20), new Uint16Array(20))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.PV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.PMT, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NPER, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.RATE, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.IPMT, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.PPMT, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.ISPMT, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.CUMIPMT, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.CUMPRINC, 6),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.FV, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.NPV, 4), encodeRet()],
    ])
    const constants = packConstants([
      [0.1, 2, -576.1904761904761],
      [0.1, 2, 1000],
      [0.1, -576.1904761904761, 1000],
      [48, -200, 8000],
      [0.1, 1, 2, 1000],
      [0.1, 1, 2, 1000],
      [0.1, 1, 2, 1000],
      [0.09 / 12, 30 * 12, 125000, 13, 24, 0],
      [0.09 / 12, 30 * 12, 125000, 13, 24, 0],
      [0.1, 2, -100, -1000],
      [0.1, 100, 200, 300],
    ])

    const outputCells = Uint32Array.from([
      cellIndex(0, 0, width),
      cellIndex(0, 1, width),
      cellIndex(0, 2, width),
      cellIndex(0, 3, width),
      cellIndex(0, 4, width),
      cellIndex(0, 5, width),
      cellIndex(0, 6, width),
      cellIndex(0, 7, width),
      cellIndex(0, 8, width),
      cellIndex(0, 9, width),
      cellIndex(0, 10, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBeCloseTo(1000.0000000000006, 12)
    expect(numbers[cellIndex(0, 1, width)]).toBeCloseTo(-576.1904761904758, 12)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(1.9999999999999982, 12)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(0.007701472488246008, 12)
    expect(numbers[cellIndex(0, 4, width)]).toBeCloseTo(-100, 12)
    expect(numbers[cellIndex(0, 5, width)]).toBeCloseTo(-476.1904761904758, 12)
    expect(numbers[cellIndex(0, 6, width)]).toBeCloseTo(-50, 12)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(-11135.232130750845, 9)
    expect(numbers[cellIndex(0, 8, width)]).toBeCloseTo(-934.1071234208765, 9)
    expect(numbers[cellIndex(0, 9, width)]).toBeCloseTo(1420, 12)
    expect(numbers[cellIndex(0, 10, width)]).toBeCloseTo(481.5927873779113, 12)
  })

  it('evaluates coupon-date and periodic bond helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 1, 0, 10, 20)
    kernel.writeCells(new Uint8Array(20), new Float64Array(20), new Uint32Array(20), new Uint16Array(20))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPDAYBS, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPDAYS, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPDAYSNC, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPNCD, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPNUM, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPPCD, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BUILTIN.PRICE, 7),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BUILTIN.YIELD, 7),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.DURATION, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.MDURATION, 6),
        encodeRet(),
      ],
    ])
    const constants = packConstants([
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39493, 43054, 0.0575, 0.065, 100, 2, 0],
      [39493, 42689, 0.0575, 95.04287, 100, 2, 0],
      [43282, 54058, 0.08, 0.09, 2, 1],
      [39448, 42370, 0.08, 0.09, 2, 1],
    ])

    const outputCells = Uint32Array.from([
      cellIndex(0, 0, width),
      cellIndex(0, 1, width),
      cellIndex(0, 2, width),
      cellIndex(0, 3, width),
      cellIndex(0, 4, width),
      cellIndex(0, 5, width),
      cellIndex(0, 6, width),
      cellIndex(0, 7, width),
      cellIndex(0, 8, width),
      cellIndex(0, 9, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBe(70)
    expect(numbers[cellIndex(0, 1, width)]).toBe(180)
    expect(numbers[cellIndex(0, 2, width)]).toBe(110)
    expect(numbers[cellIndex(0, 3, width)]).toBe(39217)
    expect(numbers[cellIndex(0, 4, width)]).toBe(6)
    expect(numbers[cellIndex(0, 5, width)]).toBe(39036)
    expect(numbers[cellIndex(0, 6, width)]).toBeCloseTo(94.63436162132213, 12)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(0.065, 7)
    expect(numbers[cellIndex(0, 8, width)]).toBeCloseTo(10.919145281591925, 12)
    expect(numbers[cellIndex(0, 9, width)]).toBeCloseTo(5.735669813918838, 12)
  })

  it('evaluates covariance and regression helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(64, 10, 1, 2, 6)
    const cellTags = new Uint8Array(64)
    const cellNumbers = new Float64Array(64)
    cellTags[cellIndex(0, 0, width)] = ValueTag.Number
    cellNumbers[cellIndex(0, 0, width)] = 5
    cellTags[cellIndex(0, 1, width)] = ValueTag.Number
    cellNumbers[cellIndex(0, 1, width)] = 1
    cellTags[cellIndex(1, 0, width)] = ValueTag.Number
    cellNumbers[cellIndex(1, 0, width)] = 8
    cellTags[cellIndex(1, 1, width)] = ValueTag.Number
    cellNumbers[cellIndex(1, 1, width)] = 2
    cellTags[cellIndex(2, 0, width)] = ValueTag.Number
    cellNumbers[cellIndex(2, 0, width)] = 11
    cellTags[cellIndex(2, 1, width)] = ValueTag.Number
    cellNumbers[cellIndex(2, 1, width)] = 3
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(64), new Uint16Array(64))
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
      ]),
      Uint32Array.from([0, 3]),
      Uint32Array.from([3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3]), Uint32Array.from([1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.CORREL, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.COVAR, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.COVARIANCE_P, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.COVARIANCE_S, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.PEARSON, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.INTERCEPT, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.SLOPE, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.RSQ, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.STEYX, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.FORECAST, 3), encodeRet()],
    ])
    const constants = packConstants([[], [], [], [], [], [], [], [], [], [4]])
    const outputCells = Uint32Array.from([
      cellIndex(3, 0, width),
      cellIndex(3, 1, width),
      cellIndex(3, 2, width),
      cellIndex(3, 3, width),
      cellIndex(3, 4, width),
      cellIndex(3, 5, width),
      cellIndex(3, 6, width),
      cellIndex(3, 7, width),
      cellIndex(3, 8, width),
      cellIndex(3, 9, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const resultTags = kernel.readTags()
    const resultNumbers = kernel.readNumbers()
    outputCells.forEach((cell) => expect(resultTags[cell]).toBe(ValueTag.Number))
    expect(resultNumbers[cellIndex(3, 0, width)]).toBe(1)
    expect(resultNumbers[cellIndex(3, 1, width)]).toBe(2)
    expect(resultNumbers[cellIndex(3, 2, width)]).toBe(2)
    expect(resultNumbers[cellIndex(3, 3, width)]).toBe(3)
    expect(resultNumbers[cellIndex(3, 4, width)]).toBe(1)
    expect(resultNumbers[cellIndex(3, 5, width)]).toBe(2)
    expect(resultNumbers[cellIndex(3, 6, width)]).toBe(3)
    expect(resultNumbers[cellIndex(3, 7, width)]).toBe(1)
    expect(resultNumbers[cellIndex(3, 8, width)]).toBe(0)
    expect(resultNumbers[cellIndex(3, 9, width)]).toBe(14)
  })

  it('evaluates TREND and GROWTH spill helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 4, 0, 1, 2)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        ValueTag.Number,
        0,
        0,
      ]),
      new Float64Array([5, 1, 4, 2, 0, 0, 8, 2, 5, 4, 0, 0, 11, 3, 0, 8, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 2, width),
        cellIndex(0, 3, width),
        cellIndex(1, 3, width),
        cellIndex(2, 3, width),
      ]),
      Uint32Array.from([0, 3, 6, 8]),
      Uint32Array.from([3, 3, 2, 3]),
      Uint32Array.from([3, 3, 2, 3]),
      Uint32Array.from([1, 1, 1, 1]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3, 2, 3]), Uint32Array.from([1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodePushRange(2), encodeCall(BUILTIN.TREND, 3), encodeRet()],
      [encodePushRange(3), encodePushRange(1), encodePushRange(2), encodeCall(BUILTIN.GROWTH, 3), encodeRet()],
    ])
    const constants = packConstants([[], []])
    const outputCells = Uint32Array.from([cellIndex(0, 4, width), cellIndex(0, 5, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readSpillRows()[cellIndex(0, 4, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(0, 4, width)]).toBe(1)
    expect(
      readSpillValues(kernel, cellIndex(0, 4, width), []).map((value) => (value.tag === ValueTag.Number ? value.value : value)),
    ).toEqual([14, 17])

    expect(kernel.readSpillRows()[cellIndex(0, 5, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(0, 5, width)]).toBe(1)
    const growthValues = readSpillValues(kernel, cellIndex(0, 5, width), []).map((value) =>
      value.tag === ValueTag.Number ? value.value : value,
    )
    expect(growthValues[0]).toBeCloseTo(16, 12)
    expect(growthValues[1]).toBeCloseTo(32, 12)
  })

  it('evaluates LINEST and LOGEST spill helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 4, 0, 1, 3)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        0,
      ]),
      new Float64Array([5, 1, 0, 0, 2, 0, 8, 2, 0, 0, 4, 0, 11, 3, 0, 0, 8, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
        cellIndex(0, 4, width),
        cellIndex(1, 4, width),
        cellIndex(2, 4, width),
      ]),
      Uint32Array.from([0, 3, 6]),
      Uint32Array.from([3, 3, 3]),
      Uint32Array.from([3, 3, 3]),
      Uint32Array.from([1, 1, 1]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3, 3]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.LINEST, 2), encodeRet()],
      [encodePushRange(2), encodePushRange(1), encodeCall(BUILTIN.LOGEST, 2), encodeRet()],
    ])
    const constants = packConstants([[], []])
    const outputCells = Uint32Array.from([cellIndex(0, 2, width), cellIndex(0, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readSpillRows()[cellIndex(0, 2, width)]).toBe(1)
    expect(kernel.readSpillCols()[cellIndex(0, 2, width)]).toBe(2)
    expect(
      readSpillValues(kernel, cellIndex(0, 2, width), []).map((value) => (value.tag === ValueTag.Number ? value.value : value)),
    ).toEqual([3, 2])

    expect(kernel.readSpillRows()[cellIndex(0, 3, width)]).toBe(1)
    expect(kernel.readSpillCols()[cellIndex(0, 3, width)]).toBe(2)
    const logestValues = readSpillValues(kernel, cellIndex(0, 3, width), []).map((value) =>
      value.tag === ValueTag.Number ? value.value : value,
    )
    expect(logestValues[0]).toBeCloseTo(2, 12)
    expect(logestValues[1]).toBeCloseTo(1, 12)
  })

  it('evaluates rank helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 3, 0, 1, 4)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([10, 20, 20, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(20),
      new Uint16Array(20),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3]), Uint32Array.from([0]), Uint32Array.from([4]))
    kernel.uploadRangeShapes(Uint32Array.from([1]), Uint32Array.from([4]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.RANK, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.RANK_EQ, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.RANK_AVG, 2), encodeRet()],
    ])
    const constants = packConstants([[20], [20], [20]])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    outputCells.forEach((cell) => expect(tags[cell]).toBe(ValueTag.Number))
    expect(numbers[cellIndex(1, 0, width)]).toBe(2)
    expect(numbers[cellIndex(1, 1, width)]).toBe(2)
    expect(numbers[cellIndex(1, 2, width)]).toBe(2.5)
  })

  it('evaluates order-statistics helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(24, 12, 0, 1, 8)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 4, 7, 8, 9, 10, 12, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7]), Uint32Array.from([0]), Uint32Array.from([8]))
    kernel.uploadRangeShapes(Uint32Array.from([8]), Uint32Array.from([1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.MEDIAN, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.SMALL, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.LARGE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTILE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTILE_INC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTILE_EXC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTRANK, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTRANK_INC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTRANK_EXC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.QUARTILE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.QUARTILE_INC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.QUARTILE_EXC, 2), encodeRet()],
    ])
    const constants = packConstants([[], [3], [2], [0.25], [0.25], [0.25], [8], [8], [8], [1], [1], [1]])
    const outputCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
      cellIndex(1, 7, width),
      cellIndex(1, 8, width),
      cellIndex(1, 9, width),
      cellIndex(1, 10, width),
      cellIndex(1, 11, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    outputCells.forEach((cell) => expect(tags[cell]).toBe(ValueTag.Number))
    expect(numbers[cellIndex(1, 0, width)]).toBe(7.5)
    expect(numbers[cellIndex(1, 1, width)]).toBe(4)
    expect(numbers[cellIndex(1, 2, width)]).toBe(10)
    expect(numbers[cellIndex(1, 3, width)]).toBe(3.5)
    expect(numbers[cellIndex(1, 4, width)]).toBe(3.5)
    expect(numbers[cellIndex(1, 5, width)]).toBe(2.5)
    expect(numbers[cellIndex(1, 6, width)]).toBe(0.571)
    expect(numbers[cellIndex(1, 7, width)]).toBe(0.571)
    expect(numbers[cellIndex(1, 8, width)]).toBe(0.555)
    expect(numbers[cellIndex(1, 9, width)]).toBe(3.5)
    expect(numbers[cellIndex(1, 10, width)]).toBe(3.5)
    expect(numbers[cellIndex(1, 11, width)]).toBe(2.5)
  })

  it('evaluates MODE.MULT and FREQUENCY spill helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(30, 2, 0, 2, 8)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 2, 3, 3, 4, 79, 85, 78, 85, 50, 81, 60, 80, 90, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(30),
      new Uint16Array(30),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]),
      Uint32Array.from([0, 6, 12]),
      Uint32Array.from([6, 6, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([6, 6, 3]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.MODE_MULT, 1), encodeRet()],
      [encodePushRange(1), encodePushRange(2), encodeCall(BUILTIN.FREQUENCY, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 2, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 2, width)]))

    expect(kernel.readSpillRows()[cellIndex(1, 0, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(1, 0, width)]).toBe(1)
    expect(readSpillValues(kernel, cellIndex(1, 0, width), [])).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
    ])
    expect(kernel.readSpillRows()[cellIndex(1, 2, width)]).toBe(4)
    expect(kernel.readSpillCols()[cellIndex(1, 2, width)]).toBe(1)
    expect(readSpillValues(kernel, cellIndex(1, 2, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 0 },
    ])
  })

  it('evaluates MODE, CONFIDENCE.NORM, IFS, SWITCH, and XOR on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 6, 5, 1, 6)
    const pooledStrings = ['', 'big', 'small', 'one', 'other'] as const
    kernel.uploadStringLengths(Uint32Array.from(pooledStrings.map((value) => value.length)))
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 3, 8, 11]),
      Uint32Array.from(pooledStrings.map((value) => value.length)),
      asciiCodes(pooledStrings.join('')),
    )
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 2, 3, 3, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([6]), Uint32Array.from([1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.MODE, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.MODE_SNGL, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.CONFIDENCE_NORM, 3), encodeRet()],
      [
        encodePushBoolean(false),
        encodePushString(1),
        encodePushBoolean(true),
        encodePushString(2),
        encodeCall(BUILTIN.IFS, 4),
        encodeRet(),
      ],
      [encodePushNumber(3), encodePushNumber(4), encodePushString(3), encodePushString(4), encodeCall(BUILTIN.SWITCH, 4), encodeRet()],
      [encodePushBoolean(true), encodePushBoolean(false), encodePushBoolean(true), encodeCall(BUILTIN.XOR, 3), encodeRet()],
    ])
    const constants = packConstants([[], [], [0.05, 1, 100], [], [1, 1], []])
    const outputCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(0.1959963986120195, 12)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(pooledStrings[kernel.readStringIds()[cellIndex(1, 3, width)]] ?? '').toBe('small')
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(pooledStrings[kernel.readStringIds()[cellIndex(1, 4, width)]] ?? '').toBe('one')
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(0)
  })

  it('evaluates PROB and TRIMMEAN on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(32, 2, 0, 3, 16)
    const cellTags = new Uint8Array(32)
    const cellNumbers = new Float64Array(32)
    ;[1, 2, 3, 4].forEach((value, index) => {
      cellTags[cellIndex(index, 0, width)] = ValueTag.Number
      cellNumbers[cellIndex(index, 0, width)] = value
    })
    ;[0.1, 0.2, 0.3, 0.4].forEach((value, index) => {
      cellTags[cellIndex(index, 1, width)] = ValueTag.Number
      cellNumbers[cellIndex(index, 1, width)] = value
    })
    ;[1, 2, 4, 7, 8, 9, 10, 12].forEach((value, index) => {
      cellTags[cellIndex(index, 2, width)] = ValueTag.Number
      cellNumbers[cellIndex(index, 2, width)] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(32), new Uint16Array(32))
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(3, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
        cellIndex(3, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 2, width),
        cellIndex(2, 2, width),
        cellIndex(3, 2, width),
        cellIndex(4, 2, width),
        cellIndex(5, 2, width),
        cellIndex(6, 2, width),
        cellIndex(7, 2, width),
      ]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 8]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 8]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.PROB, 4), encodeRet()],
      [encodePushRange(2), encodePushNumber(0), encodeCall(BUILTIN.TRIMMEAN, 2), encodeRet()],
    ])
    const constants = packConstants([[2, 3], [0.25]])
    const outputCells = Uint32Array.from([cellIndex(0, 3, width), cellIndex(1, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 3, width)]).toBe(ValueTag.Number)
    expect(tags[cellIndex(1, 3, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(0.5, 12)
    expect(numbers[cellIndex(1, 3, width)]).toBeCloseTo(40 / 6, 12)
  })

  it('evaluates cash-flow rate helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(40, 4, 3, 4, 22)

    const tags = new Uint8Array(40)
    tags.fill(0)
    for (let index = 0; index < 22; index += 1) {
      tags[index] = ValueTag.Number
    }
    const numbers = new Float64Array(40)
    numbers.set(
      [
        -70000, 12000, 15000, 18000, 21000, 26000, -120000, 39000, 30000, 21000, 37000, 46000, -10000, 2750, 4250, 3250, 2750, 39448, 39508,
        39751, 39859, 39904,
      ],
      0,
    )
    kernel.writeCells(tags, numbers, new Uint32Array(40), new Uint16Array(40))
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]),
      Uint32Array.from([0, 6, 12, 17]),
      Uint32Array.from([6, 6, 5, 5]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([6, 6, 5, 5]), Uint32Array.from([1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.IRR, 1), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MIRR, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(2), encodePushRange(3), encodeCall(BUILTIN.XNPV, 3), encodeRet()],
      [encodePushRange(2), encodePushRange(3), encodeCall(BUILTIN.XIRR, 2), encodeRet()],
    ])
    const constants = packConstants([[], [0.1, 0.12], [0.09], []])
    const outputCells = Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width), cellIndex(3, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const resultTags = kernel.readTags()
    const resultNumbers = kernel.readNumbers()
    outputCells.forEach((cell) => expect(resultTags[cell]).toBe(ValueTag.Number))
    expect(resultNumbers[cellIndex(3, 0, width)]).toBeCloseTo(0.08663094803653162, 12)
    expect(resultNumbers[cellIndex(3, 1, width)]).toBeCloseTo(0.1260941303659051, 12)
    expect(resultNumbers[cellIndex(3, 2, width)]).toBeCloseTo(2086.647602031535, 9)
    expect(resultNumbers[cellIndex(3, 3, width)]).toBeCloseTo(0.37336253351883136, 12)
  })

  it('evaluates DAYS360 and YEARFRAC on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 1, 0, 3, 8)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DAYS360, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BUILTIN.DAYS360, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.YEARFRAC, 3), encodeRet()],
    ])
    const constants = packConstants([
      [45320, 45382],
      [45320, 45382],
      [45292, 45474, 3],
    ])
    const outputCells = Uint32Array.from([cellIndex(0, 0, width), cellIndex(0, 1, width), cellIndex(0, 2, width)])

    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBe(62)
    expect(numbers[cellIndex(0, 1, width)]).toBe(61)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(182 / 365, 12)
  })

  it('evaluates COUNTBLANK, ISOWEEKNUM, and TIMEVALUE on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(12, 1, 3, 1, 4)
    kernel.uploadStrings(Uint32Array.from([0, 0]), Uint32Array.from([0, 7]), asciiCodes('1:30 PM'))
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Empty, ValueTag.String, ValueTag.Empty, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(12),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3]), Uint32Array.from([0]), Uint32Array.from([4]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([2]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.COUNTBLANK, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodeCall(BUILTIN.ISOWEEKNUM, 1),
        encodeRet(),
      ],
      [encodePushString(1), encodeCall(BUILTIN.TIMEVALUE, 1), encodeRet()],
    ])
    const constants = packConstants([[], [2024, 1, 1], []])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)])

    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(0.5625, 12)
  })

  it('returns numeric spill descriptors for SEQUENCE on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 4, 1, 1)
    kernel.writeCells(new Uint8Array(4), new Float64Array(4), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Sequence, 4),
        encodeRet(),
      ]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([0]),
    )
    kernel.uploadConstants(new Float64Array([3, 1, 1, 1]), new Uint32Array([0]), new Uint32Array([4]))

    kernel.evalBatch(new Uint32Array([0]))

    expect(kernel.readTags()[0]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[0]).toBe(1)
    expect(kernel.readSpillRows()[0]).toBe(3)
    expect(kernel.readSpillCols()[0]).toBe(1)
    expect(kernel.readSpillOffsets()[0]).toBe(0)
    expect(kernel.readSpillLengths()[0]).toBe(3)
    expect(Array.from(kernel.readSpillTags().slice(0, kernel.getSpillValueCount()))).toEqual([
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
    ])
    expect(Array.from(kernel.readSpillNumbers().slice(0, kernel.getSpillValueCount()))).toEqual([1, 2, 3])
  })

  it('evaluates chi-square inverse functions and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(width, 1, 8, 8, 1)
    kernel.writeCells(new Uint8Array(width), new Float64Array(width), new Uint32Array(width), new Uint16Array(width))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHIDIST, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.LEGACY_CHIDIST, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHISQDIST, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHIINV, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHISQ_INV_RT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHISQINV, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.LEGACY_CHIINV, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHISQ_INV, 2), encodeRet()],
    ])
    const constants = packConstants([
      [18.307, 10],
      [18.307, 10],
      [18.307, 10],
      [0.050001, 10],
      [0.050001, 10],
      [0.050001, 10],
      [0.050001, 10],
      [0.93, 1],
    ])
    const outputCells = Uint32Array.from([
      cellIndex(0, 0, width),
      cellIndex(0, 1, width),
      cellIndex(0, 2, width),
      cellIndex(0, 3, width),
      cellIndex(0, 4, width),
      cellIndex(0, 5, width),
      cellIndex(0, 6, width),
      cellIndex(0, 7, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBeCloseTo(0.0500006, 6)
    expect(numbers[cellIndex(0, 1, width)]).toBeCloseTo(0.0500006, 6)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(0.0500006, 6)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(18.306973, 6)
    expect(numbers[cellIndex(0, 4, width)]).toBeCloseTo(18.306973, 6)
    expect(numbers[cellIndex(0, 5, width)]).toBeCloseTo(18.306973, 6)
    expect(numbers[cellIndex(0, 6, width)]).toBeCloseTo(18.306973, 6)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(3.2830202867594993, 12)
  })

  it('evaluates chi-square test functions and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 2, 2, 12)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([58, 35, 11, 25, 10, 23, 45.35, 47.65, 17.56, 18.44, 16.09, 16.91, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]), Uint32Array.from([0, 6]), Uint32Array.from([6, 6]))
    kernel.uploadRangeShapes(Uint32Array.from([3, 3]), Uint32Array.from([2, 2]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.CHISQ_TEST, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.CHITEST, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.LEGACY_CHITEST, 2), encodeRet()],
    ])
    const outputCells = Uint32Array.from([cellIndex(1, 4, width), cellIndex(1, 5, width), cellIndex(1, 6, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0]), new Uint32Array([0, 0, 0]))

    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(1, 4, width), 0.0003082, 7)
    expectNumberCell(kernel, cellIndex(1, 5, width), 0.0003082, 7)
    expectNumberCell(kernel, cellIndex(1, 6, width), 0.0003082, 7)
  })

  it('evaluates f-test and z-test functions and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 4, 3, 15)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([6, 7, 9, 15, 21, 20, 28, 31, 38, 40, 1, 2, 3, 4, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]),
      Uint32Array.from([0, 5, 10]),
      Uint32Array.from([5, 5, 5]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([5, 5, 5]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.F_TEST, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.FTEST, 2), encodeRet()],
      [encodePushRange(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.Z_TEST, 3), encodeRet()],
      [encodePushRange(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.ZTEST, 3), encodeRet()],
    ])
    const outputCells = Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width), cellIndex(2, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(new Float64Array([2, 1, 2, 1]), new Uint32Array([0, 0, 0, 2]), new Uint32Array([0, 0, 2, 2]))

    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(2, 0, width), 0.648317846786175, 12)
    expectNumberCell(kernel, cellIndex(2, 1, width), 0.648317846786175, 12)
    expectNumberCell(kernel, cellIndex(2, 2, width), 0.012673617875446075, 12)
    expectNumberCell(kernel, cellIndex(2, 3, width), 0.012673617875446075, 12)
  })

  it('evaluates beta and f distribution functions and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(width, 1, 12, 12, 1)
    kernel.writeCells(new Uint8Array(width), new Float64Array(width), new Uint32Array(width), new Uint16Array(width))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BETA_DIST, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BETADIST, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BETA_INV, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BETAINV, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushBoolean(true), encodeCall(BUILTIN.F_DIST, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.F_DIST_RT, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.FDIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.LEGACY_FDIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.F_INV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.F_INV_RT, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.FINV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.LEGACY_FINV, 3), encodeRet()],
    ])
    const constants = packConstants([
      [2, 8, 10, 1, 3],
      [2, 8, 10, 1, 3],
      [0.6854705810117458, 8, 10, 1, 3],
      [0.6854705810117458, 8, 10, 1, 3],
      [15.2068649, 6, 4],
      [15.2068649, 6, 4],
      [15.2068649, 6, 4],
      [15.2068649, 6, 4],
      [0.01, 6, 4],
      [0.01, 6, 4],
      [0.01, 6, 4],
      [0.01, 6, 4],
    ])
    const outputCells = Uint32Array.from(Array.from({ length: width }, (_, index) => cellIndex(0, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    for (let index = 0; index < width; index += 1) {
      expect(tags[cellIndex(0, index, width)]).toBe(ValueTag.Number)
    }
    expect(numbers[cellIndex(0, 0, width)]).toBeCloseTo(0.6854705810117458, 9)
    expect(numbers[cellIndex(0, 1, width)]).toBeCloseTo(0.6854705810117458, 9)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(2, 9)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(2, 9)
    expect(numbers[cellIndex(0, 4, width)]).toBeCloseTo(0.99, 9)
    expect(numbers[cellIndex(0, 5, width)]).toBeCloseTo(0.01, 9)
    expect(numbers[cellIndex(0, 6, width)]).toBeCloseTo(0.01, 9)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(0.01, 9)
    expect(numbers[cellIndex(0, 8, width)]).toBeCloseTo(0.10930991466299911, 8)
    expect(numbers[cellIndex(0, 9, width)]).toBeCloseTo(15.206864870947697, 7)
    expect(numbers[cellIndex(0, 10, width)]).toBeCloseTo(15.206864870947697, 7)
    expect(numbers[cellIndex(0, 11, width)]).toBeCloseTo(15.206864870947697, 7)
  })

  it('evaluates FILTER and UNIQUE spill builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    const pooledStrings = ['A', 'a', 'B', 'C']
    kernel.init(18, 6, 4, 3, 12)
    kernel.uploadStringLengths(Uint32Array.from(pooledStrings.map((value) => value.length)))
    kernel.uploadStrings(Uint32Array.from([0, 1, 2, 3]), Uint32Array.from([1, 1, 1, 1]), asciiCodes(pooledStrings.join('')))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 3, 2, 4, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 2, 3, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 4]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 4]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.FILTER, 2), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.UNIQUE, 1), encodeRet()],
      [encodePushRange(0), encodePushRange(0), encodePushNumber(0), encodeBinary(Opcode.Gt), encodeCall(BUILTIN.FILTER, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]),
    )
    kernel.uploadConstants(new Float64Array([2]), new Uint32Array([0, 0, 0]), new Uint32Array([0, 0, 1]))

    kernel.evalBatch(Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]))

    expect(readSpillValues(kernel, cellIndex(2, 0, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
    ])
    expect(readSpillValues(kernel, cellIndex(2, 1, width), pooledStrings)).toEqual([
      { tag: ValueTag.String, value: 'A', stringId: 0 },
      { tag: ValueTag.String, value: 'B', stringId: 0 },
      { tag: ValueTag.String, value: 'C', stringId: 0 },
    ])
    expect(readSpillValues(kernel, cellIndex(2, 2, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
    ])
  })

  it('evaluates internal BYROW and BYCOL sum spill builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 2, 0, 1, 6)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 0, 0, 0, 0, 3, 4, 0, 0, 0, 0, 5, 6, 0, 0, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 6, 7, 12, 13]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([2]))
    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.BYROW_SUM, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.BYCOL_SUM, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array(0), new Uint32Array([0, 0]), new Uint32Array([0, 0]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]))

    expect(readSpillValues(kernel, cellIndex(0, 3, width), [])).toEqual([
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Number, value: 11 },
    ])
    expect(readSpillValues(kernel, cellIndex(0, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 12 },
    ])
  })

  it('evaluates internal REDUCE and SCAN sum builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 2, 1, 1, 3)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2]), Uint32Array.from([0]), Uint32Array.from([3]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([1]))
    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.REDUCE_SUM, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.SCAN_SUM, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([0]), new Uint32Array([0, 1]), new Uint32Array([1, 1]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]))

    expect(decodeValueTag(kernel.readTags()[cellIndex(0, 3, width)] ?? ValueTag.Empty)).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(0, 3, width)]).toBe(6)
    expect(readSpillValues(kernel, cellIndex(0, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 6 },
    ])
  })

  it('evaluates internal REDUCE and SCAN product builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 2, 1, 1, 3)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([2, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2]), Uint32Array.from([0]), Uint32Array.from([3]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([1]))
    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.REDUCE_PRODUCT, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.SCAN_PRODUCT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([1, 1]), new Uint32Array([0, 1]), new Uint32Array([1, 1]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]))

    expect(decodeValueTag(kernel.readTags()[cellIndex(0, 3, width)] ?? ValueTag.Empty)).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(0, 3, width)]).toBe(24)
    expect(readSpillValues(kernel, cellIndex(0, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Number, value: 24 },
    ])
  })

  it('evaluates internal MAKEARRAY sum spill builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 1, 2, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
      ]),
      new Float64Array(8),
      new Uint32Array(8),
      new Uint16Array(8),
    )
    kernel.uploadPrograms(
      Uint32Array.from([encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MAKEARRAY_SUM, 2), encodeRet()]),
      Uint32Array.from([0]),
      Uint32Array.from([4]),
      Uint32Array.from([cellIndex(0, 0, width)]),
    )
    kernel.uploadConstants(new Float64Array([2, 2]), new Uint32Array([0]), new Uint32Array([2]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 0, width)]))

    expect(readSpillValues(kernel, cellIndex(0, 0, width), [])).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
    ])
  })

  it('evaluates internal BYROW and BYCOL aggregate spill builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 2, 2, 1, 2)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 0, 0, 0, 0, 3, 4, 0, 0, 0, 0, 5, 6, 0, 0, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 6, 7, 12, 13]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([2]))
    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.BYROW_AGGREGATE, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.BYCOL_AGGREGATE, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([2, 6]), new Uint32Array([0, 1]), new Uint32Array([1, 1]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]))

    expect(readSpillValues(kernel, cellIndex(0, 3, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1.5 },
      { tag: ValueTag.Number, value: 3.5 },
      { tag: ValueTag.Number, value: 5.5 },
    ])
    expect(readSpillValues(kernel, cellIndex(0, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 3 },
    ])
  })

  it('evaluates numeric aggregate builtins over native SEQUENCE arrays on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(12, 4, 24, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
      ]),
      new Float64Array([3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(12),
      new Uint16Array(12),
    )
    kernel.uploadPrograms(
      new Uint32Array([
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Sum, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Avg, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Min, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Max, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Count, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.CountA, 1),
        encodeRet(),
      ]),
      new Uint32Array([0, 7, 14, 21, 28, 35]),
      new Uint32Array([7, 7, 7, 7, 7, 7]),
      new Uint32Array([1, 2, 3, 4, 5, 6]),
    )
    kernel.uploadConstants(
      new Float64Array([1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]),
      new Uint32Array([0, 3, 6, 9, 12, 15]),
      new Uint32Array([3, 3, 3, 3, 3, 3]),
    )

    kernel.evalBatch(new Uint32Array([1, 2, 3, 4, 5, 6]))

    expect(kernel.readTags()[1]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[1]).toBe(6)
    expect(kernel.readNumbers()[2]).toBe(2)
    expect(kernel.readNumbers()[3]).toBe(1)
    expect(kernel.readNumbers()[4]).toBe(3)
    expect(kernel.readNumbers()[5]).toBe(3)
    expect(kernel.readNumbers()[6]).toBe(3)
  })

  it('evaluates TODAY and NOW from the uploaded recalc timestamp on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 1, 1, 1)
    kernel.writeCells(new Uint8Array(4), new Float64Array(4), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([encodeCall(BUILTIN.TODAY, 0), encodeRet(), encodeCall(BUILTIN.NOW, 0), encodeRet()]),
      new Uint32Array([0, 2]),
      new Uint32Array([2, 2]),
      new Uint32Array([0, 1]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]))
    kernel.uploadVolatileNowSerial(46100.65659722222)

    kernel.evalBatch(new Uint32Array([0, 1]))

    expect(kernel.readTags()[0]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[0]).toBe(46100)
    expect(kernel.readTags()[1]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[1]).toBeCloseTo(46100.65659722222, 12)
  })

  it('evaluates TIME, HOUR, MINUTE, SECOND, and WEEKDAY on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(24, 8, 5, 1, 1)
    kernel.writeCells(
      new Uint8Array([1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0.5208333333333334, 0.5208449074074074, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TIME, 3), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.HOUR, 1), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.MINUTE, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.SECOND, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodeCall(BUILTIN.WEEKDAY, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodePushNumber(3),
        encodeCall(BUILTIN.WEEKDAY, 2),
        encodeRet(),
      ],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )
    kernel.uploadConstants(
      new Float64Array([12, 30, 0, 2026, 3, 15, 2]),
      new Uint32Array([0, 0, 0, 0, 3, 3]),
      new Uint32Array([3, 0, 0, 0, 3, 4]),
    )
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )

    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0.5208333333333334)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(12)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(30)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(7)
  })

  it('evaluates DAYS, WEEKNUM, WORKDAY, and NETWORKDAYS on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(30, 8, 1, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([46097, 46101, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(30),
      new Uint16Array(30),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DAYS, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.WEEKNUM, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WEEKNUM, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WORKDAY, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushCell(0), encodeCall(BUILTIN.WORKDAY, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.NETWORKDAYS, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushCell(0), encodeCall(BUILTIN.NETWORKDAYS, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    )
    kernel.uploadConstants(
      new Float64Array([46101, 46094, 46096, 46096, 2, 46094, 1, 46094, 1, 46094, 46101, 46094, 46101]),
      new Uint32Array([0, 2, 3, 5, 7, 9, 11]),
      new Uint32Array([2, 1, 2, 2, 2, 2, 2]),
    )
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    )

    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(7)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(12)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(11)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(46097)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(46098)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(6)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(5)
  })

  it('evaluates WORKDAY.INTL and NETWORKDAYS.INTL on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 10, 4, 4)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.WORKDAY_INTL, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BUILTIN.WORKDAY_INTL, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NETWORKDAYS_INTL, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BUILTIN.NETWORKDAYS_INTL, 4),
        encodeRet(),
      ],
    ])
    const constants = packConstants([
      [46094, 1, 7],
      [46094, 2, 7, 46096],
      [46094, 46098, 7],
      [46094, 46098, 7, 46096],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(46097)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(46099)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(2)
  })

  it('evaluates logical and rounding builtins with parity-safe scalar semantics', async () => {
    const kernel = await createKernel()
    kernel.init(8, 8, 4, 4, 4)
    kernel.writeCells(
      new Uint8Array([1, 1, 4, 0, 0, 0, 0, 0]),
      new Float64Array([123.4, 1, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(8),
      new Uint16Array([0, 0, ErrorCode.Value, 0, 0, 0, 0, 0]),
    )
    kernel.uploadPrograms(
      new Uint32Array([
        (3 << 24) | 0,
        (1 << 24) | 0,
        (20 << 24) | (8 << 8) | 2,
        255 << 24,

        (3 << 24) | 1,
        (2 << 24) | 1,
        (20 << 24) | (9 << 8) | 2,
        255 << 24,

        (3 << 24) | 1,
        (20 << 24) | (15 << 8) | 1,
        255 << 24,

        (3 << 24) | 2,
        (2 << 24) | 2,
        (20 << 24) | (13 << 8) | 2,
        255 << 24,
      ]),
      new Uint32Array([0, 4, 8, 11]),
      new Uint32Array([4, 4, 3, 4]),
      new Uint32Array([3, 4, 5, 6]),
    )
    kernel.uploadConstants(new Float64Array([-1, 0.5, 1]), new Uint32Array([0, 0, 0, 0]), new Uint32Array([2, 2, 0, 1]))

    kernel.evalBatch(new Uint32Array([3, 4, 5, 6]))

    expect(kernel.readNumbers()[3]).toBe(120)
    expect(kernel.readNumbers()[4]).toBe(1)
    expect(kernel.readTags()[5]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[5]).toBe(0)
    expect(kernel.readTags()[6]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[6]).toBe(ErrorCode.Value)
  })

  it('evaluates statistical special functions on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(24, 10, 11, 1, 1)
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BUILTIN.ERF, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.ERF, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.ErfPrecise, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.ERFC, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.ErfcPrecise, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.FISHER, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.FISHERINV, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.GAMMALN, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.GammalnPrecise, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.GAMMA, 1), encodeRet()],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 10 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[1], [0, 1], [1], [1], [1], [0.5], [0.5493061443340549], [5], [5], [5]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 10 }, (_, index) => cellIndex(1, index, width))))

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.8427006897475899, 7)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.8427006897475899, 7)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.8427006897475899, 7)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.15729931025241006, 7)
    expectNumberCell(kernel, cellIndex(1, 4, width), 0.15729931025241006, 7)
    expectNumberCell(kernel, cellIndex(1, 5, width), 0.5493061443340549, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), 0.5, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), Math.log(24), 12)
    expectNumberCell(kernel, cellIndex(1, 8, width), Math.log(24), 12)
    expectNumberCell(kernel, cellIndex(1, 9, width), 24, 10)
  })

  it('evaluates statistical scalar and dispersion builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 14
    kernel.init(32, 12, 16, 2, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 4]), Uint32Array.from([0, 4, 4]), asciiCodes('skip'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 3, 4, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1, 2, 3, 0, 1, 2, 3, 4]), new Uint32Array([0, 4]), new Uint32Array([4, 5]))
    kernel.uploadRangeShapes(new Uint32Array([4, 5]), new Uint32Array([1, 1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.STANDARDIZE, 3), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.STDEV, 1), encodeRet()],
      [encodePushNumber(0), encodePushBoolean(true), encodePushCell(5), encodeCall(BUILTIN.STDEVA, 3), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.VAR, 1), encodeRet()],
      [encodePushNumber(0), encodePushBoolean(true), encodePushCell(5), encodeCall(BUILTIN.VARA, 3), encodeRet()],
      [encodePushRange(1), encodeCall(BUILTIN.SKEW, 1), encodeRet()],
      [encodePushRange(1), encodeCall(BUILTIN.KURT, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BUILTIN.NORMDIST, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NORMINV, 3), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.NORMSDIST, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.NORMSINV, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.LOGINV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.LOGNORMDIST, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [1, 0, 1],
      [],
      [2],
      [],
      [2],
      [],
      [],
      [1, 0, 1],
      [0.8413447460685429, 0, 1],
      [1],
      [0.001],
      [0.5, 0, 1],
      [1, 0, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))))

    expectNumberCell(kernel, cellIndex(1, 0, width), 1)
    expectNumberCell(kernel, cellIndex(1, 1, width), Math.sqrt(5 / 3), 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1)
    expectNumberCell(kernel, cellIndex(1, 3, width), 5 / 3, 12)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1)
    expectNumberCell(kernel, cellIndex(1, 5, width), 0, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), -1.2, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), 0.8413447460685429, 7)
    expectNumberCell(kernel, cellIndex(1, 8, width), 1, 8)
    expectNumberCell(kernel, cellIndex(1, 9, width), 0.8413447460685429, 7)
    expectNumberCell(kernel, cellIndex(1, 10, width), -3.090232306167813, 8)
    expectNumberCell(kernel, cellIndex(1, 11, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 12, width), 0.5, 8)
  })

  it('evaluates statistical distribution builtins and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(48, 22, 64, 1, 1)
    kernel.writeCells(new Uint8Array(48), new Float64Array(48), new Uint32Array(48), new Uint16Array(48))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.CONFIDENCE, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(false), encodeCall(BUILTIN.EXPONDIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BuiltinId.ExponDist, 3), encodeRet()],
      [encodePushNumber(3), encodePushNumber(4), encodePushBoolean(false), encodeCall(BUILTIN.POISSON, 3), encodeRet()],
      [encodePushNumber(3), encodePushNumber(4), encodePushBoolean(true), encodeCall(BuiltinId.PoissonDist, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(false),
        encodeCall(BUILTIN.WEIBULL, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.WeibullDist, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(false),
        encodeCall(BUILTIN.GAMMADIST, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.GammaDist, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHIDIST, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.ChisqDistRt, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BUILTIN.CHISQ_DIST, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(false),
        encodeCall(BUILTIN.BINOMDIST, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.BinomDist, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BUILTIN.BINOM_DIST_RANGE, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.BINOM_DIST_RANGE, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.CRITBINOM, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.BinomInv, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.HYPGEOMDIST, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushBoolean(true),
        encodeCall(BuiltinId.HypgeomDist, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NEGBINOMDIST, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.NegbinomDist, 4),
        encodeRet(),
      ],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 22 }, (_, index) => cellIndex(1 + Math.floor(index / width), index % width, width))),
    )
    const constants = packConstants([
      [0.05, 1.5, 100],
      [1, 2],
      [1, 2],
      [3, 2.5],
      [1.5, 2, 3],
      [1.5, 2, 3],
      [2, 3, 2],
      [2, 3, 2],
      [3, 4],
      [3, 4],
      [3, 4],
      [2, 4, 0.5],
      [2, 4, 0.5],
      [6, 0.5, 2, 4],
      [6, 0.5, 2],
      [6, 0.5, 0.7],
      [6, 0.5, 0.7],
      [1, 4, 3, 10],
      [1, 4, 3, 10],
      [2, 3, 0.5],
      [2, 3, 0.5],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    const targetCells = Uint32Array.from(
      Array.from({ length: 22 }, (_, index) => cellIndex(1 + Math.floor(index / width), index % width, width)),
    )
    kernel.evalBatch(targetCells)

    expectNumberCell(kernel, targetCells[0], 0.2939945976810081, 9)
    expectNumberCell(kernel, targetCells[1], 0.2706705664732254, 12)
    expectNumberCell(kernel, targetCells[2], 0.8646647167633873, 12)

    expect(kernel.readTags()[targetCells[3]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[4]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[5]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[6]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[7]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[8]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[9]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[10]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[11]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[12]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[13]]).toBe(ValueTag.Error)
    expect(kernel.readTags()[targetCells[14]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[15]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[16]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[17]]).toBe(ValueTag.Error)
    expect(kernel.readTags()[targetCells[18]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[19]]).toBe(ValueTag.Error)
    expect(kernel.readTags()[targetCells[20]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[21]]).toBe(ValueTag.Error)
  })

  it('returns statistical value errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 5, 5, 1, 2)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Empty, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([1, 0.5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))
    kernel.uploadRangeShapes(new Uint32Array([2]), new Uint32Array([1]))

    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.FISHER, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.GAMMA, 1), encodeRet()],
      [
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BINOM_DIST_RANGE, 4),
        encodeRet(),
      ],
      [encodePushRange(0), encodeCall(BUILTIN.ERF, 1), encodeRet()],
      [encodePushError(ErrorCode.Ref), encodePushNumber(0), encodePushBoolean(false), encodeCall(BUILTIN.POISSON, 3), encodeRet()],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 5 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[0], [0], [4, 0.5, 3, 2], [], [1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 5 }, (_, index) => cellIndex(1, index, width))))

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 4, width), ErrorCode.Ref)
  })

  it('evaluates student-t scalar distribution builtins and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 9
    kernel.init(width, 2, width, 1, 1)
    kernel.writeCells(new Uint8Array(width * 2), new Float64Array(width * 2), new Uint32Array(width * 2), new Uint16Array(width * 2))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BUILTIN.T_DIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_DIST_RT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_DIST_2T, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TDIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_INV, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_INV_2T, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TINV, 2), encodeRet()],
    ])
    const constants = packConstants([
      [1, 1],
      [1, 1],
      [1, 1],
      [1, 1, 2],
      [0.75, 1],
      [0.5, 1],
      [0.5, 1],
    ])
    const outputCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 0.75, 12)
    expectNumberCell(kernel, outputCells[1], 0.25, 12)
    expectNumberCell(kernel, outputCells[2], 0.5, 12)
    expectNumberCell(kernel, outputCells[3], 0.5, 12)
    expectNumberCell(kernel, outputCells[4], 1, 9)
    expectNumberCell(kernel, outputCells[5], 1, 9)
    expectNumberCell(kernel, outputCells[6], 1, 9)
  })

  it('evaluates CONFIDENCE.T on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 2
    kernel.init(width, 2, width, 1, 1)
    kernel.writeCells(new Uint8Array(width * 2), new Float64Array(width * 2), new Uint32Array(width * 2), new Uint16Array(width * 2))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.CONFIDENCE_T, 3), encodeRet()],
    ])
    const constants = packConstants([[0.5, 2, 4]])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 0.764892328404345, 12)
  })

  it('evaluates GAMMA.INV and legacy alias on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 3
    kernel.init(width, 2, width, 1, 1)
    kernel.writeCells(new Uint8Array(width * 2), new Float64Array(width * 2), new Uint32Array(width * 2), new Uint16Array(width * 2))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.GAMMA_INV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.GAMMAINV, 3), encodeRet()],
    ])
    const constants = packConstants([
      [0.08030139707139418, 3, 2],
      [0.08030139707139418, 3, 2],
    ])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 2, 10)
    expectNumberCell(kernel, outputCells[1], 2, 10)
  })

  it('evaluates T.TEST and legacy alias on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(width, 4, 2, 2, 2)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 1, 0, 0, 0, 0, 2, 3, 0, 0, 0, 0, 4, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(width * 4),
      new Uint16Array(width * 4),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
      ]),
      Uint32Array.from([0, 3]),
      Uint32Array.from([3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3]), Uint32Array.from([1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_TEST, 4), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TTEST, 4), encodeRet()],
    ])
    const constants = packConstants([
      [2, 1],
      [2, 1],
    ])
    const outputCells = Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 1, 12)
    expectNumberCell(kernel, outputCells[1], 1, 12)
  })

  it('materializes pivots using the actual source width', async () => {
    const kernel = await createKernel()
    kernel.init(16, 1, 1, 1, 16)

    const strings = ['', 'Region', 'Notes', 'Product', 'Sales', 'East', 'Widget', 'West', 'Gizmo', 'priority']
    const offsets = new Uint32Array(strings.length)
    const lengths = new Uint32Array(strings.length)
    const data: number[] = []
    let offset = 0
    strings.forEach((text, index) => {
      offsets[index] = offset
      lengths[index] = text.length
      for (const char of text) {
        data.push(char.charCodeAt(0))
      }
      offset += text.length
    })
    kernel.uploadStrings(offsets, lengths, Uint16Array.from(data))

    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
      ]),
      new Float64Array([0, 0, 0, 0, 0, 0, 0, 10, 0, 0, 0, 7, 0, 0, 0, 5]),
      new Uint32Array([1, 2, 3, 4, 5, 9, 6, 0, 7, 9, 6, 0, 5, 9, 8, 0]),
      new Uint16Array(16),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from(Array.from({ length: 16 }, (_, index) => index)),
      new Uint32Array([0]),
      new Uint32Array([16]),
    )
    kernel.uploadRangeShapes(new Uint32Array([4]), new Uint32Array([4]))

    const materialized = kernel.materializePivotTable(0, 4, Uint32Array.from([0]), Uint32Array.from([3, 2]), Uint8Array.from([1, 2]))

    expect(materialized.rows).toBe(3)
    expect(materialized.cols).toBe(3)
    expect(materialized.tags[0]).toBe(ValueTag.String)
    expect(materialized.stringIds[0]).toBe(1)
    expect(materialized.tags[1]).toBe(ValueTag.String)
    expect(materialized.stringIds[1]).toBe(4)
    expect(materialized.tags[2]).toBe(ValueTag.String)
    expect(materialized.stringIds[2]).toBe(3)
    expect(materialized.tags[3]).toBe(ValueTag.String)
    expect(materialized.stringIds[3]).toBe(5)
    expect(materialized.tags[4]).toBe(ValueTag.Number)
    expect(materialized.numbers[4]).toBe(15)
    expect(materialized.tags[5]).toBe(ValueTag.Number)
    expect(materialized.numbers[5]).toBe(2)
    expect(materialized.tags[6]).toBe(ValueTag.String)
    expect(materialized.stringIds[6]).toBe(7)
    expect(materialized.tags[7]).toBe(ValueTag.Number)
    expect(materialized.numbers[7]).toBe(7)
    expect(materialized.tags[8]).toBe(ValueTag.Number)
    expect(materialized.numbers[8]).toBe(1)
  })
})
