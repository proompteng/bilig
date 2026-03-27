import { BUILTINS, type BuiltinId } from "@bilig/protocol";

export type BuiltinCapabilityCategory =
  | "aggregation"
  | "logical"
  | "information"
  | "text"
  | "date-time"
  | "lookup-reference"
  | "statistical"
  | "dynamic-array"
  | "lambda"
  | "math";

export type BuiltinJsStatus = "implemented" | "special-js-only";
export type BuiltinWasmStatus = "production" | "not-started";

export interface BuiltinCapability {
  readonly id?: BuiltinId;
  readonly name: string;
  readonly category: BuiltinCapabilityCategory;
  readonly jsStatus: BuiltinJsStatus;
  readonly wasmStatus: BuiltinWasmStatus;
  readonly needsMetadata: boolean;
  readonly needsArrayRuntime: boolean;
  readonly needsExternalAdapter: boolean;
}

const jsSpecialBuiltinNames = new Set([
  "LET",
  "LAMBDA",
  "MAKEARRAY",
  "MAP",
  "REDUCE",
  "SCAN",
  "BYROW",
  "BYCOL",
  "CELL",
  "COLUMN",
  "FORMULA",
  "FORMULATEXT",
  "INDIRECT",
  "ROW",
  "SHEET",
  "SHEETS",
  "TEXTSPLIT",
]);

const wasmProductionBuiltinNames = new Set([
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
  "SIN",
  "COS",
  "TAN",
  "ASIN",
  "ACOS",
  "ATAN",
  "ATAN2",
  "DEGREES",
  "RADIANS",
  "EXP",
  "LN",
  "LOG",
  "LOG10",
  "POWER",
  "SQRT",
  "PI",
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
  "DAYS",
  "WEEKNUM",
  "WORKDAY",
  "NETWORKDAYS",
  "REPLACE",
  "SUBSTITUTE",
  "REPT",
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
  "VALUE",
  "FILTER",
  "UNIQUE",
  "EXPAND",
  "OFFSET",
  "TAKE",
  "DROP",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "SORT",
  "SORTBY",
  "TOCOL",
  "TOROW",
  "WRAPROWS",
  "WRAPCOLS",
  "TRIMRANGE",
  "LOOKUP",
  "AREAS",
  "ARRAYTOTEXT",
  "COLUMNS",
  "ROWS",
  "TRANSPOSE",
  "HSTACK",
  "VSTACK",
  "MINIFS",
  "MAXIFS",
  "ERF",
  "ERF.PRECISE",
  "ERFC",
  "ERFC.PRECISE",
  "FISHER",
  "FISHERINV",
  "GAMMALN",
  "GAMMALN.PRECISE",
  "GAMMA",
  "CONFIDENCE",
  "EXPONDIST",
  "EXPON.DIST",
  "POISSON",
  "POISSON.DIST",
  "WEIBULL",
  "WEIBULL.DIST",
  "GAMMADIST",
  "GAMMA.DIST",
  "CHIDIST",
  "CHISQ.DIST.RT",
  "CHISQ.DIST",
  "BINOMDIST",
  "BINOM.DIST",
  "BINOM.DIST.RANGE",
  "CRITBINOM",
  "BINOM.INV",
  "HYPGEOMDIST",
  "HYPGEOM.DIST",
  "NEGBINOMDIST",
  "NEGBINOM.DIST",
]);

const aggregationBuiltinNames = new Set(["SUM", "AVG", "MIN", "MAX", "COUNT", "COUNTA"]);
const logicalBuiltinNames = new Set(["IF", "AND", "OR", "NOT", "IFERROR", "IFNA", "NA"]);
const informationBuiltinNames = new Set([
  "CELL",
  "FORMULA",
  "FORMULATEXT",
  "ISBLANK",
  "ISNUMBER",
  "ISTEXT",
  "N",
  "SHEET",
  "SHEETS",
  "T",
  "TYPE",
]);
const textBuiltinNames = new Set([
  "LEN",
  "CONCAT",
  "REPLACE",
  "SUBSTITUTE",
  "REPT",
  "EXACT",
  "LEFT",
  "RIGHT",
  "MID",
  "TRIM",
  "UPPER",
  "LOWER",
  "FIND",
  "SEARCH",
  "VALUE",
  "ARRAYTOTEXT",
]);
const dateTimeBuiltinNames = new Set([
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
  "DAYS",
  "WEEKNUM",
  "WORKDAY",
  "NETWORKDAYS",
]);
const lookupBuiltinNames = new Set([
  "COLUMN",
  "MATCH",
  "XMATCH",
  "XLOOKUP",
  "INDEX",
  "VLOOKUP",
  "HLOOKUP",
  "LOOKUP",
  "OFFSET",
  "INDIRECT",
  "ROW",
]);
const statisticalBuiltinNames = new Set([
  "COUNTIF",
  "COUNTIFS",
  "SUMIF",
  "SUMIFS",
  "AVERAGEIF",
  "AVERAGEIFS",
  "SUMPRODUCT",
  "MINIFS",
  "MAXIFS",
  "MODE",
  "MODE.SNGL",
  "STDEV",
  "STDEV.P",
  "STDEV.S",
  "STDEVA",
  "STDEVP",
  "STDEVPA",
  "VAR",
  "VAR.P",
  "VAR.S",
  "VARA",
  "VARP",
  "VARPA",
  "SKEW",
  "SKEW.P",
  "KURT",
  "NORMDIST",
  "NORM.DIST",
  "NORMINV",
  "NORM.INV",
  "NORMSDIST",
  "NORM.S.DIST",
  "NORMSINV",
  "NORM.S.INV",
  "LOGINV",
  "LOGNORM.INV",
  "LOGNORMDIST",
  "LOGNORM.DIST",
  "CONFIDENCE.NORM",
  "CONFIDENCE",
  "ERF",
  "ERF.PRECISE",
  "ERFC",
  "ERFC.PRECISE",
  "FISHER",
  "FISHERINV",
  "GAMMALN",
  "GAMMALN.PRECISE",
  "GAMMA",
  "EXPONDIST",
  "EXPON.DIST",
  "POISSON",
  "POISSON.DIST",
  "WEIBULL",
  "WEIBULL.DIST",
  "GAMMADIST",
  "GAMMA.DIST",
  "CHIDIST",
  "CHISQ.DIST.RT",
  "CHISQ.DIST",
  "BINOMDIST",
  "BINOM.DIST",
  "BINOM.DIST.RANGE",
  "CRITBINOM",
  "BINOM.INV",
  "HYPGEOMDIST",
  "HYPGEOM.DIST",
  "NEGBINOMDIST",
  "NEGBINOM.DIST",
]);
const dynamicArrayBuiltinNames = new Set([
  "SEQUENCE",
  "EXPAND",
  "FILTER",
  "UNIQUE",
  "TAKE",
  "DROP",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "SORT",
  "SORTBY",
  "TOCOL",
  "TOROW",
  "WRAPROWS",
  "WRAPCOLS",
  "AREAS",
  "COLUMNS",
  "ROWS",
  "TRANSPOSE",
  "HSTACK",
  "VSTACK",
  "TEXTSPLIT",
  "TRIMRANGE",
]);
const lambdaBuiltinNames = new Set([
  "LET",
  "LAMBDA",
  "MAKEARRAY",
  "MAP",
  "REDUCE",
  "SCAN",
  "BYROW",
  "BYCOL",
]);

function inferCategory(name: string): BuiltinCapabilityCategory {
  if (aggregationBuiltinNames.has(name)) return "aggregation";
  if (logicalBuiltinNames.has(name)) return "logical";
  if (informationBuiltinNames.has(name)) return "information";
  if (textBuiltinNames.has(name)) return "text";
  if (dateTimeBuiltinNames.has(name)) return "date-time";
  if (lookupBuiltinNames.has(name)) return "lookup-reference";
  if (statisticalBuiltinNames.has(name)) return "statistical";
  if (dynamicArrayBuiltinNames.has(name)) return "dynamic-array";
  if (lambdaBuiltinNames.has(name)) return "lambda";
  return "math";
}

function buildCapability(
  name: string,
  id?: BuiltinId,
  jsStatus: BuiltinJsStatus = "implemented",
): BuiltinCapability {
  const category = inferCategory(name);
  return {
    ...(id === undefined ? {} : { id }),
    name,
    category,
    jsStatus,
    wasmStatus: wasmProductionBuiltinNames.has(name) ? "production" : "not-started",
    needsMetadata: false,
    needsArrayRuntime: category === "dynamic-array" || category === "lambda",
    needsExternalAdapter: false,
  };
}

export const builtinCapabilityManifest: readonly BuiltinCapability[] = [
  ...BUILTINS.map((builtin) => buildCapability(builtin.name.toUpperCase(), builtin.id)),
  ...Array.from(jsSpecialBuiltinNames, (name) =>
    buildCapability(name, undefined, "special-js-only"),
  ),
];

export const builtinCapabilitiesByName = new Map(
  builtinCapabilityManifest.map((capability) => [capability.name, capability]),
);

export const builtinWasmEnabledNames = new Set(
  builtinCapabilityManifest
    .filter((capability) => capability.wasmStatus === "production")
    .map((capability) => capability.name),
);

export const builtinJsSpecialNames = new Set(
  builtinCapabilityManifest
    .filter((capability) => capability.jsStatus === "special-js-only")
    .map((capability) => capability.name),
);

export function getBuiltinCapability(name: string): BuiltinCapability | undefined {
  return builtinCapabilitiesByName.get(name.trim().toUpperCase());
}
