export const compatibilityStatuses = [
  "unsupported",
  "seeded",
  "implemented-js",
  "implemented-js-and-wasm-shadow",
  "implemented-wasm-production",
  "blocked",
] as const;

export type CompatibilityStatus = (typeof compatibilityStatuses)[number];

export const compatibilityScopes = ["canonical", "extended"] as const;

export type CompatibilityScope = (typeof compatibilityScopes)[number];

export const wasmCompatibilityStatuses = [
  "not-started",
  "shadow",
  "production",
  "blocked",
] as const;

export type WasmCompatibilityStatus = (typeof wasmCompatibilityStatuses)[number];

export const compatibilityFamilies = [
  "arithmetic",
  "comparison",
  "logical",
  "aggregation",
  "math",
  "text",
  "date-time",
  "lookup-reference",
  "statistical",
  "information",
  "dynamic-array",
  "names",
  "tables",
  "structured-reference",
  "volatile",
  "lambda",
] as const;

export type CompatibilityFamily = (typeof compatibilityFamilies)[number];

export interface FormulaCompatibilityEntry {
  id: string;
  family: CompatibilityFamily;
  formula: string;
  status: CompatibilityStatus;
  scope: CompatibilityScope;
  fixtureIds: readonly string[];
  owner: string;
  prerequisites: readonly string[];
  wasmStatus: WasmCompatibilityStatus;
  notes?: string;
}

const familyOwners: Record<CompatibilityFamily, string> = {
  arithmetic: "lane-a-arithmetic-aggregation",
  comparison: "backbone-value-model",
  logical: "lane-b-logical-information",
  aggregation: "lane-a-arithmetic-aggregation",
  math: "lane-a-arithmetic-aggregation",
  text: "lane-c-text",
  "date-time": "lane-d-date-time",
  "lookup-reference": "lane-e-lookup-reference",
  statistical: "lane-f-statistical-financial",
  information: "lane-b-logical-information",
  "dynamic-array": "lane-g-dynamic-array-metadata",
  names: "lane-g-dynamic-array-metadata",
  tables: "lane-g-dynamic-array-metadata",
  "structured-reference": "lane-g-dynamic-array-metadata",
  volatile: "lane-d-date-time",
  lambda: "lane-h-lambda",
};

const familyPrerequisites: Record<CompatibilityFamily, readonly string[]> = {
  arithmetic: ["core:value-model"],
  comparison: ["core:value-model", "core:comparison-model"],
  logical: ["core:value-model", "core:error-model"],
  aggregation: ["core:value-model", "core:range-iterators"],
  math: ["core:value-model"],
  text: ["core:value-model", "core:string-coercion"],
  "date-time": ["core:value-model", "core:date-serial-model"],
  "lookup-reference": ["core:reference-model", "core:comparison-model"],
  statistical: ["core:value-model", "core:criteria-model"],
  information: ["core:value-model", "core:error-model"],
  "dynamic-array": ["core:shape-model", "core:spill-model"],
  names: ["core:names-model"],
  tables: ["core:tables-model"],
  "structured-reference": ["core:tables-model", "core:structured-reference-model"],
  volatile: ["core:volatile-context"],
  lambda: ["core:lambda-runtime", "core:shape-model"],
};

export function deriveWasmStatus(status: CompatibilityStatus): WasmCompatibilityStatus {
  switch (status) {
    case "unsupported":
    case "seeded":
    case "implemented-js":
      return "not-started";
    case "implemented-js-and-wasm-shadow":
      return "shadow";
    case "implemented-wasm-production":
      return "production";
    case "blocked":
      return "blocked";
  }
}

function entry(
  id: string,
  family: CompatibilityFamily,
  formula: string,
  status: CompatibilityStatus,
  options: {
    scope?: CompatibilityScope;
    fixtureIds?: readonly string[];
    owner?: string;
    prerequisites?: readonly string[];
    wasmStatus?: WasmCompatibilityStatus;
    notes?: string;
  } = {},
): FormulaCompatibilityEntry {
  const base: FormulaCompatibilityEntry = {
    id,
    family,
    formula,
    status,
    scope: options.scope ?? "canonical",
    fixtureIds: options.fixtureIds ?? [id],
    owner: options.owner ?? familyOwners[family],
    prerequisites: options.prerequisites ?? familyPrerequisites[family],
    wasmStatus: options.wasmStatus ?? deriveWasmStatus(status),
  };
  return options.notes === undefined ? base : { ...base, notes: options.notes };
}

export const formulaCompatibilityRegistry: readonly FormulaCompatibilityEntry[] = [
  entry("arithmetic:add-basic", "arithmetic", "=A1+A2", "implemented-wasm-production"),
  entry("arithmetic:precedence-basic", "arithmetic", "=A1+A2*A3", "implemented-wasm-production"),
  entry("arithmetic:unary-negation", "arithmetic", "=-A1", "implemented-wasm-production"),
  entry("arithmetic:division-basic", "arithmetic", "=A1/A2", "implemented-wasm-production"),
  entry("arithmetic:power-basic", "arithmetic", "=A1^A2", "implemented-wasm-production"),
  entry("arithmetic:percent-operator", "arithmetic", "=A1*10%", "implemented-wasm-production", {
    notes:
      "Postfix percent is in the canonical formula corpus and tracked as part of the arithmetic lane.",
  }),
  entry("comparison:equality-number", "comparison", "=A1=A2", "implemented-wasm-production"),
  entry(
    "comparison:equality-text",
    "comparison",
    '="hello"="HELLO"',
    "implemented-wasm-production",
  ),
  entry("comparison:greater-than", "comparison", "=A1>A2", "implemented-wasm-production"),
  entry("comparison:less-than-or-equal", "comparison", "=A1<=A2", "implemented-wasm-production"),
  entry("logical:if-basic", "logical", '=IF(A1>0,"yes","no")', "implemented-wasm-production"),
  entry(
    "logical:ifs-basic",
    "logical",
    '=IFS(A1>1,"big",TRUE(),"small")',
    "implemented-wasm-production",
  ),
  entry("logical:and-basic", "logical", "=AND(A1,A2)", "implemented-wasm-production"),
  entry("logical:or-basic", "logical", "=OR(A1,A2)", "implemented-wasm-production"),
  entry("logical:not-basic", "logical", "=NOT(A1)", "implemented-wasm-production"),
  entry(
    "logical:switch-basic",
    "logical",
    '=SWITCH(A1,1,"one","other")',
    "implemented-wasm-production",
  ),
  entry(
    "logical:xor-basic",
    "logical",
    "=XOR(TRUE(),FALSE(),TRUE())",
    "implemented-wasm-production",
  ),
  entry("aggregation:sum-range", "aggregation", "=SUM(A1:A3)", "implemented-wasm-production"),
  entry("aggregation:avg-range", "aggregation", "=AVG(A1:A3)", "implemented-wasm-production"),
  entry("aggregation:min-range", "aggregation", "=MIN(A1:A3)", "implemented-wasm-production"),
  entry("aggregation:max-range", "aggregation", "=MAX(A1:A3)", "implemented-wasm-production"),
  entry("aggregation:count-range", "aggregation", "=COUNT(A1:A4)", "implemented-wasm-production"),
  entry("aggregation:counta-range", "aggregation", "=COUNTA(A1:A4)", "implemented-wasm-production"),
  entry(
    "aggregation:countblank-range",
    "aggregation",
    "=COUNTBLANK(A1:A4)",
    "implemented-wasm-production",
  ),
  entry("math:abs-basic", "math", "=ABS(A1)", "implemented-wasm-production"),
  entry("math:round-basic", "math", "=ROUND(A1,1)", "implemented-wasm-production"),
  entry("math:trunc-basic", "math", "=TRUNC(A1,1)", "implemented-wasm-production"),
  entry("math:floor-basic", "math", "=FLOOR(A1,2)", "implemented-wasm-production"),
  entry("math:floor-math-basic", "math", "=FLOOR.MATH(A1,2)", "implemented-wasm-production"),
  entry("math:floor-precise-basic", "math", "=FLOOR.PRECISE(A1,2)", "implemented-wasm-production"),
  entry("math:ceiling-basic", "math", "=CEILING(A1,2)", "implemented-wasm-production"),
  entry("math:ceiling-math-basic", "math", "=CEILING.MATH(A1,2)", "implemented-wasm-production"),
  entry(
    "math:ceiling-precise-basic",
    "math",
    "=CEILING.PRECISE(A1,2)",
    "implemented-wasm-production",
  ),
  entry("math:iso-ceiling-basic", "math", "=ISO.CEILING(A1,2)", "implemented-wasm-production"),
  entry("math:mod-basic", "math", "=MOD(A1,A2)", "implemented-wasm-production"),
  entry("math:bitand-basic", "math", "=BITAND(6,3)", "implemented-wasm-production"),
  entry("math:base-basic", "math", "=BASE(255,16,4)", "implemented-wasm-production"),
  entry("math:decimal-basic", "math", '=DECIMAL("00FF",16)', "implemented-wasm-production"),
  entry("math:bin2dec-basic", "math", '=BIN2DEC("1111111111")', "implemented-wasm-production"),
  entry("math:dec2bin-basic", "math", "=DEC2BIN(10,8)", "implemented-wasm-production"),
  entry("math:oct2hex-basic", "math", '=OCT2HEX("17",4)', "implemented-wasm-production"),
  entry("math:besseli-basic", "math", "=BESSELI(1.5,1)", "implemented-wasm-production"),
  entry("math:besselj-basic", "math", "=BESSELJ(1.9,2)", "implemented-wasm-production"),
  entry("math:besselk-basic", "math", "=BESSELK(1.5,1)", "implemented-wasm-production"),
  entry("math:bessely-basic", "math", "=BESSELY(2.5,1)", "implemented-wasm-production"),
  entry("math:convert-basic", "math", '=CONVERT(6,"mi","km")', "implemented-wasm-production"),
  entry(
    "math:euroconvert-basic",
    "math",
    '=EUROCONVERT(1,"FRF","DEM",TRUE,3)',
    "implemented-wasm-production",
  ),
  entry("math:acosh-basic", "math", "=ACOSH(1)", "implemented-wasm-production"),
  entry("math:fact-basic", "math", "=FACT(5)", "implemented-wasm-production"),
  entry("math:combin-basic", "math", "=COMBIN(8,3)", "implemented-wasm-production"),
  entry("math:permut-basic", "math", "=PERMUT(5,3)", "implemented-wasm-production"),
  entry("math:permutationa-basic", "math", "=PERMUTATIONA(2,3)", "implemented-wasm-production"),
  entry("math:mround-basic", "math", "=MROUND(A1,4)", "implemented-wasm-production"),
  entry("math:seriessum-basic", "math", "=SERIESSUM(2,1,2,1,2)", "implemented-wasm-production"),
  entry("math:gcd-basic", "math", "=GCD(A1:A3)", "implemented-wasm-production"),
  entry("math:product-basic", "math", "=PRODUCT(A1:A3)", "implemented-wasm-production"),
  entry("math:geomean-basic", "math", "=GEOMEAN(A1:A3)", "implemented-wasm-production"),
  entry("math:harmean-basic", "math", "=HARMEAN(A1:A3)", "implemented-wasm-production"),
  entry("math:sqrtpi-basic", "math", "=SQRTPI(A1)", "implemented-wasm-production"),
  entry("math:sumsq-basic", "math", "=SUMSQ(A1:A3)", "implemented-wasm-production"),
  entry("text:concat-operator", "text", '="bi"&"lig"', "implemented-wasm-production"),
  entry("text:concat-function", "text", '=CONCAT("bi","lig")', "implemented-wasm-production"),
  entry("text:len-basic", "text", '=LEN("bilig")', "implemented-wasm-production", {
    notes: "LEN now executes through the string-aware native runtime for scalar inputs.",
  }),
  entry("text:char-basic", "text", "=CHAR(65)", "implemented-wasm-production"),
  entry("text:code-basic", "text", '=CODE("A")', "implemented-wasm-production"),
  entry(
    "text:clean-basic",
    "text",
    "=CLEAN(CHAR(97)&CHAR(1)&CHAR(98))",
    "implemented-wasm-production",
  ),
  entry("text:asc-basic", "text", '=ASC("ＡＢＣ　１２３")', "implemented-wasm-production"),
  entry("text:jis-basic", "text", '=JIS("ABC 123")', "implemented-wasm-production"),
  entry("text:dbcs-basic", "text", '=DBCS("ABC 123")', "implemented-wasm-production"),
  entry("text:unichar-basic", "text", "=UNICHAR(66)", "implemented-wasm-production"),
  entry("text:dollar-basic", "text", "=DOLLAR(-1234.5,1)", "implemented-wasm-production"),
  entry("text:text-basic", "text", '=TEXT(1234.567,"#,##0.00")', "implemented-wasm-production"),
  entry(
    "text:text-date-basic",
    "text",
    '=TEXT(DATE(2024,3,5),"yyyy-mm-dd")',
    "implemented-wasm-production",
  ),
  entry("text:phonetic-basic", "text", '=PHONETIC("カタカナ")', "implemented-wasm-production"),
  entry("text:bahttext-basic", "text", "=BAHTTEXT(1234)", "implemented-wasm-production"),
  entry("information:t-basic", "information", '=T("alpha")', "implemented-wasm-production"),
  entry("information:n-basic", "information", "=N(TRUE())", "implemented-wasm-production"),
  entry("information:type-basic", "information", '=TYPE("alpha")', "implemented-wasm-production"),
  entry("math:delta-basic", "math", "=DELTA(4,4)", "implemented-wasm-production"),
  entry("math:gestep-basic", "math", "=GESTEP(-1)", "implemented-wasm-production"),
  entry("statistical:gauss-basic", "statistical", "=GAUSS(0)", "implemented-wasm-production"),
  entry("statistical:phi-basic", "statistical", "=PHI(0)", "implemented-wasm-production"),
  entry(
    "statistical:standardize-basic",
    "statistical",
    "=STANDARDIZE(1,0,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:confidence-norm-basic",
    "statistical",
    "=CONFIDENCE.NORM(0.05,1,100)",
    "implemented-wasm-production",
  ),
  entry("statistical:mode-basic", "statistical", "=MODE(A1:A6)", "implemented-wasm-production"),
  entry(
    "statistical:mode-sngl-basic",
    "statistical",
    "=MODE.SNGL(A1:A6)",
    "implemented-wasm-production",
  ),
  entry("statistical:stdev-basic", "statistical", "=STDEV(A1:A4)", "implemented-wasm-production"),
  entry(
    "statistical:stdeva-basic",
    "statistical",
    '=STDEVA(2,TRUE(),"skip")',
    "implemented-wasm-production",
  ),
  entry("statistical:var-basic", "statistical", "=VAR(A1:A4)", "implemented-wasm-production"),
  entry(
    "statistical:vara-basic",
    "statistical",
    '=VARA(2,TRUE(),"skip")',
    "implemented-wasm-production",
  ),
  entry("statistical:skew-basic", "statistical", "=SKEW(A1:A5)", "implemented-wasm-production"),
  entry("statistical:kurt-basic", "statistical", "=KURT(A1:A5)", "implemented-wasm-production"),
  entry(
    "statistical:normdist-basic",
    "statistical",
    "=NORMDIST(1,0,1,TRUE)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:norminv-basic",
    "statistical",
    "=NORMINV(0.8413447460685429,0,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:normsdist-basic",
    "statistical",
    "=NORMSDIST(1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:normsinv-basic",
    "statistical",
    "=NORMSINV(0.001)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:loginv-basic",
    "statistical",
    "=LOGINV(0.5,0,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:lognormdist-basic",
    "statistical",
    "=LOGNORMDIST(1,0,1)",
    "implemented-wasm-production",
  ),
  entry("date-time:serial-addition", "date-time", "=A1+7", "implemented-wasm-production"),
  entry(
    "date-time:date-constructor",
    "date-time",
    "=DATE(2026,3,15)",
    "implemented-wasm-production",
  ),
  entry("date-time:today-volatile", "date-time", "=TODAY()", "implemented-wasm-production", {
    prerequisites: ["core:value-model", "core:date-serial-model", "core:volatile-context"],
  }),
  entry(
    "lookup-reference:address-basic",
    "lookup-reference",
    "=ADDRESS(12,3)",
    "implemented-wasm-production",
  ),
  entry(
    "lookup-reference:index-basic",
    "lookup-reference",
    "=INDEX(A1:B2,2,1)",
    "implemented-wasm-production",
  ),
  entry(
    "lookup-reference:match-exact",
    "lookup-reference",
    '=MATCH("pear",A1:A3,0)',
    "implemented-wasm-production",
  ),
  entry(
    "lookup-reference:vlookup-exact",
    "lookup-reference",
    '=VLOOKUP("pear",A1:B3,2,FALSE)',
    "implemented-wasm-production",
  ),
  entry(
    "lookup-reference:xlookup-exact",
    "lookup-reference",
    '=XLOOKUP("pear",A1:A3,B1:B3)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:averageif-basic",
    "statistical",
    '=AVERAGEIF(A1:A4,">0")',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:countif-basic",
    "statistical",
    '=COUNTIF(A1:A4,">0")',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:chisqdist-basic",
    "statistical",
    "=CHISQDIST(18.307,10)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:chiinv-basic",
    "statistical",
    "=CHIINV(0.050001,10)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:chisq-inv-rt-basic",
    "statistical",
    "=CHISQ.INV.RT(0.050001,10)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:chisqinv-basic",
    "statistical",
    "=CHISQINV(0.050001,10)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:chisq-inv-basic",
    "statistical",
    "=CHISQ.INV(0.93,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:chisq-test-basic",
    "statistical",
    "=CHISQ.TEST(A1:B3,D1:E3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:beta-dist-basic",
    "statistical",
    "=BETA.DIST(2,8,10,TRUE,1,3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:beta-inv-basic",
    "statistical",
    "=BETA.INV(0.6854705810117458,8,10,1,3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:f-dist-rt-basic",
    "statistical",
    "=F.DIST.RT(15.2068649,6,4)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:fdist-basic",
    "statistical",
    "=FDIST(15.2068649,6,4)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:f-inv-basic",
    "statistical",
    "=F.INV(0.01,6,4)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:f-inv-rt-basic",
    "statistical",
    "=F.INV.RT(0.01,6,4)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:f-test-basic",
    "statistical",
    "=F.TEST(A1:A5,B1:B5)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:z-test-basic",
    "statistical",
    "=Z.TEST(D1:D5,2,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:correl-basic",
    "statistical",
    "=CORREL(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:covar-basic",
    "statistical",
    "=COVAR(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:covariance-p-basic",
    "statistical",
    "=COVARIANCE.P(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:covariance-s-basic",
    "statistical",
    "=COVARIANCE.S(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:pearson-basic",
    "statistical",
    "=PEARSON(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:intercept-basic",
    "statistical",
    "=INTERCEPT(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:slope-basic",
    "statistical",
    "=SLOPE(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry("statistical:rsq-basic", "statistical", "=RSQ(A1:A3,B1:B3)", "implemented-wasm-production"),
  entry(
    "statistical:steyx-basic",
    "statistical",
    "=STEYX(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry("statistical:rank-basic", "statistical", "=RANK(20,A1:A4)", "implemented-wasm-production"),
  entry(
    "statistical:rank-eq-basic",
    "statistical",
    "=RANK.EQ(20,A1:A4)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:rank-avg-basic",
    "statistical",
    "=RANK.AVG(20,A1:A4)",
    "implemented-wasm-production",
  ),
  entry("statistical:median-basic", "statistical", "=MEDIAN(A1:A8)", "implemented-wasm-production"),
  entry("statistical:small-basic", "statistical", "=SMALL(A1:A8,3)", "implemented-wasm-production"),
  entry("statistical:large-basic", "statistical", "=LARGE(A1:A8,2)", "implemented-wasm-production"),
  entry(
    "statistical:percentile-basic",
    "statistical",
    "=PERCENTILE(A1:A8,0.25)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:percentile-inc-basic",
    "statistical",
    "=PERCENTILE.INC(A1:A8,0.25)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:percentile-exc-basic",
    "statistical",
    "=PERCENTILE.EXC(A1:A8,0.25)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:percentrank-basic",
    "statistical",
    "=PERCENTRANK(A1:A8,8)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:percentrank-inc-basic",
    "statistical",
    "=PERCENTRANK.INC(A1:A8,8)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:percentrank-exc-basic",
    "statistical",
    "=PERCENTRANK.EXC(A1:A8,8)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:quartile-basic",
    "statistical",
    "=QUARTILE(A1:A8,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:quartile-inc-basic",
    "statistical",
    "=QUARTILE.INC(A1:A8,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:quartile-exc-basic",
    "statistical",
    "=QUARTILE.EXC(A1:A8,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:mode-mult-basic",
    "statistical",
    "=MODE.MULT(A1:A6)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:frequency-basic",
    "statistical",
    "=FREQUENCY(A1:A6,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:t-dist-basic",
    "statistical",
    "=T.DIST(1,1,TRUE)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:t-inv-2t-basic",
    "statistical",
    "=T.INV.2T(0.5,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:confidence-t-basic",
    "statistical",
    "=CONFIDENCE.T(0.5,2,4)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:gamma-inv-basic",
    "statistical",
    "=GAMMA.INV(0.08030139707139418,3,2)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:t-test-basic",
    "statistical",
    "=T.TEST(A1:A3,B1:B3,2,1)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:forecast-basic",
    "statistical",
    "=FORECAST(4,A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:forecast-linear-basic",
    "statistical",
    "=FORECAST.LINEAR(4,A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:trend-basic",
    "statistical",
    "=TREND(A1:A3,B1:B3,D1:D2)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:growth-basic",
    "statistical",
    "=GROWTH(A1:A3,B1:B3,D1:D2)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:linest-basic",
    "statistical",
    "=LINEST(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:logest-basic",
    "statistical",
    "=LOGEST(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:prob-basic",
    "statistical",
    "=PROB(A1:A4,B1:B4,2,3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:trimmean-basic",
    "statistical",
    "=TRIMMEAN(A1:A8,0.25)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:daverage-basic",
    "statistical",
    '=DAVERAGE(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dcount-basic",
    "statistical",
    '=DCOUNT(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dcounta-basic",
    "statistical",
    '=DCOUNTA(A1:C5,"Height",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dget-basic",
    "statistical",
    '=DGET(A1:C5,"Height",F1:F2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dmax-basic",
    "statistical",
    '=DMAX(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dmin-basic",
    "statistical",
    '=DMIN(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dproduct-basic",
    "statistical",
    '=DPRODUCT(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dstdev-basic",
    "statistical",
    '=DSTDEV(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dstdevp-basic",
    "statistical",
    '=DSTDEVP(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dsum-basic",
    "statistical",
    '=DSUM(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dvar-basic",
    "statistical",
    '=DVAR(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:dvarp-basic",
    "statistical",
    '=DVARP(A1:C5,"Yield",E1:E2)',
    "implemented-wasm-production",
  ),
  entry("information:isblank-basic", "information", "=ISBLANK(A1)", "implemented-wasm-production"),
  entry(
    "information:isnumber-basic",
    "information",
    "=ISNUMBER(A1)",
    "implemented-wasm-production",
  ),
  entry("information:istext-basic", "information", "=ISTEXT(A1)", "implemented-wasm-production"),
  entry(
    "dynamic-array:sequence-spill",
    "dynamic-array",
    "=SEQUENCE(3,1,1,1)",
    "implemented-wasm-production",
    {
      notes:
        "Top-level SEQUENCE spills now execute on the AssemblyScript path and reuse the existing workbook spill metadata contract; broader array families remain blocked.",
    },
  ),
  entry(
    "dynamic-array:sequence-aggregate",
    "dynamic-array",
    "=SUM(SEQUENCE(A1,1,1,1))",
    "implemented-wasm-production",
    {
      notes:
        "Numeric aggregate consumers can now read transient native SEQUENCE arrays directly on the AssemblyScript path without reviving the removed JS runtime fallback.",
    },
  ),
  entry(
    "dynamic-array:filter-basic",
    "dynamic-array",
    "=FILTER(A1:A4,A1:A4>2)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:unique-basic",
    "dynamic-array",
    "=UNIQUE(A1:A4)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:groupby-basic",
    "dynamic-array",
    "=GROUPBY(A1:A5,C1:C5,SUM,3,1)",
    "implemented-js",
    {
      notes:
        "GROUPBY now has production JS-special spill semantics and canonical fixture coverage; native grouped-array lowering is still pending.",
    },
  ),
  entry(
    "dynamic-array:pivotby-basic",
    "dynamic-array",
    "=PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)",
    "implemented-js",
    {
      notes:
        "PIVOTBY now has production JS-special pivot spill semantics and canonical fixture coverage; native grouped-array lowering is still pending.",
    },
  ),
  entry("names:defined-name-scalar", "names", "=TaxRate*A1", "implemented-wasm-production", {
    notes:
      "Scalar workbook names rebind onto the AssemblyScript path once the engine has concrete scalar metadata; reference-valued names remain blocked.",
  }),
  entry(
    "names:defined-name-case-insensitive",
    "names",
    "=taxrate*A1",
    "implemented-wasm-production",
    {
      scope: "extended",
      notes:
        "Case-insensitive scalar workbook names rebind onto the AssemblyScript path after metadata normalization.",
    },
  ),
  entry(
    "names:defined-name-multi-scalar-pack",
    "names",
    "=TaxRate+FeeRate",
    "implemented-wasm-production",
    {
      scope: "extended",
      notes:
        "Multiple scalar workbook names can participate in one AssemblyScript-routed scalar expression without widening onto reference-valued metadata.",
    },
  ),
  entry("names:defined-name-missing", "names", "=MissingRate*A1", "implemented-wasm-production", {
    scope: "extended",
    notes:
      "Missing workbook-level names now stay on the AssemblyScript path, surface #NAME?, and rebind natively once the name appears.",
  }),
  entry(
    "lookup-reference:multiple-operations-basic",
    "lookup-reference",
    "=MULTIPLE.OPERATIONS(B5,B3,C4,B2,D2)",
    "implemented-js",
    {
      scope: "extended",
      notes:
        "MULTIPLE.OPERATIONS now routes through the workbook-aware JS what-if contract and is covered by both fixture-harness and engine integration tests.",
    },
  ),
  entry(
    "tables:table-total-row-sum",
    "tables",
    "=SUM(Sales[Amount])",
    "implemented-wasm-production",
    {
      notes:
        "Table-backed aggregate formulas now compile through metadata substitution and route onto the native aggregate path once the table exists.",
    },
  ),
  entry(
    "structured-reference:table-column-ref",
    "structured-reference",
    "=SUM(Sales[Amount])",
    "implemented-wasm-production",
    {
      notes:
        "Structured column references now compile through metadata substitution and route onto the native aggregate path once the table exists.",
    },
  ),
  entry("volatile:rand-basic", "volatile", "=RAND()", "implemented-wasm-production", {
    prerequisites: ["core:volatile-context", "core:value-model"],
  }),
  entry("lambda:let-basic", "lambda", "=LET(x,2,x+3)", "implemented-wasm-production", {
    notes:
      "LET formulas with rewrite-safe bindings now lower to ordinary scalar AST before binding, which lets deterministic cases route through the native fast path without a general closure VM.",
  }),
  entry("lambda:lambda-invoke", "lambda", "=LAMBDA(x,x+1)(4)", "implemented-wasm-production", {
    notes:
      "Immediate LAMBDA invocation now rewrites to the invoked body with argument substitution before binding, so scalar deterministic cases compile onto the native path.",
  }),
  entry("lambda:map-basic", "lambda", "=MAP(A1:A3,LAMBDA(x,x*2))", "implemented-wasm-production", {
    notes:
      "MAP calls whose lambda body rewrites to an ordinary broadcasted array expression now lower before binding and execute on the native spill path.",
  }),
  entry("logical:if-true-branch", "logical", "=IF(A1,A2,A3)", "implemented-wasm-production"),
  entry("logical:if-condition-error", "logical", "=IF(1/0,1,2)", "implemented-wasm-production", {
    notes: "The native branch VM now propagates IF condition errors before either branch executes.",
  }),
  entry(
    "logical:iferror-catches-any-error",
    "logical",
    '=IFERROR(1/0,"fallback")',
    "implemented-wasm-production",
  ),
  entry(
    "logical:ifna-catches-na-only",
    "logical",
    '=IFNA(NA(),"missing")',
    "implemented-wasm-production",
  ),
  entry("logical:and-false-on-empty", "logical", "=AND(TRUE,A1)", "implemented-wasm-production"),
  entry("logical:or-true-branch", "logical", "=OR(A1,TRUE)", "implemented-wasm-production"),
  entry("logical:not-number", "logical", "=NOT(2)", "implemented-wasm-production"),
  entry("information:isblank-empty", "information", "=ISBLANK(A1)", "implemented-wasm-production"),
  entry(
    "information:isnumber-number",
    "information",
    "=ISNUMBER(42)",
    "implemented-wasm-production",
  ),
  entry(
    "information:istext-string",
    "information",
    '=ISTEXT("hello")',
    "implemented-wasm-production",
  ),
  entry(
    "text:len-counts-plain-string-length",
    "text",
    '=LEN("hello")',
    "implemented-wasm-production",
  ),
  entry("text:exact-basic", "text", '=EXACT("Alpha","alpha")', "implemented-wasm-production", {
    notes: "EXACT now routes through the string-aware WASM runtime.",
  }),
  entry("text:left-basic", "text", '=LEFT("alpha",3)', "implemented-wasm-production"),
  entry("text:right-basic", "text", '=RIGHT("alpha",2)', "implemented-wasm-production"),
  entry("text:mid-basic", "text", '=MID("alphabet",2,3)', "implemented-wasm-production"),
  entry("text:trim-basic", "text", '=TRIM("  alpha   beta  ")', "implemented-wasm-production"),
  entry("text:upper-basic", "text", '=UPPER("Alpha beta")', "implemented-wasm-production"),
  entry("text:lower-basic", "text", '=LOWER("Alpha BETA")', "implemented-wasm-production"),
  entry("text:find-basic", "text", '=FIND("ph","alphabet")', "implemented-wasm-production"),
  entry("text:search-basic", "text", '=SEARCH("PH","alphabet")', "implemented-wasm-production"),
  entry(
    "text:textbefore-basic",
    "text",
    '=TEXTBEFORE("alpha-beta","-")',
    "implemented-wasm-production",
  ),
  entry(
    "text:textafter-basic",
    "text",
    '=TEXTAFTER("alpha-beta","-")',
    "implemented-wasm-production",
  ),
  entry("text:textjoin-basic", "text", '=TEXTJOIN("-",TRUE,A1:A3)', "implemented-wasm-production"),
  entry("text:textsplit-basic", "text", '=TEXTSPLIT(A1,",","|")', "implemented-wasm-production"),
  entry("text:value-basic", "text", '=VALUE("42")', "implemented-wasm-production", {
    notes:
      "VALUE now coerces scalar text inputs on the AssemblyScript path, including trimmed decimals and exponent forms.",
  }),
  entry(
    "lookup-reference:choose-basic",
    "lookup-reference",
    '=CHOOSE(2,"red","blue","green")',
    "implemented-wasm-production",
  ),
  entry(
    "lookup-reference:xmatch-basic",
    "lookup-reference",
    '=XMATCH("pear",A1:A3,0)',
    "implemented-wasm-production",
  ),
  entry(
    "lookup-reference:hlookup-basic",
    "lookup-reference",
    '=HLOOKUP("pear",A1:C2,2,FALSE)',
    "implemented-wasm-production",
  ),
  entry(
    "lookup-reference:offset-basic",
    "lookup-reference",
    "=OFFSET(A1,1,1)",
    "implemented-wasm-production",
    {
      notes:
        "OFFSET now executes on the AssemblyScript path for numeric inputs, including simple in-bounds offset ranges.",
    },
  ),
  entry(
    "dynamic-array:take-basic",
    "dynamic-array",
    "=TAKE(A1:A4,2)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:drop-basic",
    "dynamic-array",
    "=DROP(A1:A4,2)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:choosecols-basic",
    "dynamic-array",
    "=CHOOSECOLS(A1:C2,1,3)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:chooserows-basic",
    "dynamic-array",
    "=CHOOSEROWS(A1:B3,1,3)",
    "implemented-wasm-production",
  ),
  entry(
    "statistical:sumif-basic",
    "statistical",
    '=SUMIF(A1:A4,">0",B1:B4)',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:sumifs-basic",
    "statistical",
    '=SUMIFS(C1:C4,A1:A4,">0",B1:B4,"x")',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:averageifs-basic",
    "statistical",
    '=AVERAGEIFS(C1:C4,A1:A4,">0",B1:B4,"x")',
    "implemented-wasm-production",
  ),
  entry(
    "statistical:countifs-basic",
    "statistical",
    '=COUNTIFS(A1:A4,">0",B1:B4,"x")',
    "implemented-wasm-production",
  ),
  entry("math:sumproduct-basic", "math", "=SUMPRODUCT(A1:A3,B1:B3)", "implemented-wasm-production"),
  entry("math:int-basic", "math", "=INT(-3.1)", "implemented-wasm-production"),
  entry("math:roundup-basic", "math", "=ROUNDUP(12.341,2)", "implemented-wasm-production"),
  entry("math:rounddown-basic", "math", "=ROUNDDOWN(12.349,2)", "implemented-wasm-production"),
  entry(
    "arithmetic:cross-sheet-multiply",
    "arithmetic",
    "=Sheet2!B1*3",
    "implemented-wasm-production",
    {
      scope: "extended",
      prerequisites: ["core:value-model", "core:reference-model"],
      notes:
        "Qualified scalar references stay on the native arithmetic path once the target sheet is present.",
    },
  ),
  entry(
    "arithmetic:cross-sheet-empty-cell-zero",
    "arithmetic",
    "=Sheet2!B1*3",
    "implemented-wasm-production",
    {
      scope: "extended",
      prerequisites: ["core:value-model", "core:reference-model"],
      notes:
        "Existing blank cells on another sheet coerce through the usual arithmetic empty-cell semantics.",
    },
  ),
  entry(
    "arithmetic:missing-sheet-ref-error",
    "arithmetic",
    "=Sheet2!B1*3",
    "implemented-wasm-production",
    {
      scope: "extended",
      prerequisites: ["core:value-model", "core:reference-model"],
      notes:
        "Unresolved qualified cells now stay on the native path via explicit unresolved-ref operands that emit #REF! until rebinding can occur.",
    },
  ),
  entry("date-time:now-volatile", "date-time", "=NOW()", "implemented-wasm-production", {
    prerequisites: ["core:value-model", "core:date-serial-model", "core:volatile-context"],
    notes:
      "NOW now captures a single recalc-epoch serial on the host and executes on the AssemblyScript path.",
  }),
  entry("date-time:time-basic", "date-time", "=TIME(12,30,0)", "implemented-wasm-production"),
  entry("date-time:hour-basic", "date-time", "=HOUR(A1)", "implemented-wasm-production"),
  entry("date-time:minute-basic", "date-time", "=MINUTE(A1)", "implemented-wasm-production"),
  entry("date-time:second-basic", "date-time", "=SECOND(A1)", "implemented-wasm-production"),
  entry(
    "date-time:datedif-ym",
    "date-time",
    '=DATEDIF(DATE(2020,1,15),DATE(2021,3,20),"YM")',
    "implemented-wasm-production",
  ),
  entry(
    "date-time:days360-basic",
    "date-time",
    "=DAYS360(DATE(2024,1,29),DATE(2024,3,31))",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:isoweeknum-basic",
    "date-time",
    "=ISOWEEKNUM(DATE(2024,1,1))",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:workday-intl-basic",
    "date-time",
    "=WORKDAY.INTL(A1,2,7,B1)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:timevalue-basic",
    "date-time",
    '=TIMEVALUE("1:30 PM")',
    "implemented-wasm-production",
  ),
  entry(
    "date-time:yearfrac-basic",
    "date-time",
    "=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:fvschedule-basic",
    "date-time",
    "=FVSCHEDULE(1000,0.09,0.11,0.1)",
    "implemented-wasm-production",
  ),
  entry("date-time:effect-basic", "date-time", "=EFFECT(12%,12)", "implemented-wasm-production"),
  entry(
    "date-time:nominal-basic",
    "date-time",
    "=NOMINAL(0.12682503013196977,12)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:pduration-basic",
    "date-time",
    "=PDURATION(10%,100,121)",
    "implemented-wasm-production",
  ),
  entry("date-time:rri-basic", "date-time", "=RRI(2,100,121)", "implemented-wasm-production"),
  entry("date-time:fv-basic", "date-time", "=FV(10%,2,-100,-1000)", "implemented-wasm-production"),
  entry(
    "date-time:pv-basic",
    "date-time",
    "=PV(10%,2,-576.1904761904761)",
    "implemented-wasm-production",
  ),
  entry("date-time:pmt-basic", "date-time", "=PMT(10%,2,1000)", "implemented-wasm-production"),
  entry(
    "date-time:nper-basic",
    "date-time",
    "=NPER(10%,-576.1904761904761,1000)",
    "implemented-wasm-production",
  ),
  entry("date-time:npv-basic", "date-time", "=NPV(10%,100,200,300)", "implemented-wasm-production"),
  entry("date-time:rate-basic", "date-time", "=RATE(48,-200,8000)", "implemented-wasm-production"),
  entry("date-time:irr-basic", "date-time", "=IRR(A1:A6)", "implemented-wasm-production"),
  entry("date-time:mirr-basic", "date-time", "=MIRR(A1:A6,10%,12%)", "implemented-wasm-production"),
  entry(
    "date-time:xnpv-basic",
    "date-time",
    "=XNPV(0.09,A1:A5,B1:B5)",
    "implemented-wasm-production",
  ),
  entry("date-time:xirr-basic", "date-time", "=XIRR(A1:A5,B1:B5)", "implemented-wasm-production"),
  entry("date-time:ipmt-basic", "date-time", "=IPMT(10%,1,2,1000)", "implemented-wasm-production"),
  entry("date-time:ppmt-basic", "date-time", "=PPMT(10%,1,2,1000)", "implemented-wasm-production"),
  entry(
    "date-time:ispmt-basic",
    "date-time",
    "=ISPMT(10%,1,2,1000)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:cumipmt-basic",
    "date-time",
    "=CUMIPMT(9%/12,30*12,125000,13,24,0)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:cumprinc-basic",
    "date-time",
    "=CUMPRINC(9%/12,30*12,125000,13,24,0)",
    "implemented-wasm-production",
  ),
  entry("date-time:db-basic", "date-time", "=DB(10000,1000,5,1)", "implemented-wasm-production"),
  entry("date-time:ddb-basic", "date-time", "=DDB(2400,300,10,2)", "implemented-wasm-production"),
  entry("date-time:vdb-basic", "date-time", "=VDB(2400,300,10,1,3)", "implemented-wasm-production"),
  entry("date-time:sln-basic", "date-time", "=SLN(10000,1000,9)", "implemented-wasm-production"),
  entry("date-time:syd-basic", "date-time", "=SYD(10000,1000,9,1)", "implemented-wasm-production"),
  entry(
    "date-time:disc-basic",
    "date-time",
    "=DISC(DATE(2023,1,1),DATE(2023,4,1),97,100,2)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:intrate-basic",
    "date-time",
    "=INTRATE(DATE(2023,1,1),DATE(2023,4,1),1000,1030,2)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:received-basic",
    "date-time",
    "=RECEIVED(DATE(2023,1,1),DATE(2023,4,1),1000,0.12,2)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:pricedisc-basic",
    "date-time",
    "=PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:yielddisc-basic",
    "date-time",
    "=YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:pricemat-basic",
    "date-time",
    "=PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:yieldmat-basic",
    "date-time",
    "=YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:oddfprice-basic",
    "date-time",
    "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:oddfyield-basic",
    "date-time",
    "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:oddlprice-basic",
    "date-time",
    "=ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:oddlyield-basic",
    "date-time",
    "=ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:coupdaybs-basic",
    "date-time",
    "=COUPDAYBS(DATE(2007,1,25),DATE(2009,11,15),2,4)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:coupdays-basic",
    "date-time",
    "=COUPDAYS(DATE(2007,1,25),DATE(2009,11,15),2,4)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:coupdaysnc-basic",
    "date-time",
    "=COUPDAYSNC(DATE(2007,1,25),DATE(2009,11,15),2,4)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:coupncd-basic",
    "date-time",
    "=COUPNCD(DATE(2007,1,25),DATE(2009,11,15),2,4)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:coupnum-basic",
    "date-time",
    "=COUPNUM(DATE(2007,1,25),DATE(2009,11,15),2,4)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:couppcd-basic",
    "date-time",
    "=COUPPCD(DATE(2007,1,25),DATE(2009,11,15),2,4)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:price-basic",
    "date-time",
    "=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:yield-basic",
    "date-time",
    "=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:duration-basic",
    "date-time",
    "=DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:mduration-basic",
    "date-time",
    "=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:tbillprice-basic",
    "date-time",
    "=TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:tbillyield-basic",
    "date-time",
    "=TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:tbilleq-basic",
    "date-time",
    "=TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:networkdays-intl-basic",
    "date-time",
    "=NETWORKDAYS.INTL(A1,A2,7,B1)",
    "implemented-wasm-production",
  ),
  entry(
    "date-time:weekday-basic",
    "date-time",
    "=WEEKDAY(DATE(2026,3,15))",
    "implemented-wasm-production",
  ),
  entry(
    "aggregation:cross-sheet-range-sum",
    "aggregation",
    "=SUM(Sheet2!A1:A2)",
    "implemented-wasm-production",
    {
      scope: "extended",
      prerequisites: ["core:value-model", "core:range-iterators", "core:reference-model"],
      notes: "Resolved qualified ranges now stay on the native aggregation path.",
    },
  ),
  entry(
    "aggregation:cross-sheet-empty-range-zero",
    "aggregation",
    "=SUM(Sheet2!A1:A2)",
    "implemented-wasm-production",
    {
      scope: "extended",
      prerequisites: ["core:value-model", "core:range-iterators", "core:reference-model"],
      notes:
        "Existing blank ranges on another sheet aggregate as zero once the referenced sheet exists.",
    },
  ),
  entry(
    "aggregation:missing-sheet-range-ref-error",
    "aggregation",
    "=SUM(Sheet2!A1:A2)",
    "implemented-wasm-production",
    {
      scope: "extended",
      prerequisites: ["core:value-model", "core:range-iterators", "core:reference-model"],
      notes:
        "Missing qualified ranges now stay on the native path via explicit unresolved-range operands that emit #REF! until later rebinding.",
    },
  ),
  entry("dynamic-array:sort-basic", "dynamic-array", "=SORT(A1:A4)", "implemented-wasm-production"),
  entry(
    "dynamic-array:sortby-basic",
    "dynamic-array",
    "=SORTBY(A1:A3,B1:B3)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:tocol-basic",
    "dynamic-array",
    "=TOCOL(A1:B2)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:torow-basic",
    "dynamic-array",
    "=TOROW(A1:B2)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:wraprows-basic",
    "dynamic-array",
    "=WRAPROWS(A1:A4,2)",
    "implemented-wasm-production",
  ),
  entry(
    "dynamic-array:wrapcols-basic",
    "dynamic-array",
    "=WRAPCOLS(A1:A4,2)",
    "implemented-wasm-production",
  ),
  entry("names:defined-name-range", "names", "=SUM(MyRange)", "implemented-wasm-production", {
    notes:
      "Range-valued workbook names now compile through metadata substitution and route onto the native aggregate path once the name resolves.",
  }),
  entry(
    "lambda:byrow-basic",
    "lambda",
    "=BYROW(A1:B2,LAMBDA(r,SUM(r)))",
    "implemented-wasm-production",
    {
      notes:
        "BYROW aggregate lambdas in the canonical SUM form now lower onto an internal native row-sum builtin, so the canonical spill case executes on the wasm path.",
    },
  ),
];

export function getCompatibilityEntry(id: string): FormulaCompatibilityEntry | undefined {
  return formulaCompatibilityRegistry.find((compatibilityEntry) => compatibilityEntry.id === id);
}

export function isCompatibilityStatus(value: string): value is CompatibilityStatus {
  return (compatibilityStatuses as readonly string[]).includes(value);
}

export function isWasmCompatibilityStatus(value: string): value is WasmCompatibilityStatus {
  return (wasmCompatibilityStatuses as readonly string[]).includes(value);
}
