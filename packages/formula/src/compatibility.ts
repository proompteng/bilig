export const compatibilityStatuses = [
  "unsupported",
  "seeded",
  "implemented-js",
  "implemented-js-and-wasm-shadow",
  "implemented-wasm-production",
  "blocked"
] as const;

export type CompatibilityStatus = typeof compatibilityStatuses[number];

export const compatibilityScopes = ["canonical", "extended"] as const;

export type CompatibilityScope = typeof compatibilityScopes[number];

export const wasmCompatibilityStatuses = ["not-started", "shadow", "production", "blocked"] as const;

export type WasmCompatibilityStatus = typeof wasmCompatibilityStatuses[number];

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
  "lambda"
] as const;

export type CompatibilityFamily = typeof compatibilityFamilies[number];

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
  lambda: "lane-h-lambda"
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
  lambda: ["core:lambda-runtime", "core:shape-model"]
};

function deriveWasmStatus(status: CompatibilityStatus): WasmCompatibilityStatus {
  switch (status) {
    case "implemented-js-and-wasm-shadow":
      return "shadow";
    case "implemented-wasm-production":
      return "production";
    case "blocked":
      return "blocked";
    default:
      return "not-started";
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
  } = {}
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
    wasmStatus: options.wasmStatus ?? deriveWasmStatus(status)
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
    notes: "Postfix percent is in the canonical formula corpus and tracked as part of the arithmetic lane."
  }),
  entry("comparison:equality-number", "comparison", "=A1=A2", "implemented-wasm-production"),
  entry("comparison:equality-text", "comparison", "=\"hello\"=\"HELLO\"", "implemented-js", {
    prerequisites: ["core:value-model", "core:comparison-model", "core:string-coercion"]
  }),
  entry("comparison:greater-than", "comparison", "=A1>A2", "implemented-wasm-production"),
  entry("comparison:less-than-or-equal", "comparison", "=A1<=A2", "implemented-wasm-production"),
  entry("logical:if-basic", "logical", "=IF(A1>0,\"yes\",\"no\")", "implemented-js", {
    notes: "Control-flow semantics stay JS-authoritative until branch laziness parity is proven in differential tests."
  }),
  entry("logical:and-basic", "logical", "=AND(A1,A2)", "implemented-wasm-production"),
  entry("logical:or-basic", "logical", "=OR(A1,A2)", "implemented-wasm-production"),
  entry("logical:not-basic", "logical", "=NOT(A1)", "implemented-wasm-production"),
  entry("aggregation:sum-range", "aggregation", "=SUM(A1:A3)", "implemented-wasm-production"),
  entry("aggregation:avg-range", "aggregation", "=AVG(A1:A3)", "implemented-wasm-production"),
  entry("aggregation:min-range", "aggregation", "=MIN(A1:A3)", "implemented-wasm-production"),
  entry("aggregation:max-range", "aggregation", "=MAX(A1:A3)", "implemented-wasm-production"),
  entry("aggregation:count-range", "aggregation", "=COUNT(A1:A4)", "implemented-wasm-production"),
  entry("aggregation:counta-range", "aggregation", "=COUNTA(A1:A4)", "implemented-wasm-production"),
  entry("math:abs-basic", "math", "=ABS(A1)", "implemented-wasm-production"),
  entry("math:round-basic", "math", "=ROUND(A1,1)", "implemented-wasm-production"),
  entry("math:floor-basic", "math", "=FLOOR(A1,2)", "implemented-wasm-production"),
  entry("math:ceiling-basic", "math", "=CEILING(A1,2)", "implemented-wasm-production"),
  entry("math:mod-basic", "math", "=MOD(A1,A2)", "implemented-wasm-production"),
  entry("text:concat-operator", "text", "=\"bi\"&\"lig\"", "implemented-js"),
  entry("text:concat-function", "text", "=CONCAT(\"bi\",\"lig\")", "implemented-js"),
  entry("text:len-basic", "text", "=LEN(\"bilig\")", "implemented-wasm-production", {
    notes: "LEN now executes through the string-aware native runtime for scalar inputs."
  }),
  entry("date-time:serial-addition", "date-time", "=A1+7", "implemented-wasm-production"),
  entry("date-time:date-constructor", "date-time", "=DATE(2026,3,15)", "implemented-wasm-production"),
  entry("date-time:today-volatile", "date-time", "=TODAY()", "implemented-js", {
    prerequisites: ["core:value-model", "core:date-serial-model", "core:volatile-context"]
  }),
  entry("lookup-reference:index-basic", "lookup-reference", "=INDEX(A1:B2,2,1)", "implemented-js"),
  entry("lookup-reference:match-exact", "lookup-reference", "=MATCH(\"pear\",A1:A3,0)", "implemented-js"),
  entry("lookup-reference:vlookup-exact", "lookup-reference", "=VLOOKUP(\"pear\",A1:B3,2,FALSE)", "implemented-js"),
  entry("lookup-reference:xlookup-exact", "lookup-reference", "=XLOOKUP(\"pear\",A1:A3,B1:B3)", "implemented-js"),
  entry("statistical:averageif-basic", "statistical", "=AVERAGEIF(A1:A4,\">0\")", "implemented-js"),
  entry("statistical:countif-basic", "statistical", "=COUNTIF(A1:A4,\">0\")", "implemented-js"),
  entry("information:isblank-basic", "information", "=ISBLANK(A1)", "implemented-wasm-production"),
  entry("information:isnumber-basic", "information", "=ISNUMBER(A1)", "implemented-wasm-production"),
  entry("information:istext-basic", "information", "=ISTEXT(A1)", "implemented-wasm-production"),
  entry("dynamic-array:sequence-spill", "dynamic-array", "=SEQUENCE(3,1,1,1)", "blocked", {
    notes: "Spill allocation and shape semantics are required before this family can move past blocked."
  }),
  entry("dynamic-array:filter-basic", "dynamic-array", "=FILTER(A1:A4,A1:A4>2)", "blocked"),
  entry("dynamic-array:unique-basic", "dynamic-array", "=UNIQUE(A1:A4)", "blocked"),
  entry("names:defined-name-scalar", "names", "=TaxRate*A1", "blocked"),
  entry("tables:table-total-row-sum", "tables", "=SUM(Sales[Amount])", "blocked"),
  entry("structured-reference:table-column-ref", "structured-reference", "=SUM(Sales[Amount])", "blocked"),
  entry("volatile:rand-basic", "volatile", "=RAND()", "implemented-js", {
    prerequisites: ["core:volatile-context", "core:value-model"]
  }),
  entry("lambda:let-basic", "lambda", "=LET(x,2,x+3)", "blocked"),
  entry("lambda:lambda-invoke", "lambda", "=LAMBDA(x,x+1)(4)", "blocked"),
  entry("lambda:map-basic", "lambda", "=MAP(A1:A3,LAMBDA(x,x*2))", "blocked"),
  entry("logical:if-true-branch", "logical", "=IF(A1,A2,A3)", "implemented-js"),
  entry("logical:if-condition-error", "logical", "=IF(1/0,1,2)", "implemented-js", {
    notes: "The JS evaluator now propagates IF condition errors before either branch executes."
  }),
  entry("logical:iferror-catches-any-error", "logical", "=IFERROR(1/0,\"fallback\")", "implemented-js"),
  entry("logical:ifna-catches-na-only", "logical", "=IFNA(NA(),\"missing\")", "implemented-js", {
    notes: "NA() now normalizes to #N/A on the JS oracle path, allowing IFNA parity to execute."
  }),
  entry("logical:and-false-on-empty", "logical", "=AND(TRUE,A1)", "implemented-js"),
  entry("logical:or-true-branch", "logical", "=OR(A1,TRUE)", "implemented-js"),
  entry("logical:not-number", "logical", "=NOT(2)", "implemented-js"),
  entry("information:isblank-empty", "information", "=ISBLANK(A1)", "implemented-wasm-production"),
  entry("information:isnumber-number", "information", "=ISNUMBER(42)", "implemented-wasm-production"),
  entry("information:istext-string", "information", "=ISTEXT(\"hello\")", "implemented-wasm-production"),
  entry("text:len-counts-plain-string-length", "text", "=LEN(\"hello\")", "implemented-js"),
  entry("text:exact-basic", "text", "=EXACT(\"Alpha\",\"alpha\")", "implemented-wasm-production", {
    notes: "EXACT now routes through the string-aware WASM runtime."
  }),
  entry("text:left-basic", "text", "=LEFT(\"alpha\",3)", "implemented-js"),
  entry("text:right-basic", "text", "=RIGHT(\"alpha\",2)", "implemented-js"),
  entry("text:mid-basic", "text", "=MID(\"alphabet\",2,3)", "implemented-js"),
  entry("text:trim-basic", "text", "=TRIM(\"  alpha   beta  \")", "implemented-js"),
  entry("text:upper-basic", "text", "=UPPER(\"Alpha beta\")", "implemented-js"),
  entry("text:lower-basic", "text", "=LOWER(\"Alpha BETA\")", "implemented-js"),
  entry("text:find-basic", "text", "=FIND(\"ph\",\"alphabet\")", "implemented-js"),
  entry("text:search-basic", "text", "=SEARCH(\"PH\",\"alphabet\")", "implemented-js"),
  entry("text:value-basic", "text", "=VALUE(\"42\")", "implemented-wasm-production", {
    notes: "The canonical literal case constant-folds to a native-safe number; dynamic text inputs still execute through JS until string coercion is promoted."
  }),
  entry("lookup-reference:xmatch-basic", "lookup-reference", "=XMATCH(\"pear\",A1:A3,0)", "implemented-js"),
  entry("lookup-reference:hlookup-basic", "lookup-reference", "=HLOOKUP(\"pear\",A1:C2,2,FALSE)", "implemented-js"),
  entry("lookup-reference:offset-basic", "lookup-reference", "=OFFSET(A1,1,1)", "blocked", {
    notes: "Reference-returning OFFSET depends on richer reference model and rebinding semantics."
  }),
  entry("dynamic-array:take-basic", "dynamic-array", "=TAKE(A1:A4,2)", "blocked"),
  entry("dynamic-array:drop-basic", "dynamic-array", "=DROP(A1:A4,2)", "blocked"),
  entry("dynamic-array:choosecols-basic", "dynamic-array", "=CHOOSECOLS(A1:C2,1,3)", "blocked"),
  entry("dynamic-array:chooserows-basic", "dynamic-array", "=CHOOSEROWS(A1:B3,1,3)", "blocked"),
  entry("statistical:sumif-basic", "statistical", "=SUMIF(A1:A4,\">0\",B1:B4)", "implemented-js"),
  entry("statistical:sumifs-basic", "statistical", "=SUMIFS(C1:C4,A1:A4,\">0\",B1:B4,\"x\")", "implemented-js"),
  entry("statistical:averageifs-basic", "statistical", "=AVERAGEIFS(C1:C4,A1:A4,\">0\",B1:B4,\"x\")", "implemented-js"),
  entry("statistical:countifs-basic", "statistical", "=COUNTIFS(A1:A4,\">0\",B1:B4,\"x\")", "implemented-js"),
  entry("math:sumproduct-basic", "math", "=SUMPRODUCT(A1:A3,B1:B3)", "implemented-js"),
  entry("math:int-basic", "math", "=INT(-3.1)", "implemented-wasm-production"),
  entry("math:roundup-basic", "math", "=ROUNDUP(12.341,2)", "implemented-wasm-production"),
  entry("math:rounddown-basic", "math", "=ROUNDDOWN(12.349,2)", "implemented-wasm-production"),
  entry("date-time:now-volatile", "date-time", "=NOW()", "implemented-js", {
    prerequisites: ["core:value-model", "core:date-serial-model", "core:volatile-context"],
    notes: "NOW executes on the JS oracle with deterministic capture tests; native promotion still depends on a recalc-epoch runtime contract."
  }),
  entry("date-time:time-basic", "date-time", "=TIME(12,30,0)", "implemented-wasm-production"),
  entry("date-time:hour-basic", "date-time", "=HOUR(A1)", "implemented-wasm-production"),
  entry("date-time:minute-basic", "date-time", "=MINUTE(A1)", "implemented-wasm-production"),
  entry("date-time:second-basic", "date-time", "=SECOND(A1)", "implemented-wasm-production"),
  entry("date-time:weekday-basic", "date-time", "=WEEKDAY(DATE(2026,3,15))", "implemented-wasm-production"),
  entry("dynamic-array:sort-basic", "dynamic-array", "=SORT(A1:A4)", "blocked"),
  entry("dynamic-array:sortby-basic", "dynamic-array", "=SORTBY(A1:A3,B1:B3)", "blocked"),
  entry("dynamic-array:tocol-basic", "dynamic-array", "=TOCOL(A1:B2)", "blocked"),
  entry("dynamic-array:torow-basic", "dynamic-array", "=TOROW(A1:B2)", "blocked"),
  entry("dynamic-array:wraprows-basic", "dynamic-array", "=WRAPROWS(A1:A4,2)", "blocked"),
  entry("dynamic-array:wrapcols-basic", "dynamic-array", "=WRAPCOLS(A1:A4,2)", "blocked"),
  entry("names:defined-name-range", "names", "=SUM(MyRange)", "blocked"),
  entry("lambda:byrow-basic", "lambda", "=BYROW(A1:B2,LAMBDA(r,SUM(r)))", "blocked")
];

export function getCompatibilityEntry(id: string): FormulaCompatibilityEntry | undefined {
  return formulaCompatibilityRegistry.find((entry) => entry.id === id);
}

export function isCompatibilityStatus(value: string): value is CompatibilityStatus {
  return (compatibilityStatuses as readonly string[]).includes(value);
}

export function isWasmCompatibilityStatus(value: string): value is WasmCompatibilityStatus {
  return (wasmCompatibilityStatuses as readonly string[]).includes(value);
}
