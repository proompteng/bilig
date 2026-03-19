export const compatibilityStatuses = [
  "unsupported",
  "seeded",
  "implemented-js",
  "implemented-js-and-wasm"
] as const;

export type CompatibilityStatus = typeof compatibilityStatuses[number];

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
  notes?: string;
}

function entry(
  id: string,
  family: CompatibilityFamily,
  formula: string,
  status: CompatibilityStatus,
  notes?: string
): FormulaCompatibilityEntry {
  const base = { id, family, formula, status };
  return notes === undefined ? base : { ...base, notes };
}

export const top50CompatibilityRegistry: readonly FormulaCompatibilityEntry[] = [
  entry("arithmetic:add-basic", "arithmetic", "=A1+A2", "implemented-js-and-wasm"),
  entry("arithmetic:precedence-basic", "arithmetic", "=A1+A2*A3", "implemented-js-and-wasm"),
  entry("arithmetic:unary-negation", "arithmetic", "=-A1", "implemented-js-and-wasm"),
  entry("arithmetic:division-basic", "arithmetic", "=A1/A2", "implemented-js-and-wasm"),
  entry("arithmetic:power-basic", "arithmetic", "=A1^A2", "implemented-js-and-wasm"),
  entry("arithmetic:percent-operator", "arithmetic", "=A1*10%", "unsupported", "Postfix percent grammar is tracked but not implemented."),
  entry("comparison:equality-number", "comparison", "=A1=A2", "implemented-js-and-wasm"),
  entry("comparison:equality-text", "comparison", "=\"hello\"=\"HELLO\"", "implemented-js"),
  entry("comparison:greater-than", "comparison", "=A1>A2", "implemented-js-and-wasm"),
  entry("comparison:less-than-or-equal", "comparison", "=A1<=A2", "implemented-js-and-wasm"),
  entry("logical:if-basic", "logical", "=IF(A1>0,\"yes\",\"no\")", "implemented-js"),
  entry("logical:and-basic", "logical", "=AND(A1,A2)", "implemented-js"),
  entry("logical:or-basic", "logical", "=OR(A1,A2)", "implemented-js"),
  entry("logical:not-basic", "logical", "=NOT(A1)", "implemented-js"),
  entry("aggregation:sum-range", "aggregation", "=SUM(A1:A3)", "implemented-js-and-wasm"),
  entry("aggregation:avg-range", "aggregation", "=AVG(A1:A3)", "implemented-js-and-wasm"),
  entry("aggregation:min-range", "aggregation", "=MIN(A1:A3)", "implemented-js-and-wasm"),
  entry("aggregation:max-range", "aggregation", "=MAX(A1:A3)", "implemented-js-and-wasm"),
  entry("aggregation:count-range", "aggregation", "=COUNT(A1:A4)", "implemented-js-and-wasm"),
  entry("aggregation:counta-range", "aggregation", "=COUNTA(A1:A4)", "implemented-js-and-wasm"),
  entry("math:abs-basic", "math", "=ABS(A1)", "implemented-js-and-wasm"),
  entry("math:round-basic", "math", "=ROUND(A1,1)", "implemented-js"),
  entry("math:floor-basic", "math", "=FLOOR(A1,2)", "implemented-js"),
  entry("math:ceiling-basic", "math", "=CEILING(A1,2)", "implemented-js"),
  entry("math:mod-basic", "math", "=MOD(A1,A2)", "implemented-js-and-wasm"),
  entry("text:concat-operator", "text", "=\"bi\"&\"lig\"", "implemented-js"),
  entry("text:concat-function", "text", "=CONCAT(\"bi\",\"lig\")", "implemented-js"),
  entry("text:len-basic", "text", "=LEN(\"bilig\")", "implemented-js"),
  entry("text:case-insensitive-compare", "text", "=\"a\"=\"A\"", "implemented-js"),
  entry("date-time:serial-addition", "date-time", "=A1+7", "seeded"),
  entry("date-time:date-constructor", "date-time", "=DATE(2026,3,15)", "implemented-js"),
  entry("date-time:today-volatile", "date-time", "=TODAY()", "implemented-js", "Volatile recalc invalidation is still open even though JS evaluation works."),
  entry("lookup-reference:index-basic", "lookup-reference", "=INDEX(A1:B2,2,1)", "implemented-js"),
  entry("lookup-reference:match-exact", "lookup-reference", "=MATCH(\"pear\",A1:A3,0)", "implemented-js"),
  entry("lookup-reference:vlookup-exact", "lookup-reference", "=VLOOKUP(\"pear\",A1:B3,2,FALSE)", "implemented-js"),
  entry("lookup-reference:xlookup-exact", "lookup-reference", "=XLOOKUP(\"pear\",A1:A3,B1:B3)", "implemented-js"),
  entry("statistical:averageif-basic", "statistical", "=AVERAGEIF(A1:A4,\">0\")", "implemented-js"),
  entry("statistical:countif-basic", "statistical", "=COUNTIF(A1:A4,\">0\")", "implemented-js"),
  entry("information:isblank-basic", "information", "=ISBLANK(A1)", "implemented-js"),
  entry("dynamic-array:sequence-spill", "dynamic-array", "=SEQUENCE(3,1,1,1)", "unsupported"),
  entry("dynamic-array:filter-basic", "dynamic-array", "=FILTER(A1:A4,A1:A4>2)", "unsupported"),
  entry("dynamic-array:unique-basic", "dynamic-array", "=UNIQUE(A1:A4)", "unsupported"),
  entry("names:defined-name-scalar", "names", "=TaxRate*A1", "unsupported"),
  entry("tables:table-total-row-sum", "tables", "=SUM(Sales[Amount])", "unsupported"),
  entry("structured-reference:table-column-ref", "structured-reference", "=SUM(Sales[Amount])", "unsupported"),
  entry("volatile:rand-basic", "volatile", "=RAND()", "unsupported"),
  entry("lambda:let-basic", "lambda", "=LET(x,2,x+3)", "unsupported"),
  entry("lambda:lambda-invoke", "lambda", "=LAMBDA(x,x+1)(4)", "unsupported"),
  entry("lambda:map-basic", "lambda", "=MAP(A1:A3,LAMBDA(x,x*2))", "unsupported"),
  entry("information:value-error-display", "information", "=1+\"x\"", "implemented-js", "JS evaluator already surfaces #VALUE! for invalid coercions.")
];

export function getCompatibilityEntry(id: string): FormulaCompatibilityEntry | undefined {
  return top50CompatibilityRegistry.find((entry) => entry.id === id);
}

export function isCompatibilityStatus(value: string): value is CompatibilityStatus {
  return (compatibilityStatuses as readonly string[]).includes(value);
}
