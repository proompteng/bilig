import type { CellValue, ErrorCode, LiteralInput } from '@bilig/protocol'
import { excelDateTimeFixtureSuite } from './datetime-fixtures.js'
import { canonicalLogicalFixtures } from './logical-fixtures.js'
import { canonicalExpansionFixtures } from './canonical-expansion-fixtures.js'
import { canonicalFoundationFixtures } from './canonical-foundation-fixtures.js'
import { canonicalStatisticalFixtures } from './statistical-fixtures.js'
import { canonicalTextFixtures } from './text-fixtures.js'
import { canonicalWorkbookSemanticsFixtures } from './workbook-semantics-fixtures.js'

export const excelFixtureFamilies = [
  'arithmetic',
  'comparison',
  'logical',
  'aggregation',
  'math',
  'text',
  'date-time',
  'lookup-reference',
  'statistical',
  'information',
  'dynamic-array',
  'names',
  'tables',
  'structured-reference',
  'volatile',
  'lambda',
] as const

export type ExcelFixtureFamily = (typeof excelFixtureFamilies)[number]

export const excelFixtureIdPattern = /^[a-z][a-z0-9-]*:[a-z0-9-]+$/

export type ExcelExpectedValue =
  | { kind: 'empty' }
  | { kind: 'number'; value: number }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'string'; value: string }
  | { kind: 'error'; code: ErrorCode; display: string }

export interface ExcelFixtureInputCell {
  address: string
  sheetName?: string
  input: LiteralInput
  note?: string
}

export interface ExcelFixtureDefinedName {
  name: string
  value: LiteralInput
  note?: string
}

export interface ExcelFixtureTable {
  name: string
  sheetName?: string
  startAddress: string
  endAddress: string
  columnNames: string[]
  headerRow: boolean
  totalsRow: boolean
  note?: string
}

export interface ExcelFixtureExpectedOutput {
  address: string
  sheetName?: string
  expected: ExcelExpectedValue
  note?: string
}

export interface ExcelFixtureMultipleOperationsMock {
  formulaAddress: string
  formulaSheetName?: string
  rowCellAddress: string
  rowCellSheetName?: string
  rowReplacementAddress: string
  rowReplacementSheetName?: string
  columnCellAddress?: string
  columnCellSheetName?: string
  columnReplacementAddress?: string
  columnReplacementSheetName?: string
  result: ExcelExpectedValue
}

export interface ExcelFixtureCell {
  address: string
  formula?: string
  input?: LiteralInput
  expected: ExcelExpectedValue | CellValue
}

export interface ExcelFixtureSheet {
  name: string
  cells?: ExcelFixtureCell[]
}

export interface ExcelFixtureCase {
  id: string
  family: ExcelFixtureFamily
  title: string
  formula: string
  notes?: string
  sheetName?: string
  definedNames?: ExcelFixtureDefinedName[]
  tables?: ExcelFixtureTable[]
  multipleOperations?: ExcelFixtureMultipleOperationsMock
  inputs: ExcelFixtureInputCell[]
  outputs: ExcelFixtureExpectedOutput[]
}

export interface ExcelFixtureSuite {
  id: string
  description: string
  sheets: ExcelFixtureSheet[]
  cases?: readonly ExcelFixtureCase[]
  excelBuild: string
  capturedAt: string
}

export function createExcelFixtureId(family: ExcelFixtureFamily, slug: string): string {
  const normalizedSlug = slug.trim().toLowerCase()
  if (!/^[a-z0-9-]+$/.test(normalizedSlug)) {
    throw new Error(`Invalid Excel fixture slug: ${slug}`)
  }
  const id = `${family}:${normalizedSlug}`
  if (!excelFixtureIdPattern.test(id)) {
    throw new Error(`Invalid Excel fixture id: ${id}`)
  }
  return id
}

export function isExcelFixtureId(value: string): boolean {
  return excelFixtureIdPattern.test(value)
}

export function emptyExpected(): ExcelExpectedValue {
  return { kind: 'empty' }
}

export function numberExpected(value: number): ExcelExpectedValue {
  return { kind: 'number', value }
}

export function booleanExpected(value: boolean): ExcelExpectedValue {
  return { kind: 'boolean', value }
}

export function stringExpected(value: string): ExcelExpectedValue {
  return { kind: 'string', value }
}

export function errorExpected(code: ErrorCode, display: string): ExcelExpectedValue {
  return { kind: 'error', code, display }
}

function dedupeFixtures(fixtures: readonly ExcelFixtureCase[]): ExcelFixtureCase[] {
  const seen = new Set<string>()
  const output: ExcelFixtureCase[] = []
  for (const fixture of fixtures) {
    if (seen.has(fixture.id)) {
      continue
    }
    seen.add(fixture.id)
    output.push(fixture)
  }
  return output
}

const canonicalCorpusExclusions = new Set<string>(['text:case-insensitive-compare', 'information:value-error-display'])

const canonicalBaseFixtureIds = new Set<string>([
  'arithmetic:add-basic',
  'arithmetic:precedence-basic',
  'arithmetic:unary-negation',
  'arithmetic:division-basic',
  'arithmetic:power-basic',
  'arithmetic:percent-operator',
  'comparison:equality-number',
  'comparison:equality-text',
  'comparison:greater-than',
  'comparison:less-than-or-equal',
  'logical:if-basic',
  'logical:ifs-basic',
  'logical:and-basic',
  'logical:or-basic',
  'logical:not-basic',
  'logical:switch-basic',
  'logical:xor-basic',
  'aggregation:sum-range',
  'aggregation:avg-range',
  'aggregation:min-range',
  'aggregation:max-range',
  'aggregation:count-range',
  'aggregation:counta-range',
  'aggregation:countblank-range',
  'math:abs-basic',
  'math:round-basic',
  'math:trunc-basic',
  'math:floor-math-basic',
  'math:floor-precise-basic',
  'math:ceiling-math-basic',
  'math:ceiling-precise-basic',
  'math:iso-ceiling-basic',
  'math:floor-basic',
  'math:ceiling-basic',
  'math:mod-basic',
  'math:bitand-basic',
  'math:base-basic',
  'math:decimal-basic',
  'math:bin2dec-basic',
  'math:dec2bin-basic',
  'math:oct2hex-basic',
  'math:besseli-basic',
  'math:besselj-basic',
  'math:besselk-basic',
  'math:bessely-basic',
  'math:convert-basic',
  'math:euroconvert-basic',
  'math:acosh-basic',
  'math:fact-basic',
  'math:combin-basic',
  'math:permut-basic',
  'math:permutationa-basic',
  'math:mround-basic',
  'math:seriessum-basic',
  'math:gcd-basic',
  'math:product-basic',
  'math:geomean-basic',
  'math:harmean-basic',
  'math:sqrtpi-basic',
  'math:sumsq-basic',
  'text:concat-operator',
  'text:concat-function',
  'text:len-basic',
  'text:textbefore-basic',
  'text:textafter-basic',
  'text:textjoin-basic',
  'text:textsplit-basic',
  'text:asc-basic',
  'text:jis-basic',
  'text:dbcs-basic',
  'text:char-basic',
  'text:code-basic',
  'text:clean-basic',
  'text:unichar-basic',
  'text:dollar-basic',
  'text:text-basic',
  'text:text-date-basic',
  'text:phonetic-basic',
  'text:bahttext-basic',
  'information:t-basic',
  'information:n-basic',
  'information:type-basic',
  'math:delta-basic',
  'math:gestep-basic',
  'statistical:gauss-basic',
  'statistical:phi-basic',
  'statistical:standardize-basic',
  'statistical:confidence-norm-basic',
  'statistical:mode-basic',
  'statistical:mode-sngl-basic',
  'statistical:stdev-basic',
  'statistical:stdeva-basic',
  'statistical:var-basic',
  'statistical:vara-basic',
  'statistical:skew-basic',
  'statistical:kurt-basic',
  'statistical:normdist-basic',
  'statistical:norminv-basic',
  'statistical:normsdist-basic',
  'statistical:normsinv-basic',
  'statistical:loginv-basic',
  'statistical:lognormdist-basic',
  'date-time:serial-addition',
  'date-time:date-constructor',
  'date-time:datedif-ym',
  'date-time:days360-basic',
  'date-time:isoweeknum-basic',
  'date-time:workday-intl-basic',
  'date-time:networkdays-intl-basic',
  'date-time:timevalue-basic',
  'date-time:yearfrac-basic',
  'date-time:fvschedule-basic',
  'date-time:effect-basic',
  'date-time:nominal-basic',
  'date-time:pduration-basic',
  'date-time:rri-basic',
  'date-time:fv-basic',
  'date-time:pv-basic',
  'date-time:pmt-basic',
  'date-time:nper-basic',
  'date-time:npv-basic',
  'date-time:rate-basic',
  'date-time:irr-basic',
  'date-time:mirr-basic',
  'date-time:xnpv-basic',
  'date-time:xirr-basic',
  'date-time:ipmt-basic',
  'date-time:ppmt-basic',
  'date-time:ispmt-basic',
  'date-time:cumipmt-basic',
  'date-time:cumprinc-basic',
  'date-time:db-basic',
  'date-time:ddb-basic',
  'date-time:vdb-basic',
  'date-time:sln-basic',
  'date-time:syd-basic',
  'date-time:disc-basic',
  'date-time:intrate-basic',
  'date-time:received-basic',
  'date-time:pricedisc-basic',
  'date-time:yielddisc-basic',
  'date-time:pricemat-basic',
  'date-time:yieldmat-basic',
  'date-time:oddfprice-basic',
  'date-time:oddfyield-basic',
  'date-time:oddlprice-basic',
  'date-time:oddlyield-basic',
  'date-time:coupdaybs-basic',
  'date-time:coupdays-basic',
  'date-time:coupdaysnc-basic',
  'date-time:coupncd-basic',
  'date-time:coupnum-basic',
  'date-time:couppcd-basic',
  'date-time:price-basic',
  'date-time:yield-basic',
  'date-time:duration-basic',
  'date-time:mduration-basic',
  'date-time:tbillprice-basic',
  'date-time:tbillyield-basic',
  'date-time:tbilleq-basic',
  'date-time:today-volatile',
  'lookup-reference:index-basic',
  'lookup-reference:choose-basic',
  'lookup-reference:address-basic',
  'lookup-reference:match-exact',
  'lookup-reference:vlookup-exact',
  'lookup-reference:xlookup-exact',
  'statistical:averageif-basic',
  'statistical:countif-basic',
  'statistical:chisqdist-basic',
  'statistical:chiinv-basic',
  'statistical:chisq-inv-rt-basic',
  'statistical:chisqinv-basic',
  'statistical:chisq-inv-basic',
  'statistical:chisq-test-basic',
  'statistical:beta-dist-basic',
  'statistical:beta-inv-basic',
  'statistical:f-dist-rt-basic',
  'statistical:fdist-basic',
  'statistical:f-inv-basic',
  'statistical:f-inv-rt-basic',
  'statistical:f-test-basic',
  'statistical:z-test-basic',
  'statistical:correl-basic',
  'statistical:covar-basic',
  'statistical:covariance-p-basic',
  'statistical:covariance-s-basic',
  'statistical:pearson-basic',
  'statistical:intercept-basic',
  'statistical:slope-basic',
  'statistical:rsq-basic',
  'statistical:steyx-basic',
  'statistical:rank-basic',
  'statistical:rank-eq-basic',
  'statistical:rank-avg-basic',
  'statistical:median-basic',
  'statistical:small-basic',
  'statistical:large-basic',
  'statistical:percentile-basic',
  'statistical:percentile-inc-basic',
  'statistical:percentile-exc-basic',
  'statistical:percentrank-basic',
  'statistical:percentrank-inc-basic',
  'statistical:percentrank-exc-basic',
  'statistical:quartile-basic',
  'statistical:quartile-inc-basic',
  'statistical:quartile-exc-basic',
  'statistical:mode-mult-basic',
  'statistical:frequency-basic',
  'statistical:t-dist-basic',
  'statistical:t-inv-2t-basic',
  'statistical:confidence-t-basic',
  'statistical:gamma-inv-basic',
  'statistical:t-test-basic',
  'statistical:forecast-basic',
  'statistical:forecast-linear-basic',
  'statistical:trend-basic',
  'statistical:growth-basic',
  'statistical:linest-basic',
  'statistical:logest-basic',
  'statistical:prob-basic',
  'statistical:trimmean-basic',
  'statistical:daverage-basic',
  'statistical:dcount-basic',
  'statistical:dcounta-basic',
  'statistical:dget-basic',
  'statistical:dmax-basic',
  'statistical:dmin-basic',
  'statistical:dproduct-basic',
  'statistical:dstdev-basic',
  'statistical:dstdevp-basic',
  'statistical:dsum-basic',
  'statistical:dvar-basic',
  'statistical:dvarp-basic',
  'information:isblank-basic',
  'information:isnumber-basic',
  'information:istext-basic',
  'dynamic-array:sequence-spill',
  'dynamic-array:sequence-aggregate',
  'dynamic-array:filter-basic',
  'dynamic-array:unique-basic',
  'names:defined-name-scalar',
  'tables:table-total-row-sum',
  'structured-reference:table-column-ref',
  'volatile:rand-basic',
  'lambda:let-basic',
  'lambda:lambda-invoke',
  'lambda:map-basic',
  'logical:if-true-branch',
  'logical:if-condition-error',
  'logical:iferror-catches-any-error',
  'logical:ifna-catches-na-only',
  'logical:and-false-on-empty',
  'logical:or-true-branch',
  'logical:not-number',
  'information:isblank-empty',
  'information:isnumber-number',
  'information:istext-string',
  'text:len-counts-plain-string-length',
])

const canonicalBaseFixtures = dedupeFixtures([
  ...canonicalFoundationFixtures,
  ...canonicalLogicalFixtures,
  ...canonicalTextFixtures,
  ...canonicalStatisticalFixtures,
  ...(excelDateTimeFixtureSuite.cases ?? []),
]).filter((fixture) => canonicalBaseFixtureIds.has(fixture.id) && !canonicalCorpusExclusions.has(fixture.id))

export const canonicalFormulaFixtures: readonly ExcelFixtureCase[] = dedupeFixtures([
  ...canonicalBaseFixtures,
  ...canonicalExpansionFixtures,
])

export const workbookSemanticsFormulaFixtureSuite: ExcelFixtureSuite = {
  id: 'workbook-semantics',
  description: 'Extended workbook semantics fixture slice for names and cross-sheet reference behavior.',
  sheets: [{ name: 'Sheet1' }, { name: 'Sheet2' }],
  excelBuild: 'Microsoft 365 / 2026-03-19',
  capturedAt: '2026-03-19T00:00:00.000Z',
  cases: canonicalWorkbookSemanticsFixtures,
}

export const canonicalFormulaSmokeSuite: ExcelFixtureSuite = {
  id: 'canonical-smoke',
  description: 'Representative smoke slice from the canonical formula compatibility corpus.',
  sheets: [{ name: 'Sheet1' }],
  excelBuild: 'Microsoft 365 / 2026-03-19',
  capturedAt: '2026-03-19T00:00:00.000Z',
  cases: canonicalFormulaFixtures.slice(0, 5),
}

function buildFamilySuite(id: string, description: string, families: readonly ExcelFixtureFamily[]): ExcelFixtureSuite {
  return {
    id,
    description,
    sheets: [{ name: 'Sheet1' }],
    excelBuild: 'Microsoft 365 / 2026-03-19',
    capturedAt: '2026-03-19T00:00:00.000Z',
    cases: canonicalFormulaFixtures.filter((fixture) => families.includes(fixture.family)),
  }
}

export const textFormulaFixtureSuite = buildFamilySuite('canonical-text', 'Canonical formula corpus text-function fixture slice.', ['text'])

export const lookupReferenceFormulaFixtureSuite = buildFamilySuite(
  'canonical-lookup-reference',
  'Canonical formula corpus lookup/reference fixture slice.',
  ['lookup-reference', 'statistical'],
)

export const dateTimeFormulaFixtureSuite = buildFamilySuite(
  'canonical-date-time',
  'Canonical formula corpus date/time and volatile fixture slice.',
  ['date-time', 'volatile'],
)

export const dynamicArrayFormulaFixtureSuite = buildFamilySuite(
  'canonical-dynamic-array',
  'Canonical formula corpus dynamic-array fixture slice.',
  ['dynamic-array'],
)

export const namesTablesFormulaFixtureSuite = buildFamilySuite(
  'canonical-names-tables',
  'Canonical formula corpus names, tables, and structured-reference fixture slice.',
  ['names', 'tables', 'structured-reference'],
)

export const lambdaFormulaFixtureSuite = buildFamilySuite('canonical-lambda', 'Canonical formula corpus lambda fixture slice.', ['lambda'])

export { canonicalFoundationFixtures } from './canonical-foundation-fixtures.js'
export { canonicalExpansionFixtures } from './canonical-expansion-fixtures.js'
export { canonicalWorkbookSemanticsFixtures } from './workbook-semantics-fixtures.js'
export * from './logical-fixtures.js'
export * from './statistical-fixtures.js'
export * from './text-fixtures.js'
export * from './datetime-fixtures.js'
