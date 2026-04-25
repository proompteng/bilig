import type { CellValue, LiteralInput, RecalcMetrics } from '@bilig/protocol'
import type { EvaluationResult } from '@bilig/formula'
import type { EngineCounters } from '@bilig/core'

export type RawCellContent = LiteralInput | string

export type WorkPaperContextValue = string | number | boolean | null | WorkPaperContextObject | readonly WorkPaperContextValue[]

export interface WorkPaperContextObject {
  [key: string]: WorkPaperContextValue
}

export interface WorkPaperChooseAddressMappingPolicy {
  mode: 'dense' | 'sparse'
}

export type WorkPaperSheet = readonly (readonly RawCellContent[])[]
export type WorkPaperSheets = Record<string, WorkPaperSheet>

export interface WorkPaperCellAddress {
  sheet: number
  col: number
  row: number
}

export interface WorkPaperCellRange {
  start: WorkPaperCellAddress
  end: WorkPaperCellAddress
}

export interface WorkPaperAddressFormatOptions {
  includeSheetName?: boolean
}

export type WorkPaperAddressLike = WorkPaperCellAddress | WorkPaperCellRange
export type WorkPaperAxisInterval = readonly [start: number, count?: number]
export type WorkPaperAxisSwapMapping = readonly [from: number, to: number]

export interface WorkPaperSheetDimensions {
  width: number
  height: number
}

export type WorkPaperChange = WorkPaperCellChange | WorkPaperNamedExpressionChange

export interface WorkPaperCellChange {
  kind: 'cell'
  address: WorkPaperCellAddress
  sheetName: string
  a1: string
  newValue: CellValue
}

export interface WorkPaperNamedExpressionChange {
  kind: 'named-expression'
  name: string
  scope?: number
  newValue: CellValue | CellValue[][]
}

export interface WorkPaperNamedExpression {
  name: string
  expression: RawCellContent
  scope?: number
  options?: Record<string, string | number | boolean>
}

export interface SerializedWorkPaperNamedExpression extends WorkPaperNamedExpression {}

export interface WorkPaperSimpleDate {
  year: number
  month: number
  day: number
}

export interface WorkPaperSimpleTime {
  hours: number
  minutes: number
  seconds: number
}

export interface WorkPaperDateTime extends WorkPaperSimpleDate, WorkPaperSimpleTime {}

export type WorkPaperParsedDateTime = WorkPaperSimpleDate | WorkPaperSimpleTime | WorkPaperDateTime

export type WorkPaperParseDateTime = (
  dateTimeString: string,
  dateFormat?: string,
  timeFormat?: string,
) => WorkPaperParsedDateTime | undefined

export type WorkPaperStringifyDateTime = (dateTime: WorkPaperDateTime, dateTimeFormat: string) => string | undefined

export type WorkPaperStringifyDuration = (time: WorkPaperSimpleTime, timeFormat: string) => string | undefined

export type WorkPaperFunctionArgumentType = 'STRING' | 'NUMBER' | 'BOOLEAN' | 'SCALAR' | 'NOERROR' | 'RANGE' | 'INTEGER' | 'COMPLEX' | 'ANY'

export interface WorkPaperFunctionArgument {
  argumentType: WorkPaperFunctionArgumentType
  passSubtype?: boolean
  defaultValue?: unknown
  optionalArg?: boolean
  minValue?: number
  maxValue?: number
  lessThan?: number
  greaterThan?: number
}

export interface WorkPaperFunctionMetadata {
  method: string
  parameters?: WorkPaperFunctionArgument[]
  repeatLastArgs?: number
  expandRanges?: boolean
  returnNumberType?: string
  sizeOfResultArrayMethod?: string
  isVolatile?: boolean
  isDependentOnSheetStructureChange?: boolean
  doesNotNeedArgumentsToBeComputed?: boolean
  enableArrayArithmeticForArguments?: boolean
  vectorizationForbidden?: boolean
}

export interface WorkPaperFunctionPlugin {
  implementedFunctions: Record<string, WorkPaperFunctionMetadata>
  aliases?: Record<string, string>
}

export interface WorkPaperFunctionPluginDefinition extends WorkPaperFunctionPlugin {
  id: string
  functions?: Record<string, (...args: CellValue[]) => EvaluationResult | CellValue>
}

export type WorkPaperFunctionTranslationsPackage = Record<string, Record<string, string>>

export interface WorkPaperLanguagePackage {
  readonly functions?: Record<string, string>
  readonly errors?: Record<string, string>
  readonly ui?: Record<string, string>
  readonly [key: string]: unknown
}

export type WorkPaperLicenseKeyValidityState = 'valid' | 'invalid' | 'expired' | 'missing'

export interface WorkPaperConfig {
  accentSensitive?: boolean
  caseSensitive?: boolean
  caseFirst?: 'upper' | 'lower' | 'false'
  chooseAddressMappingPolicy?: WorkPaperChooseAddressMappingPolicy
  context?: WorkPaperContextValue
  currencySymbol?: string[]
  dateFormats?: string[]
  functionArgSeparator?: string
  decimalSeparator?: '.' | ','
  evaluateNullToZero?: boolean
  functionPlugins?: WorkPaperFunctionPluginDefinition[]
  ignorePunctuation?: boolean
  language?: string
  ignoreWhiteSpace?: 'standard' | 'any'
  leapYear1900?: boolean
  licenseKey?: string
  localeLang?: string
  matchWholeCell?: boolean
  arrayColumnSeparator?: ',' | ';'
  arrayRowSeparator?: ';' | '|'
  maxRows?: number
  maxColumns?: number
  nullDate?: { year: number; month: number; day: number }
  nullYear?: number
  parseDateTime?: WorkPaperParseDateTime
  precisionEpsilon?: number
  precisionRounding?: number
  stringifyDateTime?: WorkPaperStringifyDateTime
  stringifyDuration?: WorkPaperStringifyDuration
  smartRounding?: boolean
  thousandSeparator?: '' | ',' | '.'
  timeFormats?: string[]
  useArrayArithmetic?: boolean
  useColumnIndex?: boolean
  useStats?: boolean
  undoLimit?: number
  useRegularExpressions?: boolean
  useWildcards?: boolean
}

export interface WorkPaperDetailedEventMap {
  sheetAdded: { sheetId: number; sheetName: string }
  sheetRemoved: { sheetId: number; sheetName: string; changes: WorkPaperChange[] }
  sheetRenamed: { sheetId: number; oldName: string; newName: string }
  namedExpressionAdded: { name: string; scope?: number; changes: WorkPaperChange[] }
  namedExpressionRemoved: { name: string; scope?: number; changes: WorkPaperChange[] }
  valuesUpdated: { changes: WorkPaperChange[] }
  evaluationSuspended: {}
  evaluationResumed: { changes: WorkPaperChange[] }
}

export interface WorkPaperEventMap {
  sheetAdded: [sheetName: string]
  sheetRemoved: [sheetName: string, changes: WorkPaperChange[]]
  sheetRenamed: [oldName: string, newName: string]
  namedExpressionAdded: [name: string, changes: WorkPaperChange[]]
  namedExpressionRemoved: [name: string, changes: WorkPaperChange[]]
  valuesUpdated: [changes: WorkPaperChange[]]
  evaluationSuspended: []
  evaluationResumed: [changes: WorkPaperChange[]]
}

export type WorkPaperEventName = keyof WorkPaperEventMap

export type WorkPaperListener<EventName extends WorkPaperEventName> = (...args: WorkPaperEventMap[EventName]) => void

export type WorkPaperDetailedListener<EventName extends WorkPaperEventName> = (payload: WorkPaperDetailedEventMap[EventName]) => void

export type WorkPaperCellType = 'EMPTY' | 'VALUE' | 'FORMULA' | 'ARRAY'
export type WorkPaperCellValueType = 'EMPTY' | 'NUMBER' | 'STRING' | 'BOOLEAN' | 'ERROR'
export type WorkPaperCellValueDetailedType = WorkPaperCellValueType | 'DATE' | 'TIME' | 'DATETIME'

export type WorkPaperDependencyRef =
  | { kind: 'cell'; address: WorkPaperCellAddress }
  | { kind: 'range'; range: WorkPaperCellRange }
  | { kind: 'name'; name: string }

export interface WorkPaperStats {
  batchDepth: number
  evaluationSuspended: boolean
  lastMetrics: RecalcMetrics
}

export type WorkPaperEngineCounters = EngineCounters

export interface WorkPaperGraphAdapter {
  getDependents(reference: WorkPaperAddressLike): WorkPaperDependencyRef[]
  getPrecedents(reference: WorkPaperAddressLike): WorkPaperDependencyRef[]
}

export interface WorkPaperRangeMappingAdapter {
  getValues(range: WorkPaperCellRange): CellValue[][]
  getSerialized(range: WorkPaperCellRange): RawCellContent[][]
}

export interface WorkPaperArrayMappingAdapter {
  isPartOfArray(address: WorkPaperCellAddress): boolean
  getFormula(address: WorkPaperCellAddress): string | undefined
}

export interface WorkPaperSheetMappingAdapter {
  getSheetName(sheetId: number): string | undefined
  getSheetId(name: string): number | undefined
  getSheetNames(): string[]
  countSheets(): number
}

export interface WorkPaperAddressMappingAdapter {
  has(address: WorkPaperCellAddress): boolean
  getValue(address: WorkPaperCellAddress): CellValue
  getFormula(address: WorkPaperCellAddress): string | undefined
}

export interface WorkPaperDependencyGraphAdapter {
  getCellDependents(reference: WorkPaperAddressLike): WorkPaperDependencyRef[]
  getCellPrecedents(reference: WorkPaperAddressLike): WorkPaperDependencyRef[]
}

export interface WorkPaperEvaluatorAdapter {
  recalculate(): WorkPaperChange[]
  calculateFormula(formula: string, scope?: number): CellValue | CellValue[][]
}

export interface WorkPaperColumnSearchAdapter {
  find(sheetId: number, column: number, matcher: string | ((value: CellValue) => boolean)): WorkPaperCellAddress[]
}

export interface WorkPaperLazilyTransformingAstServiceAdapter {
  normalizeFormula(formula: string): string
  validateFormula(formula: string): boolean
  getNamedExpressionsFromFormula(formula: string): string[]
}

export interface WorkPaperInternals {
  graph: WorkPaperGraphAdapter
  rangeMapping: WorkPaperRangeMappingAdapter
  arrayMapping: WorkPaperArrayMappingAdapter
  sheetMapping: WorkPaperSheetMappingAdapter
  addressMapping: WorkPaperAddressMappingAdapter
  dependencyGraph: WorkPaperDependencyGraphAdapter
  evaluator: WorkPaperEvaluatorAdapter
  columnSearch: WorkPaperColumnSearchAdapter
  lazilyTransformingAstService: WorkPaperLazilyTransformingAstServiceAdapter
}
