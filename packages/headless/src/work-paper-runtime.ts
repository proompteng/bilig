import { SpreadsheetEngine, makeCellKey, type EngineCellMutationRef, type SheetRecord } from '@bilig/core'
import {
  ErrorCode,
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  type CellRangeRef,
  type CellSnapshot,
  type CellValue,
  type LiteralInput,
  type RecalcMetrics,
  type WorkbookDefinedNameValueSnapshot,
} from '@bilig/protocol'
import {
  excelSerialToDateParts,
  formatAddress,
  formatRangeAddress,
  installExternalFunctionAdapter,
  isArrayValue,
  isCellReferenceText,
  parseCellAddress,
  parseFormula,
  parseRangeAddress,
  serializeFormula,
  translateFormulaReferences,
  type EvaluationResult,
  type FormulaNode,
  type NameRefNode,
  type CallExprNode,
} from '@bilig/formula'
import { loadInitialMixedSheet, tryLoadInitialLiteralSheet } from './initial-sheet-load.js'
import { orderWorkPaperCellChanges } from './change-order.js'
import {
  WorkPaperConfigValueTooBigError,
  WorkPaperConfigValueTooSmallError,
  WorkPaperEvaluationSuspendedError,
  WorkPaperExpectedValueOfTypeError,
  WorkPaperOperationError,
  WorkPaperParseError,
  WorkPaperSheetError,
  WorkPaperExpectedOneOfValuesError,
  WorkPaperFunctionPluginValidationError,
  WorkPaperInvalidArgumentsError,
  WorkPaperLanguageAlreadyRegisteredError,
  WorkPaperLanguageNotRegisteredError,
  WorkPaperNamedExpressionDoesNotExistError,
  WorkPaperNamedExpressionNameIsAlreadyTakenError,
  WorkPaperNamedExpressionNameIsInvalidError,
  WorkPaperNoOperationToRedoError,
  WorkPaperNoOperationToUndoError,
  WorkPaperNoRelativeAddressesAllowedError,
  WorkPaperNoSheetWithIdError,
  WorkPaperNoSheetWithNameError,
  WorkPaperNotAFormulaError,
  WorkPaperNothingToPasteError,
  WorkPaperSheetNameAlreadyTakenError,
  WorkPaperSheetSizeLimitExceededError,
  WorkPaperUnableToParseError,
} from './work-paper-errors.js'
import { buildMatrixMutationPlan } from './matrix-mutation-plan.js'
import type {
  WorkPaperAddressMappingAdapter,
  WorkPaperAddressFormatOptions,
  WorkPaperAddressLike,
  WorkPaperArrayMappingAdapter,
  WorkPaperAxisInterval,
  WorkPaperAxisSwapMapping,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperCellChange,
  WorkPaperCellType,
  WorkPaperCellValueDetailedType,
  WorkPaperCellValueType,
  WorkPaperChange,
  WorkPaperColumnSearchAdapter,
  WorkPaperConfig,
  WorkPaperDateTime,
  WorkPaperDependencyGraphAdapter,
  WorkPaperDependencyRef,
  WorkPaperEvaluatorAdapter,
  WorkPaperFunctionPluginDefinition,
  WorkPaperFunctionTranslationsPackage,
  WorkPaperGraphAdapter,
  WorkPaperLanguagePackage,
  WorkPaperLazilyTransformingAstServiceAdapter,
  WorkPaperLicenseKeyValidityState,
  WorkPaperNamedExpression,
  WorkPaperRangeMappingAdapter,
  WorkPaperSheet,
  WorkPaperSheetDimensions,
  WorkPaperSheetMappingAdapter,
  WorkPaperSheets,
  WorkPaperStats,
  WorkPaperDetailedEventMap,
  WorkPaperDetailedListener,
  WorkPaperEventName,
  WorkPaperInternals,
  WorkPaperListener,
  RawCellContent,
  SerializedWorkPaperNamedExpression,
} from './work-paper-types.js'
import { captureTrackedEngineEvent, type TrackedEngineEvent } from './tracked-engine-event-refs.js'
import { calculateWorkPaperFormulaInScratchWorkbook } from './work-paper-scratch-evaluator.js'
import { replaceWorkPaperSheetContent } from './work-paper-sheet-replacement.js'

type ListenerMap = {
  [EventName in WorkPaperEventName]: Set<WorkPaperListener<EventName>>
}

type DetailedListenerMap = {
  [EventName in WorkPaperEventName]: Set<WorkPaperDetailedListener<EventName>>
}

type DetailedEvent = {
  [EventName in WorkPaperEventName]: {
    eventName: EventName
    payload: WorkPaperDetailedEventMap[EventName]
  }
}[WorkPaperEventName]

interface EngineTrackedEventSubscription {
  subscribeTracked(
    listener: (event: {
      kind: 'batch'
      invalidation: 'cells' | 'full'
      changedCellIndices: number[] | Uint32Array
      invalidatedRanges: CellRangeRef[]
      invalidatedRows: { sheetName: string; startIndex: number; endIndex: number }[]
      invalidatedColumns: { sheetName: string; startIndex: number; endIndex: number }[]
      metrics: RecalcMetrics
      explicitChangedCount?: number
    }) => void,
  ): () => void
}

interface InternalNamedExpression {
  publicName: string
  normalizedName: string
  internalName: string
  scope?: number
  expression: RawCellContent
  options?: Record<string, string | number | boolean>
}

interface InternalFunctionBinding {
  pluginId: string
  publicName: string
  internalName: string
  implementation?: (...args: CellValue[]) => EvaluationResult | CellValue
}

interface SheetStateSnapshot {
  sheetId: number
  sheetName: string
  order: number
  cells: Map<number, CellValue>
}

type VisibilitySnapshot = Map<number, SheetStateSnapshot>
type NamedExpressionValueSnapshot = Map<string, CellValue | CellValue[][]>
const EMPTY_NAMED_EXPRESSION_VALUES: NamedExpressionValueSnapshot = new Map()
const VISIBILITY_SHEET_STRIDE = MAX_ROWS * MAX_COLS

interface ClipboardPayload {
  sourceAnchor: WorkPaperCellAddress
  serialized: RawCellContent[][]
  values: CellValue[][]
}

type QueuedEvent = Extract<
  DetailedEvent,
  {
    eventName: 'sheetAdded' | 'sheetRemoved' | 'sheetRenamed' | 'namedExpressionAdded' | 'namedExpressionRemoved'
  }
>

type HistoryTransactionRecord =
  | { kind: 'ops'; ops: unknown[]; potentialNewCells?: number }
  | { kind: 'single-op'; op: unknown; potentialNewCells?: number }

interface HistoryRecord {
  forward: HistoryTransactionRecord
  inverse: HistoryTransactionRecord
}

const DEFAULT_CONFIG: Readonly<WorkPaperConfig> = Object.freeze({
  accentSensitive: false,
  caseSensitive: false,
  caseFirst: 'false',
  chooseAddressMappingPolicy: undefined,
  context: undefined,
  currencySymbol: ['$'],
  dateFormats: [],
  functionArgSeparator: ',',
  decimalSeparator: '.',
  evaluateNullToZero: true,
  functionPlugins: [],
  ignorePunctuation: false,
  language: 'enGB',
  ignoreWhiteSpace: 'standard',
  leapYear1900: true,
  licenseKey: 'internal',
  localeLang: 'en-US',
  matchWholeCell: true,
  arrayColumnSeparator: ',',
  arrayRowSeparator: ';',
  maxRows: MAX_ROWS,
  maxColumns: MAX_COLS,
  nullDate: { year: 1899, month: 12, day: 30 },
  nullYear: 30,
  parseDateTime: undefined,
  precisionEpsilon: 1e-13,
  precisionRounding: 14,
  stringifyDateTime: undefined,
  stringifyDuration: undefined,
  smartRounding: true,
  thousandSeparator: ',',
  timeFormats: [],
  useArrayArithmetic: true,
  useColumnIndex: false,
  useStats: true,
  undoLimit: 100,
  useRegularExpressions: true,
  useWildcards: true,
})

const WORKPAPER_CONFIG_KEYS = [
  'accentSensitive',
  'caseSensitive',
  'caseFirst',
  'chooseAddressMappingPolicy',
  'context',
  'currencySymbol',
  'dateFormats',
  'functionArgSeparator',
  'decimalSeparator',
  'evaluateNullToZero',
  'functionPlugins',
  'ignorePunctuation',
  'language',
  'ignoreWhiteSpace',
  'leapYear1900',
  'licenseKey',
  'localeLang',
  'matchWholeCell',
  'arrayColumnSeparator',
  'arrayRowSeparator',
  'maxRows',
  'maxColumns',
  'nullDate',
  'nullYear',
  'parseDateTime',
  'precisionEpsilon',
  'precisionRounding',
  'stringifyDateTime',
  'stringifyDuration',
  'smartRounding',
  'thousandSeparator',
  'timeFormats',
  'useArrayArithmetic',
  'useColumnIndex',
  'useStats',
  'undoLimit',
  'useRegularExpressions',
  'useWildcards',
] as const satisfies readonly (keyof WorkPaperConfig)[]

const WORKPAPER_PUBLIC_ERROR_NAMES = new Set([
  'WorkPaperConfigValueTooBigError',
  'WorkPaperConfigValueTooSmallError',
  'WorkPaperEvaluationSuspendedError',
  'WorkPaperExpectedOneOfValuesError',
  'WorkPaperExpectedValueOfTypeError',
  'WorkPaperFunctionPluginValidationError',
  'WorkPaperInvalidAddressError',
  'WorkPaperInvalidArgumentsError',
  'WorkPaperLanguageAlreadyRegisteredError',
  'WorkPaperLanguageNotRegisteredError',
  'WorkPaperMissingTranslationError',
  'WorkPaperNamedExpressionDoesNotExistError',
  'WorkPaperNamedExpressionNameIsAlreadyTakenError',
  'WorkPaperNamedExpressionNameIsInvalidError',
  'WorkPaperNoOperationToRedoError',
  'WorkPaperNoOperationToUndoError',
  'WorkPaperNoRelativeAddressesAllowedError',
  'WorkPaperNoSheetWithIdError',
  'WorkPaperNoSheetWithNameError',
  'WorkPaperNotAFormulaError',
  'WorkPaperNothingToPasteError',
  'WorkPaperProtectedFunctionTranslationError',
  'WorkPaperSheetNameAlreadyTakenError',
  'WorkPaperSheetSizeLimitExceededError',
  'WorkPaperSourceLocationHasArrayError',
  'WorkPaperTargetLocationHasArrayError',
  'WorkPaperUnableToParseError',
  'WorkPaperConfigError',
  'WorkPaperSheetError',
  'WorkPaperNamedExpressionError',
  'WorkPaperClipboardError',
  'WorkPaperParseError',
  'WorkPaperOperationError',
])

const WORKPAPER_VERSION = '0.1.2'
const WORKPAPER_BUILD_DATE = '2026-04-10'
const WORKPAPER_RELEASE_DATE = '2026-04-10'

const globalCustomFunctions = new Map<string, (...args: CellValue[]) => EvaluationResult | CellValue | undefined>()

let customAdapterInstalled = false
let nextWorkbookId = 1

function hasTrackedEngineSubscription(value: unknown): value is EngineTrackedEventSubscription {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'subscribeTracked') === 'function'
}

function ensureCustomAdapterInstalled(): void {
  if (customAdapterInstalled) {
    return
  }
  installExternalFunctionAdapter({
    surface: 'host',
    resolveFunction(name) {
      const implementation = globalCustomFunctions.get(name.trim().toUpperCase())
      if (!implementation) {
        return undefined
      }
      return {
        kind: 'scalar',
        implementation: (...args: CellValue[]) => {
          const result = implementation(...args)
          if (!result) {
            return errorValue(ErrorCode.Value)
          }
          return scalarFromResult(result)
        },
      }
    },
  })
  customAdapterInstalled = true
}

function clonePluginDefinition(plugin: WorkPaperFunctionPluginDefinition): WorkPaperFunctionPluginDefinition {
  return {
    ...plugin,
    implementedFunctions: Object.fromEntries(
      Object.entries(plugin.implementedFunctions).map(([name, metadata]) => [name, { ...metadata }]),
    ),
    aliases: plugin.aliases ? { ...plugin.aliases } : undefined,
    functions: plugin.functions ? { ...plugin.functions } : undefined,
  }
}

function cloneConfig(config: WorkPaperConfig): WorkPaperConfig {
  return {
    ...config,
    chooseAddressMappingPolicy: config.chooseAddressMappingPolicy ? { ...config.chooseAddressMappingPolicy } : undefined,
    context: config.context !== undefined ? structuredClone(config.context) : undefined,
    currencySymbol: config.currencySymbol ? [...config.currencySymbol] : undefined,
    dateFormats: config.dateFormats ? [...config.dateFormats] : undefined,
    functionPlugins: config.functionPlugins ? config.functionPlugins.map((plugin) => clonePluginDefinition(plugin)) : undefined,
    nullDate: config.nullDate ? { ...config.nullDate } : undefined,
    timeFormats: config.timeFormats ? [...config.timeFormats] : undefined,
  }
}

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty }
}

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function scalarValueFromLiteral(value: LiteralInput): CellValue {
  if (value === null) {
    return emptyValue()
  }
  if (typeof value === 'number') {
    return { tag: ValueTag.Number, value }
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  return { tag: ValueTag.String, value, stringId: 0 }
}

function scalarFromResult(result: EvaluationResult | CellValue): CellValue {
  if (!isArrayValue(result)) {
    return result
  }
  return result.values[0] ?? emptyValue()
}

function valuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false
  }
  switch (left.tag) {
    case ValueTag.Number:
      return right.tag === ValueTag.Number && left.value === right.value
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code
    case ValueTag.Empty:
      return true
    default:
      return false
  }
}

function matrixValuesEqual(left: CellValue | CellValue[][] | undefined, right: CellValue | CellValue[][] | undefined): boolean {
  if (!left && !right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  if (isCellValueMatrix(left) !== isCellValueMatrix(right)) {
    return false
  }
  if (!isCellValueMatrix(left) && !isCellValueMatrix(right)) {
    return valuesEqual(left, right)
  }
  if (!isCellValueMatrix(left) || !isCellValueMatrix(right)) {
    return false
  }
  const leftMatrix = left
  const rightMatrix = right
  if (leftMatrix.length !== rightMatrix.length) {
    return false
  }
  return leftMatrix.every((row: CellValue[], rowIndex: number) => {
    const otherRow = rightMatrix[rowIndex]
    if (!otherRow || row.length !== otherRow.length) {
      return false
    }
    return row.every((value: CellValue, columnIndex: number) => {
      const otherValue = otherRow[columnIndex]
      if (!otherValue) {
        return false
      }
      return valuesEqual(value, otherValue)
    })
  })
}

function normalizeName(name: string): string {
  return name.trim().toUpperCase()
}

function makeNamedExpressionKey(name: string, scope?: number): string {
  return `${scope ?? 'workbook'}:${normalizeName(name)}`
}

function makeInternalScopedName(scope: number, name: string): string {
  return `__BILIG_WORKPAPER_SCOPE_${scope}_${normalizeName(name)}`
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === 'string' && content.trim().startsWith('=')
}

function isCellValueMatrix(value: CellValue | CellValue[][]): value is CellValue[][] {
  return Array.isArray(value)
}

function isWorkPaperSheetMatrix(value: RawCellContent | WorkPaperSheet): value is WorkPaperSheet {
  return Array.isArray(value)
}

function matrixContainsFormulaContent(content: WorkPaperSheet): boolean {
  return content.some((row) => row.some((cell) => isFormulaContent(cell)))
}

function isDeferredBatchLiteralContent(content: RawCellContent): boolean {
  return content === null || typeof content === 'boolean' || typeof content === 'number' || typeof content === 'string'
}

function canUseInitialMixedSheetFastPath(content: WorkPaperSheet): boolean {
  return content.some((row) => row.some((value) => typeof value === 'string' && value.trim().startsWith('=')))
}

function stripLeadingEquals(formula: string): string {
  return formula.trim().startsWith('=') ? formula.trim().slice(1) : formula.trim()
}

function assertRowAndColumn(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new WorkPaperInvalidArgumentsError(`${label} to be a non-negative integer`)
  }
}

function assertRange(range: WorkPaperCellRange): void {
  assertRowAndColumn(range.start.sheet, 'start.sheet')
  assertRowAndColumn(range.start.row, 'start.row')
  assertRowAndColumn(range.start.col, 'start.col')
  assertRowAndColumn(range.end.sheet, 'end.sheet')
  assertRowAndColumn(range.end.row, 'end.row')
  assertRowAndColumn(range.end.col, 'end.col')
  if (range.start.sheet !== range.end.sheet) {
    throw new WorkPaperInvalidArgumentsError('Ranges must stay on a single sheet')
  }
}

function isCellRange(value: WorkPaperAddressLike): value is WorkPaperCellRange {
  return 'start' in value && 'end' in value
}

function cloneCellValue(value: CellValue): CellValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return emptyValue()
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: value.value }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: value.value }
    case ValueTag.String:
      return { tag: ValueTag.String, value: value.value, stringId: value.stringId }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: value.code }
    default:
      return emptyValue()
  }
}

function transformFormulaNode(node: FormulaNode, transform: (current: FormulaNode) => FormulaNode): FormulaNode {
  const current = transform(node)
  switch (current.kind) {
    case 'BooleanLiteral':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return current
    case 'UnaryExpr':
      return {
        ...current,
        argument: transformFormulaNode(current.argument, transform),
      }
    case 'BinaryExpr':
      return {
        ...current,
        left: transformFormulaNode(current.left, transform),
        right: transformFormulaNode(current.right, transform),
      }
    case 'CallExpr':
      return {
        ...current,
        args: current.args.map((argument) => transformFormulaNode(argument, transform)),
      }
    case 'InvokeExpr':
      return {
        ...current,
        callee: transformFormulaNode(current.callee, transform),
        args: current.args.map((argument) => transformFormulaNode(argument, transform)),
      }
    default:
      return current
  }
}

function collectFormulaNameRefs(node: FormulaNode, output: Set<string>): void {
  switch (node.kind) {
    case 'BooleanLiteral':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'NameRef':
      if (node.kind === 'NameRef') {
        output.add(node.name)
      }
      return
    case 'NumberLiteral':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return
    case 'UnaryExpr':
      collectFormulaNameRefs(node.argument, output)
      return
    case 'BinaryExpr':
      collectFormulaNameRefs(node.left, output)
      collectFormulaNameRefs(node.right, output)
      return
    case 'CallExpr':
      node.args.forEach((argument) => collectFormulaNameRefs(argument, output))
      return
    case 'InvokeExpr':
      collectFormulaNameRefs(node.callee, output)
      node.args.forEach((argument) => collectFormulaNameRefs(argument, output))
      return
    default:
      return
  }
}

function isAbsoluteCellReference(value: string): boolean {
  return /^\$[A-Z]+\$[1-9][0-9]*$/.test(value.toUpperCase())
}

function isAbsoluteRowReference(value: string): boolean {
  return /^\$[1-9][0-9]*$/.test(value)
}

function isAbsoluteColumnReference(value: string): boolean {
  return /^\$[A-Z]+$/.test(value.toUpperCase())
}

function formulaHasRelativeReferences(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'BooleanLiteral':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'StringLiteral':
    case 'StructuredRef':
      return false
    case 'CellRef':
    case 'SpillRef':
      return !isAbsoluteCellReference(node.ref)
    case 'RowRef':
      return !isAbsoluteRowReference(node.ref)
    case 'ColumnRef':
      return !isAbsoluteColumnReference(node.ref)
    case 'RangeRef':
      if (node.refKind === 'cells') {
        return !isAbsoluteCellReference(node.start) || !isAbsoluteCellReference(node.end)
      }
      if (node.refKind === 'rows') {
        return !isAbsoluteRowReference(node.start) || !isAbsoluteRowReference(node.end)
      }
      return !isAbsoluteColumnReference(node.start) || !isAbsoluteColumnReference(node.end)
    case 'UnaryExpr':
      return formulaHasRelativeReferences(node.argument)
    case 'BinaryExpr':
      return formulaHasRelativeReferences(node.left) || formulaHasRelativeReferences(node.right)
    case 'CallExpr':
      return node.args.some((argument) => formulaHasRelativeReferences(argument))
    case 'InvokeExpr':
      return formulaHasRelativeReferences(node.callee) || node.args.some((argument) => formulaHasRelativeReferences(argument))
    default:
      return false
  }
}

function compareSheetNames(left: string, right: string): number {
  return left.localeCompare(right)
}

function checkWorkPaperLicenseKeyValidity(licenseKey: string | undefined): WorkPaperLicenseKeyValidityState {
  if (!licenseKey || licenseKey.trim().length === 0) {
    return 'missing'
  }
  if (licenseKey === 'internal' || licenseKey === 'gpl-v3' || licenseKey === 'internal-use-in-handsontable') {
    return 'valid'
  }
  return 'invalid'
}

function validateWorkPaperConfig(config: WorkPaperConfig): void {
  if (config.maxRows !== undefined && (!Number.isInteger(config.maxRows) || config.maxRows < 1)) {
    throw new WorkPaperConfigValueTooSmallError('maxRows', 1)
  }
  if (config.maxColumns !== undefined && (!Number.isInteger(config.maxColumns) || config.maxColumns < 1)) {
    throw new WorkPaperConfigValueTooSmallError('maxColumns', 1)
  }
  if ((config.maxRows ?? MAX_ROWS) > MAX_ROWS) {
    throw new WorkPaperConfigValueTooBigError('maxRows', MAX_ROWS)
  }
  if ((config.maxColumns ?? MAX_COLS) > MAX_COLS) {
    throw new WorkPaperConfigValueTooBigError('maxColumns', MAX_COLS)
  }
  if (config.decimalSeparator !== undefined && config.decimalSeparator !== '.' && config.decimalSeparator !== ',') {
    throw new WorkPaperExpectedOneOfValuesError('".", ","', 'decimalSeparator')
  }
  if (config.arrayColumnSeparator !== undefined && config.arrayColumnSeparator !== ',' && config.arrayColumnSeparator !== ';') {
    throw new WorkPaperExpectedOneOfValuesError('",", ";"', 'arrayColumnSeparator')
  }
  if (config.arrayRowSeparator !== undefined && config.arrayRowSeparator !== ';' && config.arrayRowSeparator !== '|') {
    throw new WorkPaperExpectedOneOfValuesError('";", "|"', 'arrayRowSeparator')
  }
  if (config.ignoreWhiteSpace !== undefined && config.ignoreWhiteSpace !== 'standard' && config.ignoreWhiteSpace !== 'any') {
    throw new WorkPaperExpectedOneOfValuesError('"standard", "any"', 'ignoreWhiteSpace')
  }
  if (config.caseFirst !== undefined && config.caseFirst !== 'upper' && config.caseFirst !== 'lower' && config.caseFirst !== 'false') {
    throw new WorkPaperExpectedOneOfValuesError('"upper", "lower", "false"', 'caseFirst')
  }
  if (
    config.chooseAddressMappingPolicy !== undefined &&
    (typeof config.chooseAddressMappingPolicy !== 'object' ||
      config.chooseAddressMappingPolicy === null ||
      (config.chooseAddressMappingPolicy.mode !== 'dense' && config.chooseAddressMappingPolicy.mode !== 'sparse'))
  ) {
    throw new WorkPaperExpectedOneOfValuesError('"dense", "sparse"', 'chooseAddressMappingPolicy.mode')
  }
  if (config.parseDateTime !== undefined && typeof config.parseDateTime !== 'function') {
    throw new WorkPaperExpectedValueOfTypeError('function', 'parseDateTime')
  }
  if (config.stringifyDateTime !== undefined && typeof config.stringifyDateTime !== 'function') {
    throw new WorkPaperExpectedValueOfTypeError('function', 'stringifyDateTime')
  }
  if (config.stringifyDuration !== undefined && typeof config.stringifyDuration !== 'function') {
    throw new WorkPaperExpectedValueOfTypeError('function', 'stringifyDuration')
  }
  if (config.context !== undefined) {
    try {
      structuredClone(config.context)
    } catch {
      throw new WorkPaperExpectedValueOfTypeError('structured-cloneable value', 'context')
    }
  }
}

function validateSheetWithinLimits(sheetName: string, sheet: WorkPaperSheet, config: WorkPaperConfig): void {
  const height = sheet.length
  const width = Math.max(0, ...sheet.map((row) => row.length))
  if (height > (config.maxRows ?? MAX_ROWS) || width > (config.maxColumns ?? MAX_COLS)) {
    throw new WorkPaperSheetSizeLimitExceededError()
  }
  sheet.forEach((row) => {
    if (!Array.isArray(row)) {
      throw new WorkPaperUnableToParseError({ sheetName, reason: 'Rows must be arrays' })
    }
  })
}

function functionPluginIds(config: WorkPaperConfig): string[] {
  return (config.functionPlugins ?? []).map((plugin) => plugin.id).toSorted()
}

function isHistoryRecordArray(value: unknown): value is HistoryRecord[] {
  return Array.isArray(value)
}

function withEventChanges(event: QueuedEvent, changes: WorkPaperChange[]): QueuedEvent {
  switch (event.eventName) {
    case 'sheetAdded':
      return event
    case 'sheetRemoved':
      return {
        eventName: 'sheetRemoved',
        payload: {
          ...event.payload,
          changes,
        },
      }
    case 'sheetRenamed':
      return event
    case 'namedExpressionAdded':
      return {
        eventName: 'namedExpressionAdded',
        payload: {
          ...event.payload,
          changes,
        },
      }
    case 'namedExpressionRemoved':
      return {
        eventName: 'namedExpressionRemoved',
        payload: {
          ...event.payload,
          changes,
        },
      }
  }
}

function quoteSheetNameIfNeeded(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function formatQualifiedCellAddress(sheetName: string | undefined, row: number, col: number): string {
  const base = formatAddress(row, col)
  return sheetName ? `${quoteSheetNameIfNeeded(sheetName)}!${base}` : base
}

class WorkPaperEmitter {
  private readonly listeners: ListenerMap = {
    sheetAdded: new Set(),
    sheetRemoved: new Set(),
    sheetRenamed: new Set(),
    namedExpressionAdded: new Set(),
    namedExpressionRemoved: new Set(),
    valuesUpdated: new Set(),
    evaluationSuspended: new Set(),
    evaluationResumed: new Set(),
  }

  private readonly detailedListeners: DetailedListenerMap = {
    sheetAdded: new Set(),
    sheetRemoved: new Set(),
    sheetRenamed: new Set(),
    namedExpressionAdded: new Set(),
    namedExpressionRemoved: new Set(),
    valuesUpdated: new Set(),
    evaluationSuspended: new Set(),
    evaluationResumed: new Set(),
  }

  on<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    this.listeners[eventName].add(listener)
  }

  off<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    this.listeners[eventName].delete(listener)
  }

  once<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    const wrapper: WorkPaperListener<EventName> = (...args) => {
      this.off(eventName, wrapper)
      listener(...args)
    }
    this.on(eventName, wrapper)
  }

  onDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    this.detailedListeners[eventName].add(listener)
  }

  offDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    this.detailedListeners[eventName].delete(listener)
  }

  onceDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    const wrapper: WorkPaperDetailedListener<EventName> = (payload) => {
      this.offDetailed(eventName, wrapper)
      listener(payload)
    }
    this.onDetailed(eventName, wrapper)
  }

  emitDetailed(event: DetailedEvent): void {
    this.dispatchDetailed(event)
  }

  private dispatchDetailed(event: DetailedEvent): void {
    switch (event.eventName) {
      case 'sheetAdded':
        this.listeners.sheetAdded.forEach((listener) => {
          listener(event.payload.sheetName)
        })
        this.detailedListeners.sheetAdded.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'sheetRemoved':
        this.listeners.sheetRemoved.forEach((listener) => {
          listener(event.payload.sheetName, event.payload.changes)
        })
        this.detailedListeners.sheetRemoved.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'sheetRenamed':
        this.listeners.sheetRenamed.forEach((listener) => {
          listener(event.payload.oldName, event.payload.newName)
        })
        this.detailedListeners.sheetRenamed.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'namedExpressionAdded':
        this.listeners.namedExpressionAdded.forEach((listener) => {
          listener(event.payload.name, event.payload.changes)
        })
        this.detailedListeners.namedExpressionAdded.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'namedExpressionRemoved':
        this.listeners.namedExpressionRemoved.forEach((listener) => {
          listener(event.payload.name, event.payload.changes)
        })
        this.detailedListeners.namedExpressionRemoved.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'valuesUpdated':
        this.listeners.valuesUpdated.forEach((listener) => {
          listener(event.payload.changes)
        })
        this.detailedListeners.valuesUpdated.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'evaluationSuspended':
        this.listeners.evaluationSuspended.forEach((listener) => {
          listener()
        })
        this.detailedListeners.evaluationSuspended.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'evaluationResumed':
        this.listeners.evaluationResumed.forEach((listener) => {
          listener(event.payload.changes)
        })
        this.detailedListeners.evaluationResumed.forEach((listener) => {
          listener(event.payload)
        })
    }
  }

  clear(): void {
    Object.values(this.listeners).forEach((listeners) => listeners.clear())
    Object.values(this.detailedListeners).forEach((listeners) => listeners.clear())
  }
}

export class WorkPaper {
  static version = WORKPAPER_VERSION
  static buildDate = WORKPAPER_BUILD_DATE
  static releaseDate = WORKPAPER_RELEASE_DATE
  static readonly languages: Record<string, WorkPaperLanguagePackage> = {}
  static readonly defaultConfig: WorkPaperConfig = cloneConfig(DEFAULT_CONFIG)

  private static readonly languageRegistry = new Map<string, WorkPaperLanguagePackage>()
  private static readonly functionPluginRegistry = new Map<string, WorkPaperFunctionPluginDefinition>()

  readonly workbookId = nextWorkbookId++
  private engine: SpreadsheetEngine
  private readonly emitter = new WorkPaperEmitter()
  private readonly namedExpressions = new Map<string, InternalNamedExpression>()
  private readonly functionSnapshot = new Map<string, InternalFunctionBinding>()
  private readonly functionAliasLookup = new Map<string, InternalFunctionBinding>()
  private readonly internalFunctionLookup = new Map<string, InternalFunctionBinding>()
  readonly internals: WorkPaperInternals
  private config: WorkPaperConfig
  private clipboard: ClipboardPayload | null = null
  private visibilityCache: VisibilitySnapshot | null = null
  private namedExpressionValueCache: NamedExpressionValueSnapshot | null = null
  private sheetRecordsCache: readonly SheetRecord[] | null = null
  private batchDepth = 0
  private batchStartVisibility: VisibilitySnapshot | null = null
  private batchStartNamedValues: NamedExpressionValueSnapshot | null = null
  private batchUsesTrackedFastPath = false
  private batchUndoStackLength = 0
  private pendingBatchOps: EngineCellMutationRef[] = []
  private pendingBatchPotentialNewCells = 0
  private evaluationSuspended = false
  private suspendedVisibility: VisibilitySnapshot | null = null
  private suspendedNamedValues: NamedExpressionValueSnapshot | null = null
  private suspendedUsesTrackedFastPath = false
  private suspendedCellMutationRefs: EngineCellMutationRef[] = []
  private suspendedCellMutationPotentialNewCells = 0
  private queuedEvents: QueuedEvent[] = []
  private trackedEngineEvents: TrackedEngineEvent[] = []
  private engineEventCaptureEnabled = true
  private unsubscribeEngineEvents: (() => void) | null = null
  private disposed = false

  private constructor(configInput: WorkPaperConfig = {}) {
    ensureCustomAdapterInstalled()
    validateWorkPaperConfig(configInput)
    this.config = {
      ...cloneConfig(DEFAULT_CONFIG),
      ...cloneConfig(configInput),
    }
    this.engine = new SpreadsheetEngine({
      workbookName: 'Workbook',
      useColumnIndex: this.config.useColumnIndex,
      trackReplicaVersions: false,
    })
    this.attachEngineEventTracking()
    this.captureFunctionRegistry()
    this.internals = Object.freeze({
      graph: Object.freeze<WorkPaperGraphAdapter>({
        getDependents: (reference) => this.getCellDependents(reference),
        getPrecedents: (reference) => this.getCellPrecedents(reference),
      }),
      rangeMapping: Object.freeze<WorkPaperRangeMappingAdapter>({
        getValues: (range) => this.getRangeValues(range),
        getSerialized: (range) => this.getRangeSerialized(range),
      }),
      arrayMapping: Object.freeze<WorkPaperArrayMappingAdapter>({
        isPartOfArray: (address) => this.isCellPartOfArray(address),
        getFormula: (address) => this.getCellFormula(address),
      }),
      sheetMapping: Object.freeze<WorkPaperSheetMappingAdapter>({
        getSheetName: (sheetId) => this.getSheetName(sheetId),
        getSheetId: (name) => this.getSheetId(name),
        getSheetNames: () => this.getSheetNames(),
        countSheets: () => this.countSheets(),
      }),
      addressMapping: Object.freeze<WorkPaperAddressMappingAdapter>({
        has: (address) => !this.isCellEmpty(address) || this.doesCellHaveFormula(address),
        getValue: (address) => this.getCellValue(address),
        getFormula: (address) => this.getCellFormula(address),
      }),
      dependencyGraph: Object.freeze<WorkPaperDependencyGraphAdapter>({
        getCellDependents: (reference) => this.getCellDependents(reference),
        getCellPrecedents: (reference) => this.getCellPrecedents(reference),
      }),
      evaluator: Object.freeze<WorkPaperEvaluatorAdapter>({
        recalculate: () => this.rebuildAndRecalculate(),
        calculateFormula: (formula, scope) => this.calculateFormula(formula, scope),
      }),
      columnSearch: Object.freeze<WorkPaperColumnSearchAdapter>({
        find: (sheetId, column, matcher) => {
          const dimensions = this.getSheetDimensions(sheetId)
          const matches: WorkPaperCellAddress[] = []
          for (let row = 0; row < dimensions.height; row += 1) {
            const address = { sheet: sheetId, row, col: column }
            const value = this.getCellValue(address)
            const isMatch = typeof matcher === 'string' ? value.tag === ValueTag.String && value.value === matcher : matcher(value)
            if (isMatch) {
              matches.push(address)
            }
          }
          return matches
        },
      }),
      lazilyTransformingAstService: Object.freeze<WorkPaperLazilyTransformingAstServiceAdapter>({
        normalizeFormula: (formula) => this.normalizeFormula(formula),
        validateFormula: (formula) => this.validateFormula(formula),
        getNamedExpressionsFromFormula: (formula) => this.getNamedExpressionsFromFormula(formula),
      }),
    })
  }

  static buildEmpty(configInput: WorkPaperConfig = {}, namedExpressions: readonly SerializedWorkPaperNamedExpression[] = []): WorkPaper {
    const workbook = new WorkPaper(configInput)
    workbook.withEngineEventCaptureDisabled(() => {
      namedExpressions.forEach((expression) => {
        workbook.upsertNamedExpressionInternal(expression, { duringInitialization: true })
      })
    })
    workbook.clearHistoryStacks()
    workbook.resetChangeTrackingCaches()
    return workbook
  }

  static buildFromArray(
    sheet: WorkPaperSheet,
    configInput: WorkPaperConfig = {},
    namedExpressions: readonly SerializedWorkPaperNamedExpression[] = [],
  ): WorkPaper {
    return this.buildFromSheets({ Sheet1: sheet }, configInput, namedExpressions)
  }

  static buildFromSheets(
    sheets: WorkPaperSheets,
    configInput: WorkPaperConfig = {},
    namedExpressions: readonly SerializedWorkPaperNamedExpression[] = [],
  ): WorkPaper {
    const workbook = new WorkPaper(configInput)
    Object.entries(sheets).forEach(([sheetName, sheet]) => {
      validateSheetWithinLimits(sheetName, sheet, workbook.config)
    })
    workbook.withEngineEventCaptureDisabled(() => {
      Object.keys(sheets).forEach((sheetName) => {
        workbook.engine.createSheet(sheetName)
      })
      namedExpressions.forEach((expression) => {
        workbook.upsertNamedExpressionInternal(expression, { duringInitialization: true })
      })
      Object.entries(sheets).forEach(([sheetName, sheet]) => {
        const sheetId = workbook.requireSheetId(sheetName)
        if (tryLoadInitialLiteralSheet(workbook.engine, sheetId, sheet)) {
          return
        }
        if (canUseInitialMixedSheetFastPath(sheet)) {
          loadInitialMixedSheet({
            engine: workbook.engine,
            sheetId,
            content: sheet,
            rewriteFormula: (formula, destination) => workbook.rewriteFormulaForStorage(formula, destination.sheet),
          })
          return
        }
        workbook.replaceSheetContentInternal(sheetId, sheet, { duringInitialization: true })
      })
    })
    workbook.clearHistoryStacks()
    workbook.resetChangeTrackingCaches()
    return workbook
  }

  static getLanguage(languageCode: string): WorkPaperLanguagePackage {
    const language = this.languageRegistry.get(languageCode)
    if (!language) {
      throw new WorkPaperLanguageNotRegisteredError(languageCode)
    }
    return structuredClone(language)
  }

  static registerLanguage(languageCode: string, languagePackage: WorkPaperLanguagePackage): void {
    if (this.languageRegistry.has(languageCode)) {
      throw new WorkPaperLanguageAlreadyRegisteredError(languageCode)
    }
    this.languageRegistry.set(languageCode, structuredClone(languagePackage))
    this.languages[languageCode] = structuredClone(languagePackage)
  }

  static unregisterLanguage(languageCode: string): void {
    if (!this.languageRegistry.delete(languageCode)) {
      throw new WorkPaperLanguageNotRegisteredError(languageCode)
    }
    delete this.languages[languageCode]
  }

  static getRegisteredLanguagesCodes(): string[] {
    return [...this.languageRegistry.keys()].toSorted(compareSheetNames)
  }

  static registerFunctionPlugin(plugin: WorkPaperFunctionPluginDefinition, translations?: WorkPaperFunctionTranslationsPackage): void {
    this.functionPluginRegistry.set(plugin.id, clonePluginDefinition(plugin))
    if (translations) {
      this.loadFunctionTranslations(translations)
    }
  }

  static unregisterFunctionPlugin(plugin: WorkPaperFunctionPluginDefinition | string): void {
    const pluginId = typeof plugin === 'string' ? plugin : plugin.id
    this.functionPluginRegistry.delete(pluginId)
  }

  static registerFunction(
    functionId: string,
    plugin: WorkPaperFunctionPluginDefinition,
    translations?: WorkPaperFunctionTranslationsPackage,
  ): void {
    const existing = this.functionPluginRegistry.get(plugin.id)
    const nextPlugin = clonePluginDefinition(existing ?? plugin)
    if (!nextPlugin.implementedFunctions[functionId]) {
      throw WorkPaperFunctionPluginValidationError.functionNotDeclaredInPlugin(functionId, plugin.id)
    }
    this.functionPluginRegistry.set(nextPlugin.id, nextPlugin)
    if (translations) {
      this.loadFunctionTranslations(translations)
    }
  }

  static unregisterFunction(functionId: string): void {
    const normalized = functionId.trim().toUpperCase()
    this.functionPluginRegistry.forEach((plugin, pluginId) => {
      if (!plugin.implementedFunctions[normalized]) {
        return
      }
      const nextPlugin = clonePluginDefinition(plugin)
      delete nextPlugin.implementedFunctions[normalized]
      if (nextPlugin.functions) {
        delete nextPlugin.functions[normalized]
      }
      if (nextPlugin.aliases) {
        Object.entries(nextPlugin.aliases).forEach(([alias, target]) => {
          if (target.trim().toUpperCase() === normalized || alias.trim().toUpperCase() === normalized) {
            delete nextPlugin.aliases![alias]
          }
        })
      }
      this.functionPluginRegistry.set(pluginId, nextPlugin)
    })
  }

  static unregisterAllFunctions(): void {
    this.functionPluginRegistry.clear()
  }

  static getRegisteredFunctionNames(languageCode?: string): string[] {
    const normalized = languageCode ?? 'enGB'
    const language = this.languageRegistry.get(normalized)
    const functions = [...this.functionPluginRegistry.values()].flatMap((plugin) => Object.keys(plugin.implementedFunctions))
    if (!language?.functions) {
      return functions.toSorted(compareSheetNames)
    }
    return functions.map((name) => language.functions?.[name] ?? name).toSorted(compareSheetNames)
  }

  static getFunctionPlugin(functionId: string): WorkPaperFunctionPluginDefinition | undefined {
    const normalized = functionId.trim().toUpperCase()
    const plugin = [...this.functionPluginRegistry.values()].find(
      (candidate) => candidate.implementedFunctions[normalized] !== undefined || candidate.aliases?.[normalized] !== undefined,
    )
    return plugin ? clonePluginDefinition(plugin) : undefined
  }

  static getAllFunctionPlugins(): WorkPaperFunctionPluginDefinition[] {
    return [...this.functionPluginRegistry.values()].map((plugin) => clonePluginDefinition(plugin))
  }

  private static loadFunctionTranslations(translations: WorkPaperFunctionTranslationsPackage): void {
    Object.entries(translations).forEach(([languageCode, functionTranslations]) => {
      const existing = this.languageRegistry.get(languageCode)
      if (!existing) {
        throw new WorkPaperLanguageNotRegisteredError(languageCode)
      }
      const nextLanguage: WorkPaperLanguagePackage = {
        ...structuredClone(existing),
        functions: {
          ...existing.functions,
          ...functionTranslations,
        },
      }
      this.languageRegistry.set(languageCode, nextLanguage)
      this.languages[languageCode] = structuredClone(nextLanguage)
    })
  }

  on<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    this.assertNotDisposed()
    this.emitter.on(eventName, listener)
  }

  once<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    this.assertNotDisposed()
    this.emitter.once(eventName, listener)
  }

  off<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    this.emitter.off(eventName, listener)
  }

  onDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    this.assertNotDisposed()
    this.emitter.onDetailed(eventName, listener)
  }

  onceDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    this.assertNotDisposed()
    this.emitter.onceDetailed(eventName, listener)
  }

  offDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    this.emitter.offDetailed(eventName, listener)
  }

  getConfig(): WorkPaperConfig {
    return cloneConfig(this.config)
  }

  get graph(): WorkPaperGraphAdapter {
    return this.internals.graph
  }

  get rangeMapping(): WorkPaperRangeMappingAdapter {
    return this.internals.rangeMapping
  }

  get arrayMapping(): WorkPaperArrayMappingAdapter {
    return this.internals.arrayMapping
  }

  get sheetMapping(): WorkPaperSheetMappingAdapter {
    return this.internals.sheetMapping
  }

  get addressMapping(): WorkPaperAddressMappingAdapter {
    return this.internals.addressMapping
  }

  get dependencyGraph(): WorkPaperDependencyGraphAdapter {
    return this.internals.dependencyGraph
  }

  get evaluator(): WorkPaperEvaluatorAdapter {
    return this.internals.evaluator
  }

  get columnSearch(): WorkPaperColumnSearchAdapter {
    return this.internals.columnSearch
  }

  get lazilyTransformingAstService(): WorkPaperLazilyTransformingAstServiceAdapter {
    return this.internals.lazilyTransformingAstService
  }

  get licenseKeyValidityState(): WorkPaperLicenseKeyValidityState {
    return checkWorkPaperLicenseKeyValidity(this.config.licenseKey)
  }

  updateConfig(next: WorkPaperConfig): void {
    this.assertNotDisposed()
    const merged = {
      ...this.config,
      ...cloneConfig(next),
    }
    const changedKeys = WORKPAPER_CONFIG_KEYS.filter((key) => Object.hasOwn(next, key) && this.config[key] !== merged[key])
    if (changedKeys.length === 0) {
      return
    }
    validateWorkPaperConfig(merged)
    if (this.canApplyRuntimeOnlyConfigUpdate(changedKeys)) {
      this.applyRuntimeOnlyConfigUpdate(merged)
      return
    }
    this.rebuildWithConfig(merged)
  }

  getStats(): WorkPaperStats {
    this.assertNotDisposed()
    return {
      batchDepth: this.batchDepth,
      evaluationSuspended: this.evaluationSuspended,
      lastMetrics: structuredClone(this.engine.getLastMetrics()),
    }
  }

  rebuildAndRecalculate(): WorkPaperChange[] {
    this.assertNotDisposed()
    if (this.shouldSuppressEvents()) {
      try {
        this.engine.recalculateNow()
      } catch (error) {
        throw new WorkPaperOperationError(this.messageOf(error, 'Recalculation failed'))
      }
      return []
    }
    const beforeVisibility = this.ensureVisibilityCache()
    const beforeNames = this.namedExpressions.size > 0 ? this.ensureNamedExpressionValueCache() : EMPTY_NAMED_EXPRESSION_VALUES
    this.drainTrackedEngineEvents()
    try {
      this.engine.recalculateNow()
    } catch (error) {
      throw new WorkPaperOperationError(this.messageOf(error, 'Recalculation failed'))
    }
    const afterVisibility = this.captureVisibilitySnapshot()
    const afterNames = this.namedExpressions.size > 0 ? this.captureNamedExpressionValueSnapshot() : EMPTY_NAMED_EXPRESSION_VALUES
    this.visibilityCache = afterVisibility
    this.namedExpressionValueCache = afterNames
    const changes = [
      ...this.computeCellChanges(beforeVisibility, afterVisibility),
      ...this.computeNamedExpressionChanges(beforeNames, afterNames),
    ]
    if (changes.length > 0) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  batch(batchOperations: () => void): WorkPaperChange[] {
    this.assertNotDisposed()
    const isOutermost = this.batchDepth === 0
    if (isOutermost) {
      this.batchUsesTrackedFastPath = this.canUseTrackedMutationFastPath()
      if (this.batchUsesTrackedFastPath) {
        this.batchStartVisibility = null
        this.batchStartNamedValues = EMPTY_NAMED_EXPRESSION_VALUES
      } else {
        this.batchStartVisibility = this.ensureVisibilityCache()
        this.batchStartNamedValues = this.ensureNamedExpressionValueCache()
      }
      this.batchUndoStackLength = this.getUndoStack().length
      this.drainTrackedEngineEvents()
    }
    this.batchDepth += 1
    try {
      batchOperations()
    } finally {
      this.batchDepth -= 1
      if (isOutermost) {
        this.flushPendingBatchOps()
        this.mergeUndoHistory(this.batchUndoStackLength)
      }
    }
    if (!isOutermost) {
      return []
    }
    const changes = this.batchUsesTrackedFastPath
      ? this.computeTrackedChangesWithoutVisibilityCache(this.drainTrackedEngineEvents())
      : this.computeChangesAfterMutation(this.batchStartVisibility ?? new Map(), this.batchStartNamedValues ?? new Map())
    this.batchUsesTrackedFastPath = false
    this.batchStartVisibility = null
    this.batchStartNamedValues = null
    if (!this.evaluationSuspended) {
      this.flushQueuedEvents()
      if (changes.length > 0) {
        this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
      }
    }
    return changes
  }

  suspendEvaluation(): void {
    this.assertNotDisposed()
    if (this.evaluationSuspended) {
      return
    }
    this.evaluationSuspended = true
    this.flushPendingBatchOps()
    if (this.visibilityCache === null && this.namedExpressions.size === 0) {
      this.suspendedVisibility = null
      this.suspendedNamedValues = EMPTY_NAMED_EXPRESSION_VALUES
      this.suspendedUsesTrackedFastPath = true
    } else {
      this.suspendedVisibility = this.ensureVisibilityCache()
      this.suspendedNamedValues = this.ensureNamedExpressionValueCache()
      this.suspendedUsesTrackedFastPath = false
    }
    this.drainTrackedEngineEvents()
    this.emitter.emitDetailed({ eventName: 'evaluationSuspended', payload: {} })
  }

  resumeEvaluation(): WorkPaperChange[] {
    this.assertNotDisposed()
    if (!this.evaluationSuspended) {
      return []
    }
    this.flushSuspendedCellMutations()
    const changes = this.suspendedUsesTrackedFastPath
      ? this.computeTrackedChangesWithoutVisibilityCache(this.drainTrackedEngineEvents())
      : this.computeChangesAfterMutation(this.suspendedVisibility ?? new Map(), this.suspendedNamedValues ?? new Map())
    this.evaluationSuspended = false
    this.suspendedVisibility = null
    this.suspendedNamedValues = null
    this.suspendedUsesTrackedFastPath = false
    this.flushQueuedEvents()
    this.emitter.emitDetailed({ eventName: 'evaluationResumed', payload: { changes } })
    if (changes.length > 0) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  isEvaluationSuspended(): boolean {
    return this.evaluationSuspended
  }

  undo(): WorkPaperChange[] {
    this.assertNotDisposed()
    return this.captureChanges(undefined, () => {
      if (!this.engine.undo()) {
        throw new WorkPaperNoOperationToUndoError()
      }
    })
  }

  redo(): WorkPaperChange[] {
    this.assertNotDisposed()
    return this.captureChanges(undefined, () => {
      if (!this.engine.redo()) {
        throw new WorkPaperNoOperationToRedoError()
      }
    })
  }

  isThereSomethingToUndo(): boolean {
    return this.getUndoStack().length > 0
  }

  isThereSomethingToRedo(): boolean {
    return this.getRedoStack().length > 0
  }

  clearUndoStack(): void {
    this.getUndoStack().length = 0
  }

  clearRedoStack(): void {
    this.getRedoStack().length = 0
  }

  copy(range: WorkPaperCellRange): CellValue[][] {
    this.assertReadable()
    assertRange(range)
    const serialized = this.getRangeSerialized(range)
    const values = this.getRangeValues(range)
    this.clipboard = {
      sourceAnchor: { ...range.start },
      serialized,
      values,
    }
    return values
  }

  cut(range: WorkPaperCellRange): CellValue[][] {
    this.assertReadable()
    const values = this.copy(range)
    this.batch(() => {
      this.setCellContents(range.start, this.buildNullMatrixForRange(range))
    })
    return values
  }

  paste(targetLeftCorner: WorkPaperCellAddress): WorkPaperChange[] {
    this.assertNotDisposed()
    if (!this.clipboard) {
      throw new WorkPaperNothingToPasteError()
    }
    return this.captureChanges(undefined, () => {
      this.applySerializedMatrix(targetLeftCorner, this.clipboard!.serialized, this.clipboard!.sourceAnchor)
    })
  }

  isClipboardEmpty(): boolean {
    return this.clipboard === null
  }

  clearClipboard(): void {
    this.clipboard = null
  }

  getFillRangeData(source: WorkPaperCellRange, target: WorkPaperCellRange, offsetsFromTarget = false): RawCellContent[][] {
    assertRange(source)
    assertRange(target)
    const sourceSerialized = this.getRangeSerialized(source)
    const targetHeight = target.end.row - target.start.row + 1
    const targetWidth = target.end.col - target.start.col + 1
    const sourceHeight = Math.max(sourceSerialized.length, 1)
    const sourceWidth = Math.max(sourceSerialized[0]?.length ?? 0, 1)
    const output: RawCellContent[][] = []
    for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
      const row: RawCellContent[] = []
      for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
        const targetRow = target.start.row + rowOffset
        const targetCol = target.start.col + colOffset
        const sourceRow =
          (((targetRow - (offsetsFromTarget ? target.start.row : source.start.row)) % sourceHeight) + sourceHeight) % sourceHeight
        const sourceCol =
          (((targetCol - (offsetsFromTarget ? target.start.col : source.start.col)) % sourceWidth) + sourceWidth) % sourceWidth
        const raw = sourceSerialized[sourceRow]?.[sourceCol] ?? null
        if (typeof raw === 'string' && raw.startsWith('=')) {
          row.push(
            `=${translateFormulaReferences(
              raw.slice(1),
              targetRow - (source.start.row + sourceRow),
              targetCol - (source.start.col + sourceCol),
            )}`,
          )
        } else {
          row.push(raw)
        }
      }
      output.push(row)
    }
    return output
  }

  getCellValue(address: WorkPaperCellAddress): CellValue {
    this.assertReadable()
    return cloneCellValue(this.engine.getCellValue(this.sheetName(address.sheet), this.a1(address)))
  }

  getCellFormula(address: WorkPaperCellAddress): string | undefined {
    this.prepareReadableState()
    const cell = this.engine.getCell(this.sheetName(address.sheet), this.a1(address))
    if (!cell.formula) {
      return undefined
    }
    return `=${this.restorePublicFormula(cell.formula, address.sheet)}`
  }

  getCellHyperlink(address: WorkPaperCellAddress): string | undefined {
    const formula = this.getCellFormula(address)
    if (!formula) {
      return undefined
    }
    const parsed = parseFormula(stripLeadingEquals(formula))
    if (parsed.kind !== 'CallExpr' || parsed.callee.trim().toUpperCase() !== 'HYPERLINK') {
      return undefined
    }
    const firstArgument = parsed.args[0]
    return firstArgument?.kind === 'StringLiteral' ? firstArgument.value : undefined
  }

  getCellSerialized(address: WorkPaperCellAddress): RawCellContent {
    this.prepareReadableState()
    return this.cellSnapshotToRawContent(this.engine.getCell(this.sheetName(address.sheet), this.a1(address)), address.sheet)
  }

  getRangeValues(range: WorkPaperCellRange): CellValue[][] {
    this.assertReadable()
    const ref = this.rangeRef(range)
    return this.engine.getRangeValues(ref)
  }

  getRangeFormulas(range: WorkPaperCellRange): Array<Array<string | undefined>> {
    return this.getDenseRange(range, (address) => this.getCellFormula(address))
  }

  getRangeSerialized(range: WorkPaperCellRange): RawCellContent[][] {
    return this.getDenseRange(range, (address) => this.getCellSerialized(address))
  }

  getSheetValues(sheetId: number): CellValue[][] {
    this.assertReadable()
    const dimensions = this.getSheetDimensions(sheetId)
    if (dimensions.width === 0 || dimensions.height === 0) {
      return []
    }
    return this.getRangeValues({
      start: { sheet: sheetId, row: 0, col: 0 },
      end: { sheet: sheetId, row: dimensions.height - 1, col: dimensions.width - 1 },
    })
  }

  getSheetFormulas(sheetId: number): Array<Array<string | undefined>> {
    const dimensions = this.getSheetDimensions(sheetId)
    if (dimensions.width === 0 || dimensions.height === 0) {
      return []
    }
    return this.getRangeFormulas({
      start: { sheet: sheetId, row: 0, col: 0 },
      end: { sheet: sheetId, row: dimensions.height - 1, col: dimensions.width - 1 },
    })
  }

  getSheetSerialized(sheetId: number): RawCellContent[][] {
    const dimensions = this.getSheetDimensions(sheetId)
    if (dimensions.width === 0 || dimensions.height === 0) {
      return []
    }
    return this.getRangeSerialized({
      start: { sheet: sheetId, row: 0, col: 0 },
      end: { sheet: sheetId, row: dimensions.height - 1, col: dimensions.width - 1 },
    })
  }

  getAllSheetsValues(): Record<string, CellValue[][]> {
    this.assertReadable()
    return Object.fromEntries(this.listSheetRecords().map((sheet) => [sheet.name, this.getSheetValues(sheet.id)]))
  }

  getAllSheetsFormulas(): Record<string, Array<Array<string | undefined>>> {
    return Object.fromEntries(this.listSheetRecords().map((sheet) => [sheet.name, this.getSheetFormulas(sheet.id)]))
  }

  getAllSheetsSerialized(): Record<string, RawCellContent[][]> {
    return Object.fromEntries(this.listSheetRecords().map((sheet) => [sheet.name, this.getSheetSerialized(sheet.id)]))
  }

  getAllSheetsDimensions(): Record<string, WorkPaperSheetDimensions> {
    return Object.fromEntries(this.listSheetRecords().map((sheet) => [sheet.name, this.getSheetDimensions(sheet.id)]))
  }

  getSheetDimensions(sheetId: number): WorkPaperSheetDimensions {
    this.prepareReadableState()
    const sheet = this.sheetRecord(sheetId)
    let width = 0
    let height = 0
    sheet.grid.forEachCellEntry((_cellIndex: number, row: number, col: number) => {
      height = Math.max(height, row + 1)
      width = Math.max(width, col + 1)
    })
    return { width, height }
  }

  simpleCellAddressFromString(value: string, defaultSheetId?: number): WorkPaperCellAddress | undefined {
    this.assertNotDisposed()
    const defaultSheetName =
      defaultSheetId !== undefined
        ? this.sheetName(defaultSheetId)
        : this.listSheetRecords().length === 1
          ? this.listSheetRecords()[0]!.name
          : undefined
    try {
      const parsed = parseCellAddress(value, defaultSheetName)
      const sheetName = parsed.sheetName ?? defaultSheetName
      if (!sheetName) {
        return undefined
      }
      return {
        sheet: this.requireSheetId(sheetName),
        row: parsed.row,
        col: parsed.col,
      }
    } catch {
      return undefined
    }
  }

  simpleCellRangeFromString(value: string, defaultSheetId?: number): WorkPaperCellRange | undefined {
    this.assertNotDisposed()
    const defaultSheetName =
      defaultSheetId !== undefined
        ? this.sheetName(defaultSheetId)
        : this.listSheetRecords().length === 1
          ? this.listSheetRecords()[0]!.name
          : undefined
    try {
      const parsed = parseRangeAddress(value, defaultSheetName)
      if (parsed.kind !== 'cells') {
        return undefined
      }
      const sheetName = parsed.sheetName ?? defaultSheetName
      if (!sheetName) {
        return undefined
      }
      const sheetId = this.requireSheetId(sheetName)
      return {
        start: { sheet: sheetId, row: parsed.start.row, col: parsed.start.col },
        end: { sheet: sheetId, row: parsed.end.row, col: parsed.end.col },
      }
    } catch {
      return undefined
    }
  }

  simpleCellAddressToString(address: WorkPaperCellAddress, optionsOrContextSheetId: WorkPaperAddressFormatOptions | number = {}): string {
    this.assertNotDisposed()
    const includeSheetName =
      typeof optionsOrContextSheetId === 'number'
        ? optionsOrContextSheetId !== address.sheet
        : optionsOrContextSheetId.includeSheetName === true
    return formatQualifiedCellAddress(includeSheetName ? this.sheetName(address.sheet) : undefined, address.row, address.col)
  }

  simpleCellRangeToString(range: WorkPaperCellRange, optionsOrContextSheetId: WorkPaperAddressFormatOptions | number = {}): string {
    const includeSheetName =
      typeof optionsOrContextSheetId === 'number'
        ? optionsOrContextSheetId !== range.start.sheet
        : optionsOrContextSheetId.includeSheetName === true
    const sheetName = includeSheetName ? this.sheetName(range.start.sheet) : undefined
    return formatRangeAddress({
      kind: 'cells',
      sheetName,
      start: {
        row: range.start.row,
        col: range.start.col,
        text: formatAddress(range.start.row, range.start.col),
      },
      end: {
        row: range.end.row,
        col: range.end.col,
        text: formatAddress(range.end.row, range.end.col),
      },
    })
  }

  getCellDependents(address: WorkPaperAddressLike): WorkPaperDependencyRef[] {
    this.flushPendingBatchOps()
    if (!isCellRange(address)) {
      return this.toDependencyRefs(this.engine.getDependents(this.sheetName(address.sheet), this.a1(address)).directDependents)
    }
    return this.collectRangeDependencies(
      address,
      (cellAddress) => this.engine.getDependents(this.sheetName(cellAddress.sheet), this.a1(cellAddress)).directDependents,
    )
  }

  getCellPrecedents(address: WorkPaperAddressLike): WorkPaperDependencyRef[] {
    this.flushPendingBatchOps()
    if (!isCellRange(address)) {
      return this.getDirectPrecedentRefs(address)
    }
    return this.collectRangeDependencies(address, (cellAddress) => this.getDirectPrecedentStrings(cellAddress))
  }

  getSheetName(sheetId: number): string | undefined {
    return this.engine.workbook.getSheetById(sheetId)?.name
  }

  getSheetNames(): string[] {
    return this.listSheetRecords().map((sheet) => sheet.name)
  }

  getSheetId(name: string): number | undefined {
    return this.engine.workbook.getSheet(name)?.id
  }

  doesSheetExist(name: string): boolean {
    return this.engine.workbook.getSheet(name) !== undefined
  }

  countSheets(): number {
    return this.listSheetRecords().length
  }

  getCellType(address: WorkPaperCellAddress): WorkPaperCellType {
    this.flushPendingBatchOps()
    const cell = this.engine.getCell(this.sheetName(address.sheet), this.a1(address))
    if (this.isCellEmpty(address)) {
      return 'EMPTY'
    }
    if (this.isCellPartOfArray(address)) {
      return 'ARRAY'
    }
    return cell.formula ? 'FORMULA' : 'VALUE'
  }

  doesCellHaveSimpleValue(address: WorkPaperCellAddress): boolean {
    this.flushPendingBatchOps()
    const cell = this.engine.getCell(this.sheetName(address.sheet), this.a1(address))
    return !cell.formula && !this.isCellEmpty(address)
  }

  doesCellHaveFormula(address: WorkPaperCellAddress): boolean {
    this.flushPendingBatchOps()
    return this.engine.getCell(this.sheetName(address.sheet), this.a1(address)).formula !== undefined
  }

  isCellEmpty(address: WorkPaperCellAddress): boolean {
    this.flushPendingBatchOps()
    return this.engine.getCellValue(this.sheetName(address.sheet), this.a1(address)).tag === ValueTag.Empty
  }

  isCellPartOfArray(address: WorkPaperCellAddress): boolean {
    this.flushPendingBatchOps()
    return this.engine.getSpillRanges().some((spill: { sheetName: string; address: string; rows: number; cols: number }) => {
      if (this.requireSheetId(spill.sheetName) !== address.sheet) {
        return false
      }
      const owner = parseCellAddress(spill.address, spill.sheetName)
      return (
        address.row >= owner.row && address.row < owner.row + spill.rows && address.col >= owner.col && address.col < owner.col + spill.cols
      )
    })
  }

  getCellValueType(address: WorkPaperCellAddress): WorkPaperCellValueType {
    const value = this.getCellValue(address)
    switch (value.tag) {
      case ValueTag.Number:
        return 'NUMBER'
      case ValueTag.String:
        return 'STRING'
      case ValueTag.Boolean:
        return 'BOOLEAN'
      case ValueTag.Error:
        return 'ERROR'
      case ValueTag.Empty:
      default:
        return 'EMPTY'
    }
  }

  getCellValueDetailedType(address: WorkPaperCellAddress): WorkPaperCellValueDetailedType {
    const type = this.getCellValueType(address)
    if (type !== 'NUMBER') {
      return type
    }
    const format = this.getCellValueFormat(address)?.toLowerCase() ?? ''
    if (format.includes('yy') || format.includes('dd')) {
      if (format.includes('h') || format.includes('s')) {
        return 'DATETIME'
      }
      return 'DATE'
    }
    if (format.includes('h') || format.includes('s')) {
      return 'TIME'
    }
    return type
  }

  getCellValueFormat(address: WorkPaperCellAddress): string | undefined {
    this.flushPendingBatchOps()
    const cell = this.engine.getCell(this.sheetName(address.sheet), this.a1(address))
    return cell.format
  }

  getNamedExpressionValue(name: string, scope?: number): CellValue | CellValue[][] | undefined {
    this.assertReadable()
    const expression = this.namedExpressions.get(makeNamedExpressionKey(name, scope))
    return expression ? this.evaluateNamedExpression(expression) : undefined
  }

  getNamedExpressionFormula(name: string, scope?: number): string | undefined {
    const expression = this.namedExpressions.get(makeNamedExpressionKey(name, scope))
    if (!expression) {
      return undefined
    }
    return isFormulaContent(expression.expression) ? expression.expression : undefined
  }

  getNamedExpression(name: string, scope?: number): WorkPaperNamedExpression | undefined {
    const expression = this.namedExpressions.get(makeNamedExpressionKey(name, scope))
    if (!expression) {
      return undefined
    }
    return {
      name: expression.publicName,
      expression: expression.expression,
      scope: expression.scope,
      options: expression.options ? structuredClone(expression.options) : undefined,
    }
  }

  addNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ): WorkPaperChange[] {
    if (!this.isItPossibleToAddNamedExpression(expressionName, expression, scope)) {
      throw new WorkPaperNamedExpressionNameIsAlreadyTakenError(expressionName)
    }
    return this.captureChanges(
      {
        eventName: 'namedExpressionAdded',
        payload: {
          name: expressionName.trim(),
          scope,
          changes: [],
        },
      },
      () => {
        this.upsertNamedExpressionInternal({ name: expressionName, expression, scope, options }, { duringInitialization: false })
      },
    )
  }

  changeNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ): WorkPaperChange[] {
    if (!this.isItPossibleToChangeNamedExpression(expressionName, expression, scope)) {
      throw new WorkPaperNamedExpressionDoesNotExistError(expressionName)
    }
    return this.captureChanges(undefined, () => {
      this.upsertNamedExpressionInternal({ name: expressionName, expression, scope, options }, { duringInitialization: false })
    })
  }

  removeNamedExpression(expressionName: string, scope?: number): WorkPaperChange[] {
    if (!this.isItPossibleToRemoveNamedExpression(expressionName, scope)) {
      throw new WorkPaperNamedExpressionDoesNotExistError(expressionName)
    }
    const existing = this.namedExpressionRecord(expressionName, scope)
    return this.captureChanges(
      {
        eventName: 'namedExpressionRemoved',
        payload: {
          name: existing.publicName,
          scope: existing.scope,
          changes: [],
        },
      },
      () => {
        this.namedExpressions.delete(makeNamedExpressionKey(expressionName, scope))
        this.engine.deleteDefinedName(existing.internalName)
      },
    )
  }

  listNamedExpressions(scope?: number): string[] {
    return [...this.namedExpressions.values()]
      .filter((expression) => expression.scope === scope)
      .map((expression) => expression.publicName)
      .toSorted(compareSheetNames)
  }

  getAllNamedExpressionsSerialized(): SerializedWorkPaperNamedExpression[] {
    return [...this.namedExpressions.values()]
      .map((expression) => ({
        name: expression.publicName,
        expression: expression.expression,
        scope: expression.scope,
        options: expression.options ? structuredClone(expression.options) : undefined,
      }))
      .toSorted((left, right) => (left.scope ?? -1) - (right.scope ?? -1) || left.name.localeCompare(right.name))
  }

  normalizeFormula(formula: string): string {
    if (!formula.trim().startsWith('=')) {
      throw new WorkPaperNotAFormulaError()
    }
    try {
      return `=${serializeFormula(parseFormula(stripLeadingEquals(formula)))}`
    } catch (error) {
      throw new WorkPaperParseError(this.messageOf(error, `Unable to normalize formula`))
    }
  }

  calculateFormula(formula: string, scope?: number): CellValue | CellValue[][] {
    if (!formula.trim().startsWith('=')) {
      throw new WorkPaperNotAFormulaError()
    }
    try {
      return calculateWorkPaperFormulaInScratchWorkbook({
        createWorkbook: (config) => {
          const temporaryWorkbook = new WorkPaper(config)
          return {
            engine: temporaryWorkbook.engine,
            registerNamedExpression: (expression) => {
              temporaryWorkbook.upsertNamedExpressionInternal(expression, {
                duringInitialization: true,
              })
            },
            requireSheetId: (sheetName) => temporaryWorkbook.requireSheetId(sheetName),
            replaceSheetContent: (sheetId, sheet) => {
              temporaryWorkbook.replaceSheetContentInternal(sheetId, sheet, {
                duringInitialization: true,
              })
            },
            clearHistoryStacks: () => temporaryWorkbook.clearHistoryStacks(),
            applyRawContent: (address, content) => temporaryWorkbook.applyRawContent(address, content),
            getRangeValues: (range) => temporaryWorkbook.getRangeValues(range),
            getCellValue: (address) => temporaryWorkbook.getCellValue(address),
            dispose: () => temporaryWorkbook.dispose(),
          }
        },
        config: this.getConfig(),
        serializedSheets: this.getAllSheetsSerialized(),
        namedExpressions: this.getAllNamedExpressionsSerialized(),
        formula,
        scope,
      })
    } catch (error) {
      throw new WorkPaperParseError(this.messageOf(error, 'Unable to calculate formula'))
    }
  }

  getNamedExpressionsFromFormula(formula: string): string[] {
    if (!formula.trim().startsWith('=')) {
      throw new WorkPaperNotAFormulaError()
    }
    try {
      const parsed = parseFormula(stripLeadingEquals(formula))
      const names = new Set<string>()
      collectFormulaNameRefs(parsed, names)
      return [...names].toSorted(compareSheetNames)
    } catch (error) {
      throw new WorkPaperParseError(this.messageOf(error, 'Unable to inspect formula'))
    }
  }

  validateFormula(formula: string): boolean {
    if (!formula.trim().startsWith('=')) {
      return false
    }
    try {
      parseFormula(stripLeadingEquals(formula))
      return true
    } catch {
      return false
    }
  }

  getRegisteredFunctionNames(languageCode?: string): string[] {
    const code = languageCode ?? this.config.language ?? 'enGB'
    const language = WorkPaper.languageRegistry.get(code)
    const functions = [...this.functionSnapshot.values()]
      .filter((binding) => binding.publicName === binding.publicName.toUpperCase())
      .map((binding) => binding.publicName)
      .toSorted(compareSheetNames)
    if (!language?.functions) {
      return functions
    }
    return functions.map((name) => language.functions?.[name] ?? name)
  }

  getFunctionPlugin(functionId: string): WorkPaperFunctionPluginDefinition | undefined {
    const binding = this.functionAliasLookup.get(functionId.trim().toUpperCase())
    if (!binding) {
      return undefined
    }
    const plugin = WorkPaper.functionPluginRegistry.get(binding.pluginId)
    return plugin ? clonePluginDefinition(plugin) : undefined
  }

  getAllFunctionPlugins(): WorkPaperFunctionPluginDefinition[] {
    const pluginIds = new Set([...this.functionSnapshot.values()].map((binding) => binding.pluginId))
    return [...pluginIds]
      .map((pluginId) => WorkPaper.functionPluginRegistry.get(pluginId))
      .filter((plugin): plugin is WorkPaperFunctionPluginDefinition => plugin !== undefined)
      .map((plugin) => clonePluginDefinition(plugin))
  }

  numberToDateTime(value: number): WorkPaperDateTime | undefined {
    const dateParts = excelSerialToDateParts(value)
    if (!dateParts) {
      return undefined
    }
    const whole = Math.floor(value)
    const fraction = value - whole
    const totalSeconds = Math.round(Math.max(0, fraction) * 86_400)
    const hours = Math.floor(totalSeconds / 3_600) % 24
    const minutes = Math.floor((totalSeconds % 3_600) / 60)
    const seconds = totalSeconds % 60
    return {
      year: dateParts.year,
      month: dateParts.month,
      day: dateParts.day,
      hours,
      minutes,
      seconds,
    }
  }

  numberToDate(value: number): Omit<WorkPaperDateTime, 'hours' | 'minutes' | 'seconds'> | undefined {
    const dateTime = this.numberToDateTime(value)
    if (!dateTime) {
      return undefined
    }
    const { year, month, day } = dateTime
    return { year, month, day }
  }

  numberToTime(value: number): Pick<WorkPaperDateTime, 'hours' | 'minutes' | 'seconds'> | undefined {
    const dateTime = this.numberToDateTime(value)
    if (!dateTime) {
      return undefined
    }
    const { hours, minutes, seconds } = dateTime
    return { hours, minutes, seconds }
  }

  setCellContents(address: WorkPaperCellAddress, content: RawCellContent | WorkPaperSheet): WorkPaperChange[] {
    this.assertNotDisposed()
    const sheet = this.sheetRecord(address.sheet)
    assertRowAndColumn(address.row, 'address.row')
    assertRowAndColumn(address.col, 'address.col')
    if (!isWorkPaperSheetMatrix(content)) {
      if (address.row >= (this.config.maxRows ?? MAX_ROWS) || address.col >= (this.config.maxColumns ?? MAX_COLS)) {
        throw new WorkPaperOperationError('Cell contents cannot be set')
      }
      const existingCellIndex = sheet.grid.get(address.row, address.col)
      if (this.enqueueSuspendedLiteralMutation(address.sheet, address.row, address.col, content, existingCellIndex)) {
        return []
      }
      if (this.enqueueDeferredBatchLiteral(address.sheet, address.row, address.col, content, existingCellIndex)) {
        return []
      }
      const mutate = () => {
        this.flushPendingBatchOps()
        const mutation: EngineCellMutationRef['mutation'] =
          content === null
            ? { kind: 'clearCell', row: address.row, col: address.col }
            : typeof content === 'string' && content.trim().startsWith('=')
              ? {
                  kind: 'setCellFormula',
                  row: address.row,
                  col: address.col,
                  formula: this.rewriteFormulaForStorage(stripLeadingEquals(content), address.sheet),
                }
              : {
                  kind: 'setCellValue',
                  row: address.row,
                  col: address.col,
                  value: content,
                }
        this.applyCellMutationRefs([{ sheetId: address.sheet, mutation }], {
          captureUndo: true,
          potentialNewCells: content === null || existingCellIndex !== -1 ? 0 : 1,
          source: 'local',
          returnUndoOps: false,
          reuseRefs: true,
        })
      }
      if (this.canUseTrackedMutationFastPath()) {
        return this.captureTrackedChangesWithoutVisibilityCache(mutate)
      }
      return this.captureChanges(undefined, () => {
        mutate()
      })
    }
    if (!this.isItPossibleToSetCellContents(address, content)) {
      throw new WorkPaperOperationError('Cell contents cannot be set')
    }
    return this.captureChanges(undefined, () => {
      if (isWorkPaperSheetMatrix(content)) {
        this.flushPendingBatchOps()
        this.applyMatrixContents(address, content)
        return
      }
    })
  }

  swapRowIndexes(sheetId: number, rowA: number, rowB: number): WorkPaperChange[]
  swapRowIndexes(sheetId: number, rowMappings: readonly WorkPaperAxisSwapMapping[]): WorkPaperChange[]
  swapRowIndexes(sheetId: number, rowAOrMappings: number | readonly WorkPaperAxisSwapMapping[], rowB?: number): WorkPaperChange[] {
    const mappings = this.normalizeAxisSwapMappings('row', rowAOrMappings, rowB)
    if (!this.isItPossibleToSwapRowIndexes(sheetId, mappings)) {
      throw new WorkPaperOperationError('Rows cannot be swapped')
    }
    return this.batch(() => {
      mappings.forEach(([rowA, mappedRowB]) => {
        if (rowA === mappedRowB) {
          return
        }
        if (rowA < mappedRowB) {
          this.moveRows(sheetId, rowA, 1, mappedRowB)
          this.moveRows(sheetId, mappedRowB - 1, 1, rowA)
        } else {
          this.moveRows(sheetId, rowA, 1, mappedRowB)
          this.moveRows(sheetId, mappedRowB + 1, 1, rowA)
        }
      })
    })
  }

  setRowOrder(sheetId: number, rowOrder: readonly number[]): WorkPaperChange[] {
    if (!this.isItPossibleToSetRowOrder(sheetId, rowOrder)) {
      throw new WorkPaperOperationError('Row order is invalid')
    }
    const current = rowOrder.toSorted((left, right) => left - right)
    return this.batch(() => {
      rowOrder.forEach((targetOriginalIndex, targetIndex) => {
        const currentIndex = current.indexOf(targetOriginalIndex)
        if (currentIndex === targetIndex) {
          return
        }
        this.moveRows(sheetId, currentIndex, 1, targetIndex)
        const [moved] = current.splice(currentIndex, 1)
        current.splice(targetIndex, 0, moved!)
      })
    })
  }

  swapColumnIndexes(sheetId: number, columnA: number, columnB: number): WorkPaperChange[]
  swapColumnIndexes(sheetId: number, columnMappings: readonly WorkPaperAxisSwapMapping[]): WorkPaperChange[]
  swapColumnIndexes(sheetId: number, columnAOrMappings: number | readonly WorkPaperAxisSwapMapping[], columnB?: number): WorkPaperChange[] {
    const mappings = this.normalizeAxisSwapMappings('column', columnAOrMappings, columnB)
    if (!this.isItPossibleToSwapColumnIndexes(sheetId, mappings)) {
      throw new WorkPaperOperationError('Columns cannot be swapped')
    }
    return this.batch(() => {
      mappings.forEach(([columnA, mappedColumnB]) => {
        if (columnA === mappedColumnB) {
          return
        }
        if (columnA < mappedColumnB) {
          this.moveColumns(sheetId, columnA, 1, mappedColumnB)
          this.moveColumns(sheetId, mappedColumnB - 1, 1, columnA)
        } else {
          this.moveColumns(sheetId, columnA, 1, mappedColumnB)
          this.moveColumns(sheetId, mappedColumnB + 1, 1, columnA)
        }
      })
    })
  }

  setColumnOrder(sheetId: number, columnOrder: readonly number[]): WorkPaperChange[] {
    if (!this.isItPossibleToSetColumnOrder(sheetId, columnOrder)) {
      throw new WorkPaperOperationError('Column order is invalid')
    }
    const current = columnOrder.toSorted((left, right) => left - right)
    return this.batch(() => {
      columnOrder.forEach((targetOriginalIndex, targetIndex) => {
        const currentIndex = current.indexOf(targetOriginalIndex)
        if (currentIndex === targetIndex) {
          return
        }
        this.moveColumns(sheetId, currentIndex, 1, targetIndex)
        const [moved] = current.splice(currentIndex, 1)
        current.splice(targetIndex, 0, moved!)
      })
    })
  }

  addRows(sheetId: number, start: number, count?: number): WorkPaperChange[]
  addRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  addRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    const indexes = this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals)
    if (!this.isItPossibleToAddRows(sheetId, ...indexes)) {
      throw new WorkPaperOperationError('Rows cannot be added')
    }
    return this.batchStructuralChanges(() => {
      indexes.forEach(([start, amount]) => {
        this.engine.insertRows(this.sheetName(sheetId), start, amount)
      })
    })
  }

  removeRows(sheetId: number, start: number, count?: number): WorkPaperChange[]
  removeRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  removeRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    const indexes = this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals)
    if (!this.isItPossibleToRemoveRows(sheetId, ...indexes)) {
      throw new WorkPaperOperationError('Rows cannot be removed')
    }
    return this.batchStructuralChanges(() => {
      indexes
        .toSorted((left, right) => right[0] - left[0])
        .forEach(([start, amount]) => {
          this.engine.deleteRows(this.sheetName(sheetId), start, amount)
        })
    })
  }

  addColumns(sheetId: number, start: number, count?: number): WorkPaperChange[]
  addColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  addColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    const indexes = this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals)
    if (!this.isItPossibleToAddColumns(sheetId, ...indexes)) {
      throw new WorkPaperOperationError('Columns cannot be added')
    }
    return this.batchStructuralChanges(() => {
      indexes.forEach(([start, amount]) => {
        this.engine.insertColumns(this.sheetName(sheetId), start, amount)
      })
    })
  }

  removeColumns(sheetId: number, start: number, count?: number): WorkPaperChange[]
  removeColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  removeColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    const indexes = this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals)
    if (!this.isItPossibleToRemoveColumns(sheetId, ...indexes)) {
      throw new WorkPaperOperationError('Columns cannot be removed')
    }
    return this.batchStructuralChanges(() => {
      indexes
        .toSorted((left, right) => right[0] - left[0])
        .forEach(([start, amount]) => {
          this.engine.deleteColumns(this.sheetName(sheetId), start, amount)
        })
    })
  }

  moveCells(source: WorkPaperCellRange, target: WorkPaperCellAddress): WorkPaperChange[] {
    if (!this.isItPossibleToMoveCells(source, target)) {
      throw new WorkPaperOperationError('Cells cannot be moved')
    }
    const sourceHeight = source.end.row - source.start.row
    const sourceWidth = source.end.col - source.start.col
    return this.captureChanges(undefined, () => {
      this.engine.moveRange(sourceRangeRef(this.sheetName(source.start.sheet), source), {
        sheetName: this.sheetName(target.sheet),
        startAddress: formatAddress(target.row, target.col),
        endAddress: formatAddress(target.row + sourceHeight, target.col + sourceWidth),
      })
    })
  }

  moveRows(sheetId: number, start: number, count: number, target: number): WorkPaperChange[] {
    if (!this.isItPossibleToMoveRows(sheetId, start, count, target)) {
      throw new WorkPaperOperationError('Rows cannot be moved')
    }
    return this.canUseTrackedStructuralFastPath()
      ? this.batchStructuralChanges(() => {
          this.engine.moveRows(this.sheetName(sheetId), start, count, target)
        })
      : this.captureChanges(undefined, () => {
          this.engine.moveRows(this.sheetName(sheetId), start, count, target)
        })
  }

  moveColumns(sheetId: number, start: number, count: number, target: number): WorkPaperChange[] {
    if (!this.isItPossibleToMoveColumns(sheetId, start, count, target)) {
      throw new WorkPaperOperationError('Columns cannot be moved')
    }
    return this.canUseTrackedStructuralFastPath()
      ? this.batchStructuralChanges(() => {
          this.engine.moveColumns(this.sheetName(sheetId), start, count, target)
        })
      : this.captureChanges(undefined, () => {
          this.engine.moveColumns(this.sheetName(sheetId), start, count, target)
        })
  }

  addSheet(sheetName?: string): string {
    this.assertNotDisposed()
    const name = sheetName?.trim() || this.nextSheetName()
    if (!this.isItPossibleToAddSheet(name)) {
      throw new WorkPaperSheetNameAlreadyTakenError(name)
    }
    const beforeVisibility = this.ensureVisibilityCache()
    const beforeNames = this.ensureNamedExpressionValueCache()
    this.drainTrackedEngineEvents()
    this.engine.createSheet(name)
    this.sheetRecordsCache = null
    const sheetId = this.requireSheetId(name)
    const payload: WorkPaperDetailedEventMap['sheetAdded'] = { sheetId, sheetName: name }
    if (this.shouldSuppressEvents()) {
      this.queuedEvents.push({ eventName: 'sheetAdded', payload })
    } else {
      this.emitter.emitDetailed({ eventName: 'sheetAdded', payload })
    }
    const changes = this.computeChangesAfterMutation(beforeVisibility, beforeNames)
    if (!this.shouldSuppressEvents() && changes.length > 0) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return name
  }

  removeSheet(sheetId: number): WorkPaperChange[] {
    if (!this.isItPossibleToRemoveSheet(sheetId)) {
      throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be removed`)
    }
    const sheetName = this.sheetName(sheetId)
    return this.captureChanges(
      {
        eventName: 'sheetRemoved',
        payload: {
          sheetId,
          sheetName,
          changes: [],
        },
      },
      () => {
        this.engine.deleteSheet(sheetName)
        this.sheetRecordsCache = null
      },
    )
  }

  clearSheet(sheetId: number): WorkPaperChange[] {
    if (!this.isItPossibleToClearSheet(sheetId)) {
      throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be cleared`)
    }
    return this.captureChanges(undefined, () => {
      const dimensions = this.getSheetDimensions(sheetId)
      if (dimensions.width === 0 || dimensions.height === 0) {
        return
      }
      this.engine.clearRange({
        sheetName: this.sheetName(sheetId),
        startAddress: 'A1',
        endAddress: formatAddress(dimensions.height - 1, dimensions.width - 1),
      })
    })
  }

  setSheetContent(sheetId: number, content: WorkPaperSheet): WorkPaperChange[] {
    if (!this.isItPossibleToReplaceSheetContent(sheetId, content)) {
      throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be replaced`)
    }
    return this.captureChanges(undefined, () => {
      this.replaceSheetContentInternal(sheetId, content, { duringInitialization: false })
    })
  }

  renameSheet(sheetId: number, nextName: string): WorkPaperChange[] {
    if (!this.isItPossibleToRenameSheet(sheetId, nextName)) {
      throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be renamed to '${nextName}'`)
    }
    const oldName = this.sheetName(sheetId)
    const newName = nextName.trim()
    return this.captureChanges(
      {
        eventName: 'sheetRenamed',
        payload: {
          sheetId,
          oldName,
          newName,
        },
      },
      () => {
        this.engine.renameSheet(oldName, newName)
        this.sheetRecordsCache = null
      },
    )
  }

  isItPossibleToSetCellContents(address: WorkPaperCellAddress, content?: RawCellContent | WorkPaperSheet): boolean
  isItPossibleToSetCellContents(range: WorkPaperCellRange): boolean
  isItPossibleToSetCellContents(addressOrRange: WorkPaperAddressLike, content?: RawCellContent | WorkPaperSheet): boolean {
    this.assertNotDisposed()
    if (isCellRange(addressOrRange)) {
      assertRange(addressOrRange)
      this.sheetRecord(addressOrRange.start.sheet)
      return addressOrRange.end.row < (this.config.maxRows ?? MAX_ROWS) && addressOrRange.end.col < (this.config.maxColumns ?? MAX_COLS)
    }
    this.sheetRecord(addressOrRange.sheet)
    assertRowAndColumn(addressOrRange.row, 'address.row')
    assertRowAndColumn(addressOrRange.col, 'address.col')
    if (content === undefined) {
      return addressOrRange.row < (this.config.maxRows ?? MAX_ROWS) && addressOrRange.col < (this.config.maxColumns ?? MAX_COLS)
    }
    if (Array.isArray(content)) {
      if (!content.every((row) => Array.isArray(row))) {
        throw new WorkPaperInvalidArgumentsError('Content matrix must be a two-dimensional array')
      }
      const height = content.length
      const width = Math.max(0, ...content.map((row) => row.length))
      return (
        addressOrRange.row + height <= (this.config.maxRows ?? MAX_ROWS) &&
        addressOrRange.col + width <= (this.config.maxColumns ?? MAX_COLS)
      )
    }
    return addressOrRange.row < (this.config.maxRows ?? MAX_ROWS) && addressOrRange.col < (this.config.maxColumns ?? MAX_COLS)
  }

  isItPossibleToSwapRowIndexes(sheetId: number, rowA: number, rowB: number): boolean
  isItPossibleToSwapRowIndexes(sheetId: number, rowMappings: readonly WorkPaperAxisSwapMapping[]): boolean
  isItPossibleToSwapRowIndexes(sheetId: number, rowAOrMappings: number | readonly WorkPaperAxisSwapMapping[], rowB?: number): boolean {
    this.sheetRecord(sheetId)
    const mappings = this.normalizeAxisSwapMappings('row', rowAOrMappings, rowB)
    return mappings.every(([rowA, mappedRowB]) => {
      assertRowAndColumn(rowA, 'rowA')
      assertRowAndColumn(mappedRowB, 'rowB')
      return true
    })
  }

  isItPossibleToSetRowOrder(sheetId: number, rowOrder: readonly number[]): boolean {
    this.sheetRecord(sheetId)
    if (new Set(rowOrder).size !== rowOrder.length || rowOrder.some((value) => !Number.isInteger(value) || value < 0)) {
      return false
    }
    return true
  }

  isItPossibleToSwapColumnIndexes(sheetId: number, columnA: number, columnB: number): boolean
  isItPossibleToSwapColumnIndexes(sheetId: number, columnMappings: readonly WorkPaperAxisSwapMapping[]): boolean
  isItPossibleToSwapColumnIndexes(
    sheetId: number,
    columnAOrMappings: number | readonly WorkPaperAxisSwapMapping[],
    columnB?: number,
  ): boolean {
    this.sheetRecord(sheetId)
    const mappings = this.normalizeAxisSwapMappings('column', columnAOrMappings, columnB)
    return mappings.every(([columnA, mappedColumnB]) => {
      assertRowAndColumn(columnA, 'columnA')
      assertRowAndColumn(mappedColumnB, 'columnB')
      return true
    })
  }

  isItPossibleToSetColumnOrder(sheetId: number, columnOrder: readonly number[]): boolean {
    this.sheetRecord(sheetId)
    if (new Set(columnOrder).size !== columnOrder.length || columnOrder.some((value) => !Number.isInteger(value) || value < 0)) {
      return false
    }
    return true
  }

  isItPossibleToAddRows(sheetId: number, start: number, count?: number): boolean
  isItPossibleToAddRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): boolean
  isItPossibleToAddRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): boolean {
    this.sheetRecord(sheetId)
    return this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(([start, count]) => {
      assertRowAndColumn(start, 'start')
      assertRowAndColumn(count, 'count')
      return count > 0 && start + count <= (this.config.maxRows ?? MAX_ROWS)
    })
  }

  isItPossibleToRemoveRows(sheetId: number, start: number, count?: number): boolean
  isItPossibleToRemoveRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): boolean
  isItPossibleToRemoveRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): boolean {
    this.sheetRecord(sheetId)
    return this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(([start, count]) => {
      assertRowAndColumn(start, 'start')
      assertRowAndColumn(count, 'count')
      return count > 0
    })
  }

  isItPossibleToAddColumns(sheetId: number, start: number, count?: number): boolean
  isItPossibleToAddColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): boolean
  isItPossibleToAddColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): boolean {
    this.sheetRecord(sheetId)
    return this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(([start, count]) => {
      assertRowAndColumn(start, 'start')
      assertRowAndColumn(count, 'count')
      return count > 0 && start + count <= (this.config.maxColumns ?? MAX_COLS)
    })
  }

  isItPossibleToRemoveColumns(sheetId: number, start: number, count?: number): boolean
  isItPossibleToRemoveColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): boolean
  isItPossibleToRemoveColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): boolean {
    this.sheetRecord(sheetId)
    return this.normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(([start, count]) => {
      assertRowAndColumn(start, 'start')
      assertRowAndColumn(count, 'count')
      return count > 0
    })
  }

  isItPossibleToMoveCells(source: WorkPaperCellRange, target: WorkPaperCellAddress): boolean {
    assertRange(source)
    assertRowAndColumn(target.sheet, 'target.sheet')
    assertRowAndColumn(target.row, 'target.row')
    assertRowAndColumn(target.col, 'target.col')
    return source.start.sheet === target.sheet
  }

  isItPossibleToMoveRows(sheetId: number, start: number, count: number, target: number): boolean {
    this.sheetRecord(sheetId)
    assertRowAndColumn(start, 'start')
    assertRowAndColumn(count, 'count')
    assertRowAndColumn(target, 'target')
    return count > 0
  }

  isItPossibleToMoveColumns(sheetId: number, start: number, count: number, target: number): boolean {
    this.sheetRecord(sheetId)
    assertRowAndColumn(start, 'start')
    assertRowAndColumn(count, 'count')
    assertRowAndColumn(target, 'target')
    return count > 0
  }

  isItPossibleToAddSheet(sheetName: string): boolean {
    const trimmed = sheetName.trim()
    if (trimmed.length === 0) {
      throw new WorkPaperInvalidArgumentsError('Sheet name must be non-empty')
    }
    return !this.doesSheetExist(trimmed)
  }

  isItPossibleToRemoveSheet(sheetId: number): boolean {
    return this.engine.workbook.getSheetById(sheetId) !== undefined
  }

  isItPossibleToClearSheet(sheetId: number): boolean {
    return this.engine.workbook.getSheetById(sheetId) !== undefined
  }

  isItPossibleToReplaceSheetContent(sheetId: number, content: WorkPaperSheet): boolean {
    this.sheetRecord(sheetId)
    if (!content.every((row) => Array.isArray(row))) {
      throw new WorkPaperInvalidArgumentsError('Sheet content must be a two-dimensional array')
    }
    const height = content.length
    const width = Math.max(0, ...content.map((row) => row.length))
    return height <= (this.config.maxRows ?? MAX_ROWS) && width <= (this.config.maxColumns ?? MAX_COLS)
  }

  isItPossibleToRenameSheet(sheetId: number, nextName: string): boolean {
    this.sheetRecord(sheetId)
    const trimmed = nextName.trim()
    if (trimmed.length === 0) {
      throw new WorkPaperInvalidArgumentsError('Sheet name must be non-empty')
    }
    const existing = this.engine.workbook.getSheet(trimmed)
    return !existing || existing.id === sheetId
  }

  isItPossibleToAddNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): boolean {
    this.validateNamedExpression(expressionName, expression, scope)
    return !this.namedExpressions.has(makeNamedExpressionKey(expressionName, scope))
  }

  isItPossibleToChangeNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): boolean {
    this.validateNamedExpression(expressionName, expression, scope)
    return this.namedExpressions.has(makeNamedExpressionKey(expressionName, scope))
  }

  isItPossibleToRemoveNamedExpression(expressionName: string, scope?: number): boolean {
    return this.namedExpressions.has(makeNamedExpressionKey(expressionName, scope))
  }

  destroy(): void {
    this.dispose()
  }

  dispose(): void {
    if (this.disposed) {
      return
    }
    this.disposed = true
    this.unsubscribeEngineEvents?.()
    this.unsubscribeEngineEvents = null
    this.emitter.clear()
    this.clearFunctionBindings()
    this.clipboard = null
    this.visibilityCache = null
    this.namedExpressionValueCache = null
    this.queuedEvents = []
    this.trackedEngineEvents = []
    this.namedExpressions.clear()
  }

  private attachEngineEventTracking(): void {
    this.unsubscribeEngineEvents?.()
    this.trackedEngineEvents = []
    if (!hasTrackedEngineSubscription(this.engine.events)) {
      throw new WorkPaperOperationError('Tracked engine event subscription is unavailable')
    }
    this.unsubscribeEngineEvents = this.engine.events.subscribeTracked((event) => {
      if (!this.engineEventCaptureEnabled) {
        return
      }
      this.trackedEngineEvents.push(captureTrackedEngineEvent(event))
    })
  }

  private withEngineEventCaptureDisabled<T>(callback: () => T): T {
    const previous = this.engineEventCaptureEnabled
    this.engineEventCaptureEnabled = false
    this.trackedEngineEvents = []
    try {
      return callback()
    } finally {
      this.engineEventCaptureEnabled = previous
      this.trackedEngineEvents = []
    }
  }

  private drainTrackedEngineEvents(): TrackedEngineEvent[] {
    const events = this.trackedEngineEvents
    this.trackedEngineEvents = []
    return events
  }

  private resetChangeTrackingCaches(): void {
    this.sheetRecordsCache = null
    this.visibilityCache = null
    this.namedExpressionValueCache = null
    this.drainTrackedEngineEvents()
  }

  private ensureVisibilityCache(): VisibilitySnapshot {
    if (!this.visibilityCache) {
      this.visibilityCache = this.captureVisibilitySnapshot()
    }
    return this.visibilityCache
  }

  private ensureNamedExpressionValueCache(): NamedExpressionValueSnapshot {
    if (!this.namedExpressionValueCache) {
      this.namedExpressionValueCache =
        this.namedExpressions.size > 0 ? this.captureNamedExpressionValueSnapshot() : EMPTY_NAMED_EXPRESSION_VALUES
    }
    return this.namedExpressionValueCache
  }

  private flushPendingBatchOps(): void {
    if (this.pendingBatchOps.length === 0) {
      return
    }
    const ops = this.pendingBatchOps
    const potentialNewCells = this.pendingBatchPotentialNewCells
    this.pendingBatchOps = []
    this.pendingBatchPotentialNewCells = 0
    this.engine.applyCellMutationsAtWithOptions(ops, {
      captureUndo: true,
      potentialNewCells: potentialNewCells > 0 ? potentialNewCells : undefined,
      source: 'local',
      returnUndoOps: false,
      reuseRefs: true,
    })
  }

  private applyCellMutationRefs(
    refs: readonly EngineCellMutationRef[],
    options: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      reuseRefs?: boolean
    },
  ): void {
    if (this.evaluationSuspended && (options.source ?? 'local') === 'local') {
      for (let index = 0; index < refs.length; index += 1) {
        const ref = refs[index]
        if (!ref) {
          continue
        }
        const mutation = ref.mutation
        this.suspendedCellMutationRefs.push({
          sheetId: ref.sheetId,
          mutation:
            mutation.kind === 'setCellValue'
              ? {
                  kind: 'setCellValue',
                  row: mutation.row,
                  col: mutation.col,
                  value: mutation.value,
                }
              : mutation.kind === 'setCellFormula'
                ? {
                    kind: 'setCellFormula',
                    row: mutation.row,
                    col: mutation.col,
                    formula: mutation.formula,
                  }
                : {
                    kind: 'clearCell',
                    row: mutation.row,
                    col: mutation.col,
                  },
        })
      }
      this.suspendedCellMutationPotentialNewCells +=
        options.potentialNewCells ?? refs.reduce((count, ref) => (ref?.mutation.kind === 'clearCell' ? count : count + 1), 0)
      return
    }
    this.engine.applyCellMutationsAtWithOptions(refs, options)
  }

  private flushSuspendedCellMutations(): void {
    if (this.suspendedCellMutationRefs.length === 0) {
      return
    }
    const refs = this.suspendedCellMutationRefs
    const potentialNewCells = this.suspendedCellMutationPotentialNewCells
    this.suspendedCellMutationRefs = []
    this.suspendedCellMutationPotentialNewCells = 0
    this.engine.applyCellMutationsAtWithOptions(refs, {
      captureUndo: true,
      potentialNewCells: potentialNewCells > 0 ? potentialNewCells : undefined,
      source: 'local',
      returnUndoOps: false,
      reuseRefs: true,
    })
  }

  private enqueueSuspendedLiteralMutation(
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    existingCellIndex = this.engine.workbook.getSheetById(sheetId)?.grid.get(row, col) ?? -1,
  ): boolean {
    if (!this.evaluationSuspended || !isDeferredBatchLiteralContent(content) || isFormulaContent(content)) {
      return false
    }
    if (content === null) {
      this.suspendedCellMutationRefs.push({ sheetId, mutation: { kind: 'clearCell', row, col } })
      return true
    }
    this.suspendedCellMutationRefs.push({
      sheetId,
      mutation: { kind: 'setCellValue', row, col, value: content },
    })
    if (existingCellIndex === -1) {
      this.suspendedCellMutationPotentialNewCells += 1
    }
    return true
  }

  private enqueueDeferredBatchLiteral(
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    existingCellIndex = this.engine.workbook.getSheetById(sheetId)?.grid.get(row, col) ?? -1,
  ): boolean {
    if (this.batchDepth === 0 || this.evaluationSuspended || !isDeferredBatchLiteralContent(content) || isFormulaContent(content)) {
      return false
    }
    if (content === null) {
      this.pendingBatchOps.push({ sheetId, mutation: { kind: 'clearCell', row, col } })
      return true
    }
    this.pendingBatchOps.push({
      sheetId,
      mutation: { kind: 'setCellValue', row, col, value: content },
    })
    if (existingCellIndex === -1) {
      this.pendingBatchPotentialNewCells += 1
    }
    return true
  }

  private prepareReadableState(): void {
    this.assertNotDisposed()
    this.flushPendingBatchOps()
  }

  private assertNotDisposed(): void {
    if (this.disposed) {
      throw new WorkPaperOperationError('Workbook has been disposed')
    }
  }

  private assertReadable(): void {
    this.prepareReadableState()
    if (this.evaluationSuspended) {
      throw new WorkPaperEvaluationSuspendedError()
    }
  }

  private sheetRecord(sheetId: number) {
    const sheet = this.engine.workbook.getSheetById(sheetId)
    if (!sheet) {
      throw new WorkPaperNoSheetWithIdError(sheetId)
    }
    return sheet
  }

  private sheetName(sheetId: number): string {
    return this.sheetRecord(sheetId).name
  }

  private requireSheetId(name: string): number {
    const sheetId = this.getSheetId(name)
    if (sheetId === undefined) {
      throw new WorkPaperNoSheetWithNameError(name)
    }
    return sheetId
  }

  private a1(address: Pick<WorkPaperCellAddress, 'row' | 'col'>): string {
    return formatAddress(address.row, address.col)
  }

  private rangeRef(range: WorkPaperCellRange): CellRangeRef {
    assertRange(range)
    return sourceRangeRef(this.sheetName(range.start.sheet), range)
  }

  private getDirectPrecedentStrings(address: WorkPaperCellAddress): string[] {
    const precedents = new Set<string>(this.engine.getDependencies(this.sheetName(address.sheet), this.a1(address)).directPrecedents)
    const formula = this.getCellFormula(address)
    if (formula) {
      this.getNamedExpressionsFromFormula(formula).forEach((name) => {
        precedents.add(name)
      })
    }
    return [...precedents]
  }

  private getDirectPrecedentRefs(address: WorkPaperCellAddress): WorkPaperDependencyRef[] {
    return this.toDependencyRefs(this.getDirectPrecedentStrings(address))
  }

  private listSheetRecords() {
    if (this.sheetRecordsCache) {
      return this.sheetRecordsCache
    }
    this.sheetRecordsCache = [...this.engine.workbook.sheetsByName.values()].toSorted(
      (left, right) => left.order - right.order || left.name.localeCompare(right.name),
    )
    return this.sheetRecordsCache
  }

  private getDenseRange<Value>(range: WorkPaperCellRange, read: (address: WorkPaperCellAddress) => Value): Value[][] {
    assertRange(range)
    const height = range.end.row - range.start.row + 1
    const width = range.end.col - range.start.col + 1
    return Array.from({ length: height }, (_row, rowOffset) =>
      Array.from({ length: width }, (_column, colOffset) =>
        read({
          sheet: range.start.sheet,
          row: range.start.row + rowOffset,
          col: range.start.col + colOffset,
        }),
      ),
    )
  }

  private captureVisibilitySnapshot(): VisibilitySnapshot {
    const snapshot = new Map<number, SheetStateSnapshot>()
    const strings = this.engine.strings
    const cellStore = this.engine.workbook.cellStore
    this.listSheetRecords().forEach((sheet) => {
      const cells = new Map<number, CellValue>()
      sheet.grid.forEachCellEntry((cellIndex: number, row: number, col: number) => {
        const value = cellStore.getValue(cellIndex, (id) => strings.get(id))
        if (value.tag === ValueTag.Empty) {
          return
        }
        cells.set(makeCellKey(sheet.id, row, col), value)
      })
      snapshot.set(sheet.id, {
        sheetId: sheet.id,
        sheetName: sheet.name,
        order: sheet.order,
        cells,
      })
    })
    return snapshot
  }

  private captureNamedExpressionValueSnapshot(): NamedExpressionValueSnapshot {
    if (this.namedExpressions.size === 0) {
      return EMPTY_NAMED_EXPRESSION_VALUES
    }
    const snapshot = new Map<string, CellValue | CellValue[][]>()
    ;[...this.namedExpressions.values()].forEach((expression) => {
      snapshot.set(
        makeNamedExpressionKey(expression.publicName, expression.scope),
        cloneNamedExpressionValue(this.evaluateNamedExpression(expression)),
      )
    })
    return snapshot
  }

  private computeCellChanges(beforeVisibility: VisibilitySnapshot, afterVisibility: VisibilitySnapshot): WorkPaperChange[] {
    const cellChanges: WorkPaperChange[] = []
    afterVisibility.forEach((afterSheet, sheetId) => {
      const beforeSheet = beforeVisibility.get(sheetId)
      const cellKeys = new Set<number>([...(beforeSheet?.cells.keys() ?? []), ...afterSheet.cells.keys()])
      ;[...cellKeys]
        .toSorted((left, right) => left - right)
        .forEach((cellKey) => {
          const beforeValue = beforeSheet?.cells.get(cellKey) ?? emptyValue()
          const afterValue = afterSheet.cells.get(cellKey) ?? emptyValue()
          if (valuesEqual(beforeValue, afterValue)) {
            return
          }
          const localKey = cellKey - afterSheet.sheetId * VISIBILITY_SHEET_STRIDE
          const row = Math.floor(localKey / MAX_COLS)
          const col = localKey % MAX_COLS
          const address = formatAddress(row, col)
          cellChanges.push({
            kind: 'cell',
            address: { sheet: sheetId, row, col },
            sheetName: afterSheet.sheetName,
            a1: address,
            newValue: afterValue,
          })
        })
    })
    return orderWorkPaperCellChanges(cellChanges, this.listSheetRecords())
  }

  private readTrackedCellChange(cellIndex: number): WorkPaperCellChange | undefined {
    const sheetId = this.engine.workbook.cellStore.sheetIds[cellIndex]
    const row = this.engine.workbook.cellStore.rows[cellIndex]
    const col = this.engine.workbook.cellStore.cols[cellIndex]
    if (sheetId === undefined || row === undefined || col === undefined) {
      return undefined
    }
    const sheetName = this.engine.workbook.getSheetNameById(sheetId)
    if (sheetName === undefined) {
      return undefined
    }
    return {
      kind: 'cell',
      address: { sheet: sheetId, row, col },
      sheetName,
      a1: formatAddress(row, col),
      newValue: this.engine.workbook.cellStore.getValue(cellIndex, (id) => this.engine.strings.get(id)),
    }
  }

  private computeCellChangesFromTrackedEvents(
    beforeVisibility: VisibilitySnapshot,
    events: readonly TrackedEngineEvent[],
  ): { changes: WorkPaperChange[]; nextVisibility: VisibilitySnapshot } | null {
    if (events.some((event) => event.invalidation === 'full')) {
      return null
    }

    const nextVisibility = beforeVisibility
    const ensureMutableSheet = (sheetId: number, sheetName: string): SheetStateSnapshot => {
      const existing = nextVisibility.get(sheetId)
      if (existing) {
        return existing
      }
      const created: SheetStateSnapshot = {
        sheetId,
        sheetName,
        order: this.sheetRecord(sheetId).order,
        cells: new Map<number, CellValue>(),
      }
      nextVisibility.set(sheetId, created)
      return created
    }
    if (events.length === 1) {
      const event = events[0]!
      let hasDuplicateCellKey = false
      if (event.changedCellIndices.length <= 4) {
        for (let index = 0; index < event.changedCellIndices.length; index += 1) {
          const change = this.readTrackedCellChange(event.changedCellIndices[index]!)
          if (!change) {
            continue
          }
          const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
          for (let priorIndex = 0; priorIndex < index; priorIndex += 1) {
            const prior = this.readTrackedCellChange(event.changedCellIndices[priorIndex]!)
            if (!prior) {
              continue
            }
            if (makeCellKey(prior.address.sheet, prior.address.row, prior.address.col) === cellKey) {
              hasDuplicateCellKey = true
              break
            }
          }
          if (hasDuplicateCellKey) {
            break
          }
        }
      } else {
        const seenCellKeys = new Set<number>()
        for (let index = 0; index < event.changedCellIndices.length; index += 1) {
          const change = this.readTrackedCellChange(event.changedCellIndices[index]!)
          if (!change) {
            continue
          }
          const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
          if (seenCellKeys.has(cellKey)) {
            hasDuplicateCellKey = true
            break
          }
          seenCellKeys.add(cellKey)
        }
      }
      if (!hasDuplicateCellKey) {
        const directChanges: WorkPaperCellChange[] = []
        let alreadySorted = true
        let previousSheetOrder = -1
        let previousRow = -1
        let previousCol = -1
        for (let index = 0; index < event.changedCellIndices.length; index += 1) {
          const change = this.readTrackedCellChange(event.changedCellIndices[index]!)
          if (!change) {
            continue
          }
          const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
          const sheet = ensureMutableSheet(change.address.sheet, change.sheetName)
          if (
            sheet.order < previousSheetOrder ||
            (sheet.order === previousSheetOrder &&
              (change.address.row < previousRow || (change.address.row === previousRow && change.address.col < previousCol)))
          ) {
            alreadySorted = false
          }
          if (change.newValue.tag === ValueTag.Empty) {
            sheet.cells.delete(cellKey)
          } else {
            sheet.cells.set(cellKey, change.newValue)
          }
          directChanges[index] = {
            kind: 'cell',
            address: change.address,
            sheetName: change.sheetName,
            a1: change.a1,
            newValue: change.newValue,
          }
          previousSheetOrder = sheet.order
          previousRow = change.address.row
          previousCol = change.address.col
        }
        return {
          changes: alreadySorted
            ? directChanges
            : orderWorkPaperCellChanges(directChanges, this.listSheetRecords(), event.explicitChangedCount),
          nextVisibility,
        }
      }
    }
    const latestChangesByKey = new Map<number, WorkPaperCellChange>()
    for (const event of events) {
      for (let index = 0; index < event.changedCellIndices.length; index += 1) {
        const change = this.readTrackedCellChange(event.changedCellIndices[index]!)
        if (!change) {
          continue
        }
        const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
        latestChangesByKey.delete(cellKey)
        latestChangesByKey.set(cellKey, {
          kind: 'cell',
          address: change.address,
          sheetName: change.sheetName,
          a1: change.a1,
          newValue: change.newValue,
        })
      }
    }
    for (const change of latestChangesByKey.values()) {
      const cellKey = makeCellKey(change.address.sheet, change.address.row, change.address.col)
      const sheet = ensureMutableSheet(change.address.sheet, change.sheetName)
      if (change.newValue.tag === ValueTag.Empty) {
        sheet.cells.delete(cellKey)
      } else {
        sheet.cells.set(cellKey, change.newValue)
      }
    }
    const directChanges = [...latestChangesByKey.values()]
    return {
      changes: orderWorkPaperCellChanges(directChanges, this.listSheetRecords()),
      nextVisibility,
    }
  }

  private computeNamedExpressionChanges(
    beforeNames: NamedExpressionValueSnapshot,
    afterNames: NamedExpressionValueSnapshot,
  ): WorkPaperChange[] {
    const namedExpressionChanges: WorkPaperChange[] = []
    afterNames.forEach((afterValue, key) => {
      const beforeValue = beforeNames.get(key)
      if (matrixValuesEqual(beforeValue, afterValue)) {
        return
      }
      const expression = this.namedExpressions.get(key)
      if (!expression) {
        return
      }
      namedExpressionChanges.push({
        kind: 'named-expression',
        name: expression.publicName,
        scope: expression.scope,
        newValue: cloneNamedExpressionValue(afterValue),
      })
    })
    return namedExpressionChanges.toSorted(compareWorkPaperNamedExpressionChanges)
  }

  private canUseTrackedStructuralFastPath(): boolean {
    return this.batchDepth === 0 && !this.evaluationSuspended && this.visibilityCache === null && this.namedExpressions.size === 0
  }

  private canUseTrackedMutationFastPath(): boolean {
    return this.batchDepth === 0 && !this.evaluationSuspended && this.visibilityCache === null && this.namedExpressions.size === 0
  }

  private downgradeTrackedBatchFastPath(): void {
    if (!this.batchUsesTrackedFastPath || this.batchDepth === 0) {
      return
    }
    this.batchStartVisibility = this.ensureVisibilityCache()
    this.batchStartNamedValues = this.namedExpressions.size > 0 ? this.ensureNamedExpressionValueCache() : EMPTY_NAMED_EXPRESSION_VALUES
    this.batchUsesTrackedFastPath = false
  }

  private computeTrackedChangesWithoutVisibilityCache(events: readonly TrackedEngineEvent[]): WorkPaperChange[] {
    const fastPath = this.computeCellChangesFromTrackedEvents(new Map(), events)
    if (!fastPath) {
      throw new WorkPaperOperationError('Mutation emitted an unsupported invalidation pattern for tracked changes')
    }
    return fastPath.changes
  }

  private captureTrackedChangesWithoutVisibilityCache(mutate: () => void): WorkPaperChange[] {
    this.assertNotDisposed()
    this.drainTrackedEngineEvents()
    try {
      mutate()
    } catch (error) {
      if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
        throw error
      }
      throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
    }
    const changes = this.computeTrackedChangesWithoutVisibilityCache(this.drainTrackedEngineEvents())
    if (changes.length > 0) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  private batchStructuralChanges(batchOperations: () => void): WorkPaperChange[] {
    if (!this.canUseTrackedStructuralFastPath()) {
      this.downgradeTrackedBatchFastPath()
      return this.batch(batchOperations)
    }
    this.assertNotDisposed()
    const undoStackStart = this.getUndoStack().length
    this.drainTrackedEngineEvents()
    this.batchDepth += 1
    try {
      batchOperations()
    } finally {
      this.batchDepth -= 1
      this.flushPendingBatchOps()
      this.mergeUndoHistory(undoStackStart)
    }
    const changes = this.computeTrackedChangesWithoutVisibilityCache(this.drainTrackedEngineEvents())
    this.flushQueuedEvents()
    if (changes.length > 0) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  private computeChangesAfterMutation(beforeVisibility: VisibilitySnapshot, beforeNames: NamedExpressionValueSnapshot): WorkPaperChange[] {
    const hasNamedExpressions = this.namedExpressions.size > 0
    const afterNames = hasNamedExpressions ? this.captureNamedExpressionValueSnapshot() : EMPTY_NAMED_EXPRESSION_VALUES
    const fastPath = this.computeCellChangesFromTrackedEvents(beforeVisibility, this.drainTrackedEngineEvents())
    let cellChanges: WorkPaperChange[]
    if (fastPath) {
      cellChanges = fastPath.changes
      this.visibilityCache = fastPath.nextVisibility
    } else {
      const afterVisibility = this.captureVisibilitySnapshot()
      cellChanges = this.computeCellChanges(beforeVisibility, afterVisibility)
      this.visibilityCache = afterVisibility
    }
    this.namedExpressionValueCache = afterNames
    return hasNamedExpressions ? [...cellChanges, ...this.computeNamedExpressionChanges(beforeNames, afterNames)] : cellChanges
  }

  private captureChanges(semanticEvent: QueuedEvent | undefined, mutate: () => void): WorkPaperChange[] {
    this.assertNotDisposed()
    this.downgradeTrackedBatchFastPath()
    if (semanticEvent !== undefined) {
      this.flushPendingBatchOps()
    }
    if (this.shouldSuppressEvents()) {
      try {
        mutate()
      } catch (error) {
        if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
          throw error
        }
        throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
      }
      if (semanticEvent) {
        this.queuedEvents.push(semanticEvent)
      }
      return []
    }
    const beforeVisibility = this.ensureVisibilityCache()
    const beforeNames = this.namedExpressions.size > 0 ? this.ensureNamedExpressionValueCache() : EMPTY_NAMED_EXPRESSION_VALUES
    this.drainTrackedEngineEvents()
    try {
      mutate()
    } catch (error) {
      if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
        throw error
      }
      throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
    }
    const changes =
      semanticEvent === undefined
        ? this.computeChangesAfterMutation(beforeVisibility, beforeNames)
        : (() => {
            const afterVisibility = this.captureVisibilitySnapshot()
            const afterNames = this.captureNamedExpressionValueSnapshot()
            this.visibilityCache = afterVisibility
            this.namedExpressionValueCache = afterNames
            return [
              ...this.computeCellChanges(beforeVisibility, afterVisibility),
              ...this.computeNamedExpressionChanges(beforeNames, afterNames),
            ]
          })()
    if (semanticEvent) {
      const event = withEventChanges(semanticEvent, changes)
      if (this.shouldSuppressEvents()) {
        this.queuedEvents.push(event)
      } else {
        this.emitter.emitDetailed(event)
      }
    }
    if (!this.shouldSuppressEvents() && changes.length > 0) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  private shouldSuppressEvents(): boolean {
    return this.batchDepth > 0 || this.evaluationSuspended
  }

  private flushQueuedEvents(): void {
    const events = [...this.queuedEvents]
    this.queuedEvents.length = 0
    events.forEach((event) => {
      this.emitter.emitDetailed(event)
    })
  }

  private getUndoStack(): HistoryRecord[] {
    const stack = Reflect.get(this.engine, 'undoStack')
    if (!isHistoryRecordArray(stack)) {
      return []
    }
    return stack
  }

  private getRedoStack(): HistoryRecord[] {
    const stack = Reflect.get(this.engine, 'redoStack')
    if (!isHistoryRecordArray(stack)) {
      return []
    }
    return stack
  }

  private clearHistoryStacks(): void {
    this.getUndoStack().length = 0
    this.getRedoStack().length = 0
  }

  private historyTransactionOps(record: HistoryTransactionRecord): unknown[] {
    switch (record.kind) {
      case 'ops':
        return record.ops
      case 'single-op':
        return [record.op]
    }
  }

  private mergeUndoHistory(startIndex: number): void {
    const undoStack = this.getUndoStack()
    if (undoStack.length - startIndex <= 1) {
      return
    }
    const entries = undoStack.splice(startIndex)
    const merged: HistoryRecord = {
      forward: {
        kind: 'ops',
        ops: entries.flatMap((entry) => this.historyTransactionOps(entry.forward)),
        potentialNewCells: sumNumbers(entries.map((entry) => entry.forward.potentialNewCells)),
      },
      inverse: {
        kind: 'ops',
        ops: entries.toReversed().flatMap((entry) => this.historyTransactionOps(entry.inverse)),
        potentialNewCells: sumNumbers(entries.map((entry) => entry.inverse.potentialNewCells)),
      },
    }
    undoStack.push(merged)
  }

  private nextSheetName(): string {
    let index = 1
    while (this.doesSheetExist(`Sheet${index}`)) {
      index += 1
    }
    return `Sheet${index}`
  }

  private buildNullMatrixForRange(range: WorkPaperCellRange): RawCellContent[][] {
    const height = range.end.row - range.start.row + 1
    const width = range.end.col - range.start.col + 1
    return Array.from({ length: height }, () => Array.from({ length: width }, () => null))
  }

  private applySerializedMatrix(
    targetLeftCorner: WorkPaperCellAddress,
    serialized: RawCellContent[][],
    sourceAnchor: WorkPaperCellAddress,
  ): void {
    this.flushPendingBatchOps()
    if (matrixContainsFormulaContent(serialized)) {
      serialized.forEach((row, rowOffset) => {
        row.forEach((raw, columnOffset) => {
          const destination = {
            sheet: targetLeftCorner.sheet,
            row: targetLeftCorner.row + rowOffset,
            col: targetLeftCorner.col + columnOffset,
          }
          let nextValue = raw
          if (typeof raw === 'string' && raw.startsWith('=')) {
            nextValue = `=${translateFormulaReferences(
              raw.slice(1),
              destination.row - (sourceAnchor.row + rowOffset),
              destination.col - (sourceAnchor.col + columnOffset),
            )}`
          }
          this.applyRawContent(destination, nextValue)
        })
      })
      return
    }

    const { refs, potentialNewCells } = buildMatrixMutationPlan({
      target: targetLeftCorner,
      content: serialized,
      rewriteFormula: (formula, destination, rowOffset, columnOffset) =>
        this.rewriteFormulaForStorage(
          translateFormulaReferences(
            stripLeadingEquals(formula),
            destination.row - (sourceAnchor.row + rowOffset),
            destination.col - (sourceAnchor.col + columnOffset),
          ),
          destination.sheet,
        ),
    })
    if (refs.length === 0) {
      return
    }
    this.applyCellMutationRefs(refs, {
      captureUndo: true,
      potentialNewCells,
      source: 'local',
      returnUndoOps: false,
      reuseRefs: true,
    })
  }

  private applyMatrixContents(
    address: WorkPaperCellAddress,
    content: WorkPaperSheet,
    options: {
      captureUndo?: boolean
      deferLiteralAddresses?: ReadonlySet<string>
      skipNulls?: boolean
    } = {},
  ): void {
    this.flushPendingBatchOps()
    const { leadingRefs, formulaRefs, refs, potentialNewCells, trailingLiteralRefs } = buildMatrixMutationPlan({
      target: address,
      content,
      deferLiteralAddresses: options.deferLiteralAddresses,
      skipNulls: options.skipNulls,
      rewriteFormula: (formula, destination) => this.rewriteFormulaForStorage(stripLeadingEquals(formula), destination.sheet),
    })
    if (refs.length === 0) {
      return
    }
    const applyPlannedRefs = (
      phaseRefs: readonly (typeof refs)[number][],
      applyOptions: {
        captureUndo?: boolean
        potentialNewCells?: number
        source?: 'local' | 'restore'
        returnUndoOps?: boolean
        reuseRefs?: boolean
      },
    ): void => {
      if (phaseRefs.length === 0) {
        return
      }
      this.applyCellMutationRefs(phaseRefs, applyOptions)
    }
    const phaseSource = options.captureUndo === false ? 'restore' : 'local'

    if (formulaRefs.length === 0) {
      applyPlannedRefs(refs, {
        captureUndo: options.captureUndo,
        potentialNewCells,
        source: phaseSource,
        returnUndoOps: false,
        reuseRefs: true,
      })
      return
    }

    applyPlannedRefs(leadingRefs, {
      captureUndo: options.captureUndo,
      potentialNewCells,
      source: phaseSource,
      returnUndoOps: false,
      reuseRefs: true,
    })
    applyPlannedRefs(formulaRefs, {
      captureUndo: options.captureUndo,
      potentialNewCells,
      source: phaseSource,
      returnUndoOps: false,
      reuseRefs: true,
    })
    applyPlannedRefs(trailingLiteralRefs, {
      captureUndo: options.captureUndo,
      potentialNewCells,
      source: phaseSource,
      returnUndoOps: false,
      reuseRefs: true,
    })
  }

  private replaceSheetContentInternal(sheetId: number, content: WorkPaperSheet, options: { duringInitialization: boolean }): void {
    replaceWorkPaperSheetContent({
      sheetId,
      sheetName: this.sheetName(sheetId),
      content,
      duringInitialization: options.duringInitialization,
      listSpills: () => this.engine.workbook.listSpills(),
      getSheetDimensions: (nextSheetId) => this.getSheetDimensions(nextSheetId),
      clearRange: (input) => this.engine.clearRange(input),
      applyMatrixContents: (address, nextContent, applyOptions) => this.applyMatrixContents(address, nextContent, applyOptions),
      clearHistoryStacks: () => this.clearHistoryStacks(),
      getUndoStackLength: () => this.getUndoStack().length,
      mergeUndoHistory: (undoStackStart) => this.mergeUndoHistory(undoStackStart),
    })
  }

  private applyRawContent(address: WorkPaperCellAddress, content: RawCellContent): void {
    const existingCellIndex = this.engine.workbook.getSheetById(address.sheet)?.grid.get(address.row, address.col) ?? -1
    const mutation: EngineCellMutationRef['mutation'] =
      content === null
        ? { kind: 'clearCell', row: address.row, col: address.col }
        : typeof content === 'string' && content.trim().startsWith('=')
          ? {
              kind: 'setCellFormula',
              row: address.row,
              col: address.col,
              formula: this.rewriteFormulaForStorage(stripLeadingEquals(content), address.sheet),
            }
          : {
              kind: 'setCellValue',
              row: address.row,
              col: address.col,
              value: content,
            }
    this.applyCellMutationRefs([{ sheetId: address.sheet, mutation }], {
      captureUndo: true,
      potentialNewCells: content === null || existingCellIndex !== -1 ? 0 : 1,
      source: 'local',
      returnUndoOps: false,
      reuseRefs: true,
    })
  }

  private captureFunctionRegistry(): void {
    const allowedPluginIds =
      this.config.functionPlugins && this.config.functionPlugins.length > 0
        ? new Set(this.config.functionPlugins.map((plugin) => plugin.id))
        : undefined
    WorkPaper.functionPluginRegistry.forEach((plugin) => {
      if (allowedPluginIds && !allowedPluginIds.has(plugin.id)) {
        return
      }
      Object.keys(plugin.implementedFunctions).forEach((functionId) => {
        const normalized = functionId.trim().toUpperCase()
        const internalName = `__BILIG_WORKPAPER_FN_${this.workbookId}_${normalized}`
        const implementation = plugin.functions?.[normalized]
        const binding: InternalFunctionBinding = {
          pluginId: plugin.id,
          publicName: normalized,
          internalName,
          implementation,
        }
        this.functionSnapshot.set(normalized, binding)
        this.functionAliasLookup.set(normalized, binding)
        this.internalFunctionLookup.set(internalName, binding)
        if (implementation) {
          globalCustomFunctions.set(internalName, implementation)
        }
      })
      Object.entries(plugin.aliases ?? {}).forEach(([alias, target]) => {
        const binding = this.functionSnapshot.get(target.trim().toUpperCase())
        if (!binding) {
          return
        }
        this.functionAliasLookup.set(alias.trim().toUpperCase(), binding)
      })
    })
  }

  private clearFunctionBindings(): void {
    this.internalFunctionLookup.forEach((_binding, internalName) => {
      globalCustomFunctions.delete(internalName)
    })
    this.functionSnapshot.clear()
    this.functionAliasLookup.clear()
    this.internalFunctionLookup.clear()
  }

  private validateCurrentSheetsWithinLimits(nextConfig: WorkPaperConfig): void {
    this.listSheetRecords().forEach((sheet) => {
      const dimensions = this.getSheetDimensions(sheet.id)
      if (dimensions.height > (nextConfig.maxRows ?? MAX_ROWS) || dimensions.width > (nextConfig.maxColumns ?? MAX_COLS)) {
        throw new WorkPaperSheetSizeLimitExceededError()
      }
    })
  }

  private canReuseSnapshotRebuild(nextConfig: WorkPaperConfig): boolean {
    if (this.config.language !== nextConfig.language) {
      return false
    }
    const currentPluginIds = functionPluginIds(this.config)
    const nextPluginIds = functionPluginIds(nextConfig)
    if (currentPluginIds.length !== nextPluginIds.length) {
      return false
    }
    for (let index = 0; index < currentPluginIds.length; index += 1) {
      if (currentPluginIds[index] !== nextPluginIds[index]) {
        return false
      }
    }
    return true
  }

  private canApplyRuntimeOnlyConfigUpdate(changedKeys: readonly (keyof WorkPaperConfig)[]): boolean {
    return changedKeys.every((key) => key === 'useColumnIndex' || key === 'useStats')
  }

  private applyRuntimeOnlyConfigUpdate(nextConfig: WorkPaperConfig): void {
    if (this.config.useColumnIndex !== nextConfig.useColumnIndex) {
      ;(this.engine as SpreadsheetEngine & { setUseColumnIndexEnabled(enabled: boolean): void }).setUseColumnIndexEnabled(
        nextConfig.useColumnIndex ?? false,
      )
    }
    this.config = cloneConfig(nextConfig)
  }

  private rebuildWithConfig(nextConfig: WorkPaperConfig): void {
    this.validateCurrentSheetsWithinLimits(nextConfig)
    const canReuseSnapshot = this.canReuseSnapshotRebuild(nextConfig)
    const snapshot = canReuseSnapshot ? this.engine.exportSnapshot() : null
    const serializedSheets = canReuseSnapshot ? null : this.getAllSheetsSerialized()
    if (serializedSheets) {
      Object.entries(serializedSheets).forEach(([sheetName, sheet]) => {
        validateSheetWithinLimits(sheetName, sheet, nextConfig)
      })
    }
    const serializedNamedExpressions = canReuseSnapshot ? null : this.getAllNamedExpressionsSerialized()
    const suspended = this.evaluationSuspended
    const clipboard = this.clipboard
      ? {
          sourceAnchor: { ...this.clipboard.sourceAnchor },
          serialized: this.clipboard.serialized.map((row) => [...row]),
          values: this.clipboard.values.map((row) => row.map((value) => cloneCellValue(value))),
        }
      : null

    this.clearFunctionBindings()
    if (!canReuseSnapshot) {
      this.namedExpressions.clear()
    }
    this.engine = new SpreadsheetEngine({
      workbookName: 'Workbook',
      useColumnIndex: this.config.useColumnIndex,
      trackReplicaVersions: false,
    })
    this.attachEngineEventTracking()
    this.config = cloneConfig(nextConfig)
    this.captureFunctionRegistry()

    this.withEngineEventCaptureDisabled(() => {
      if (snapshot) {
        this.engine.importSnapshot(snapshot)
      } else {
        Object.keys(serializedSheets!).forEach((sheetName) => {
          this.engine.createSheet(sheetName)
        })
        serializedNamedExpressions!.forEach((expression) => {
          this.upsertNamedExpressionInternal(expression, { duringInitialization: true })
        })
        Object.entries(serializedSheets!).forEach(([sheetName, sheet]) => {
          const sheetId = this.requireSheetId(sheetName)
          if (tryLoadInitialLiteralSheet(this.engine, sheetId, sheet)) {
            return
          }
          this.replaceSheetContentInternal(sheetId, sheet, { duringInitialization: true })
        })
      }
    })
    this.clearHistoryStacks()
    this.resetChangeTrackingCaches()
    this.clipboard = clipboard
    if (suspended) {
      this.suspendedVisibility = this.ensureVisibilityCache()
      this.suspendedNamedValues = this.ensureNamedExpressionValueCache()
    }
  }

  private normalizeAxisIntervals(
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    restIntervals: readonly WorkPaperAxisInterval[] = [],
  ): Array<[number, number]> {
    if (typeof startOrInterval === 'number') {
      if (Array.isArray(countOrInterval)) {
        throw new WorkPaperInvalidArgumentsError('Axis interval count must be a number')
      }
      const resolvedCount = typeof countOrInterval === 'number' ? countOrInterval : 1
      return [[startOrInterval, resolvedCount]]
    }
    if (typeof countOrInterval === 'number') {
      throw new WorkPaperInvalidArgumentsError('Axis interval count is only valid with a numeric start')
    }
    return [startOrInterval, ...(countOrInterval ? [countOrInterval] : []), ...restIntervals].map(
      ([start, count]) => [start, count ?? 1] as [number, number],
    )
  }

  private normalizeAxisSwapMappings(
    label: 'row' | 'column',
    startOrMappings: number | readonly WorkPaperAxisSwapMapping[],
    end?: number,
  ): WorkPaperAxisSwapMapping[] {
    if (typeof startOrMappings === 'number') {
      if (end === undefined) {
        throw new WorkPaperInvalidArgumentsError(`${label} swap requires two indexes`)
      }
      return [[startOrMappings, end]]
    }
    return [...startOrMappings]
  }

  private collectRangeDependencies(
    range: WorkPaperCellRange,
    readDependencies: (address: WorkPaperCellAddress) => readonly string[],
  ): WorkPaperDependencyRef[] {
    assertRange(range)
    const seen = new Set<string>()
    const collected: WorkPaperDependencyRef[] = []
    this.getDenseRange(range, (address) => address).forEach((row) => {
      row.forEach((address) => {
        this.toDependencyRefs(readDependencies(address)).forEach((dependency) => {
          const key =
            dependency.kind === 'cell'
              ? `cell:${dependency.address.sheet}:${dependency.address.row}:${dependency.address.col}`
              : dependency.kind === 'range'
                ? `range:${dependency.range.start.sheet}:${dependency.range.start.row}:${dependency.range.start.col}:${dependency.range.end.row}:${dependency.range.end.col}`
                : `name:${dependency.name}`
          if (seen.has(key)) {
            return
          }
          seen.add(key)
          collected.push(dependency)
        })
      })
    })
    return collected
  }

  private rewriteFormulaForStorage(formula: string, ownerSheetId: number): string {
    if (this.namedExpressions.size === 0 && this.functionAliasLookup.size === 0) {
      return formula
    }
    try {
      const transformed = transformFormulaNode(parseFormula(stripLeadingEquals(formula)), (node) => {
        if (node.kind === 'NameRef') {
          return this.rewriteNameRefForStorage(node, ownerSheetId)
        }
        if (node.kind === 'CallExpr') {
          return this.rewriteCallForStorage(node)
        }
        return node
      })
      return serializeFormula(transformed)
    } catch (error) {
      throw new WorkPaperParseError(this.messageOf(error, 'Unable to store formula'))
    }
  }

  private restorePublicFormula(formula: string, ownerSheetId: number): string {
    if (this.namedExpressions.size === 0 && this.functionAliasLookup.size === 0) {
      return formula
    }
    const transformed = transformFormulaNode(parseFormula(formula), (node) => {
      if (node.kind === 'NameRef') {
        return this.rewriteNameRefForPublic(node, ownerSheetId)
      }
      if (node.kind === 'CallExpr') {
        return this.rewriteCallForPublic(node)
      }
      return node
    })
    return serializeFormula(transformed)
  }

  private rewriteNameRefForStorage(node: NameRefNode, ownerSheetId: number): FormulaNode {
    const scoped = this.namedExpressions.get(makeNamedExpressionKey(node.name, ownerSheetId))
    if (scoped) {
      return { ...node, name: scoped.internalName }
    }
    const workbookScoped = this.namedExpressions.get(makeNamedExpressionKey(node.name))
    if (workbookScoped) {
      return { ...node, name: workbookScoped.internalName }
    }
    return node
  }

  private rewriteNameRefForPublic(node: NameRefNode, ownerSheetId: number): FormulaNode {
    const exact = [...this.namedExpressions.values()].find(
      (expression) => expression.internalName === node.name && expression.scope === ownerSheetId,
    )
    if (exact) {
      return { ...node, name: exact.publicName }
    }
    return node
  }

  private rewriteCallForStorage(node: CallExprNode): FormulaNode {
    const binding = this.functionAliasLookup.get(node.callee.trim().toUpperCase())
    if (!binding) {
      return node
    }
    return { ...node, callee: binding.internalName }
  }

  private rewriteCallForPublic(node: CallExprNode): FormulaNode {
    const binding = this.internalFunctionLookup.get(node.callee.trim().toUpperCase())
    if (!binding) {
      return node
    }
    return { ...node, callee: binding.publicName }
  }

  private validateNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): void {
    const trimmed = expressionName.trim()
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed) || isCellReferenceText(trimmed)) {
      throw new WorkPaperNamedExpressionNameIsInvalidError(expressionName)
    }
    if (scope !== undefined) {
      this.sheetRecord(scope)
    }
    if (isFormulaContent(expression)) {
      try {
        const parsed = parseFormula(stripLeadingEquals(expression))
        if (formulaHasRelativeReferences(parsed)) {
          throw new WorkPaperNoRelativeAddressesAllowedError()
        }
      } catch (error) {
        if (error instanceof WorkPaperNoRelativeAddressesAllowedError) {
          throw error
        }
        throw new WorkPaperUnableToParseError({
          expressionName,
          reason: this.messageOf(error, `Invalid named expression formula for '${expressionName}'`),
        })
      }
    }
  }

  private upsertNamedExpressionInternal(expression: SerializedWorkPaperNamedExpression, options: { duringInitialization: boolean }): void {
    this.validateNamedExpression(expression.name, expression.expression, expression.scope)
    const trimmed = expression.name.trim()
    const internalName = expression.scope === undefined ? trimmed : makeInternalScopedName(expression.scope, trimmed)
    const record: InternalNamedExpression = {
      publicName: trimmed,
      normalizedName: normalizeName(trimmed),
      internalName,
      scope: expression.scope,
      expression: expression.expression,
      options: expression.options ? structuredClone(expression.options) : undefined,
    }
    this.namedExpressions.set(makeNamedExpressionKey(trimmed, expression.scope), record)
    this.engine.setDefinedName(internalName, this.toDefinedNameSnapshot(record.expression, record.scope))
    if (options.duringInitialization) {
      this.clearHistoryStacks()
    }
  }

  private toDefinedNameSnapshot(expression: RawCellContent, scope?: number): WorkbookDefinedNameValueSnapshot {
    if (expression === null || typeof expression === 'number' || typeof expression === 'boolean') {
      return expression
    }
    if (typeof expression === 'string' && expression.trim().startsWith('=')) {
      return {
        kind: 'formula',
        formula: `=${this.rewriteFormulaForStorage(stripLeadingEquals(expression), scope ?? this.listSheetRecords()[0]?.id ?? 1)}`,
      }
    }
    return expression
  }

  private namedExpressionRecord(name: string, scope?: number): InternalNamedExpression {
    const direct = this.namedExpressions.get(makeNamedExpressionKey(name, scope))
    if (direct) {
      return direct
    }
    throw new WorkPaperNamedExpressionDoesNotExistError(name)
  }

  private evaluateNamedExpression(expression: InternalNamedExpression): CellValue | CellValue[][] {
    const raw = expression.expression
    if (raw === null || typeof raw === 'number' || typeof raw === 'boolean') {
      return scalarValueFromLiteral(raw)
    }
    if (typeof raw === 'string' && !raw.trim().startsWith('=')) {
      return scalarValueFromLiteral(raw)
    }
    return this.calculateFormula(raw, expression.scope)
  }

  private cellSnapshotToRawContent(cell: CellSnapshot, ownerSheetId: number): RawCellContent {
    if (cell.formula) {
      return `=${this.restorePublicFormula(cell.formula, ownerSheetId)}`
    }
    if (cell.input !== undefined) {
      return cell.input
    }
    switch (cell.value.tag) {
      case ValueTag.Empty:
      case ValueTag.Error:
        return null
      case ValueTag.Number:
        return cell.value.value
      case ValueTag.Boolean:
        return cell.value.value
      case ValueTag.String:
        return cell.value.value
    }
  }

  private toDependencyRefs(values: readonly string[]): WorkPaperDependencyRef[] {
    return values.map((value) => {
      try {
        const parsedCell = parseCellAddress(value)
        return {
          kind: 'cell',
          address: {
            sheet: this.requireSheetId(parsedCell.sheetName ?? this.listSheetRecords()[0]!.name),
            row: parsedCell.row,
            col: parsedCell.col,
          },
        } satisfies WorkPaperDependencyRef
      } catch {
        try {
          const parsedRange = parseRangeAddress(value)
          if (parsedRange.kind === 'cells') {
            return {
              kind: 'range',
              range: {
                start: {
                  sheet: this.requireSheetId(parsedRange.sheetName ?? this.listSheetRecords()[0]!.name),
                  row: parsedRange.start.row,
                  col: parsedRange.start.col,
                },
                end: {
                  sheet: this.requireSheetId(parsedRange.sheetName ?? this.listSheetRecords()[0]!.name),
                  row: parsedRange.end.row,
                  col: parsedRange.end.col,
                },
              },
            } satisfies WorkPaperDependencyRef
          }
        } catch {
          return { kind: 'name', name: value } satisfies WorkPaperDependencyRef
        }
      }
      return { kind: 'name', name: value } satisfies WorkPaperDependencyRef
    })
  }

  private messageOf(error: unknown, fallback: string): string {
    return error instanceof Error && error.message.length > 0 ? error.message : fallback
  }
}

function cloneNamedExpressionValue(value: CellValue | CellValue[][]): CellValue | CellValue[][] {
  if (!Array.isArray(value)) {
    return cloneCellValue(value)
  }
  return value.map((row) => row.map((cell) => cloneCellValue(cell)))
}

function compareWorkPaperNamedExpressionChanges(left: WorkPaperChange, right: WorkPaperChange): number {
  if (left.kind !== 'named-expression' || right.kind !== 'named-expression') {
    return 0
  }
  return (left.scope ?? -1) - (right.scope ?? -1) || left.name.localeCompare(right.name)
}

function sourceRangeRef(sheetName: string, range: WorkPaperCellRange): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(range.start.row, range.start.col),
    endAddress: formatAddress(range.end.row, range.end.col),
  }
}

function sumNumbers(values: Array<number | undefined>): number | undefined {
  const filtered = values.filter((value): value is number => typeof value === 'number')
  if (filtered.length === 0) {
    return undefined
  }
  return filtered.reduce((sum, value) => sum + value, 0)
}
