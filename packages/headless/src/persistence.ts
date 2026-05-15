import type { LiteralInput } from '@bilig/protocol'
import { WorkPaper } from './work-paper.js'
import { WorkPaperPersistenceError } from './work-paper-errors.js'
import { isFormulaContent } from './work-paper-runtime-helpers.js'
import type {
  RawCellContent,
  SerializedWorkPaperNamedExpression,
  WorkPaperCalculationSettings,
  WorkPaperChooseAddressMappingPolicy,
  WorkPaperConfig,
  WorkPaperContextValue,
  WorkPaperSheet,
} from './work-paper-types.js'

export const WORK_PAPER_DOCUMENT_FORMAT = 'bilig.headless.work-paper.document.v1' as const

export const PERSISTABLE_WORK_PAPER_CONFIG_KEYS = [
  'accentSensitive',
  'caseSensitive',
  'caseFirst',
  'calculationSettings',
  'chooseAddressMappingPolicy',
  'context',
  'currencySymbol',
  'dateFormats',
  'functionArgSeparator',
  'decimalSeparator',
  'evaluateNullToZero',
  'evaluationTimeoutMs',
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
  'precisionEpsilon',
  'precisionRounding',
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

type PersistableWorkPaperConfigKey = (typeof PERSISTABLE_WORK_PAPER_CONFIG_KEYS)[number]

export type PersistableWorkPaperConfig = Pick<WorkPaperConfig, PersistableWorkPaperConfigKey>

export interface PersistedWorkPaperNamedExpression {
  name: string
  expression: RawCellContent
  scopeSheetName?: string
  options?: Record<string, string | number | boolean>
}

export interface PersistedWorkPaperSheet {
  name: string
  content: WorkPaperSheet
}

export interface PersistedWorkPaperDocument {
  format: typeof WORK_PAPER_DOCUMENT_FORMAT
  sheets: PersistedWorkPaperSheet[]
  namedExpressions: PersistedWorkPaperNamedExpression[]
  config?: PersistableWorkPaperConfig
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isDenseArray(value: unknown): value is unknown[] {
  if (!Array.isArray(value)) {
    return false
  }
  for (let index = 0; index < value.length; index += 1) {
    if (!Object.hasOwn(value, index)) {
      return false
    }
  }
  return true
}

function isJsonNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function isJsonInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value)
}

function isPositiveJsonInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value > 0
}

function isNonNegativeJsonInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'boolean' || isJsonNumber(value) || typeof value === 'string'
}

function isRawCellContent(value: unknown): value is RawCellContent {
  return isLiteralInput(value) || typeof value === 'string'
}

function isWorkPaperSheet(value: unknown): value is WorkPaperSheet {
  return isDenseArray(value) && value.every((row) => isDenseArray(row) && row.every((cell) => isRawCellContent(cell)))
}

function isNamedExpressionOptions(value: unknown): value is Record<string, string | number | boolean> {
  return (
    isRecord(value) && Object.values(value).every((entry) => typeof entry === 'string' || isJsonNumber(entry) || typeof entry === 'boolean')
  )
}

function isWorkPaperContextValue(value: unknown): value is WorkPaperContextValue {
  return (
    value === null ||
    typeof value === 'string' ||
    isJsonNumber(value) ||
    typeof value === 'boolean' ||
    (isDenseArray(value) && value.every((entry) => isWorkPaperContextValue(entry))) ||
    (isRecord(value) && Object.values(value).every((entry) => isWorkPaperContextValue(entry)))
  )
}

function isChooseAddressMappingPolicy(value: unknown): value is WorkPaperChooseAddressMappingPolicy {
  return isRecord(value) && (value['mode'] === 'dense' || value['mode'] === 'sparse')
}

function isCalculationSettings(value: unknown): value is WorkPaperCalculationSettings {
  return (
    isRecord(value) &&
    (value['mode'] === undefined || value['mode'] === 'automatic' || value['mode'] === 'manual') &&
    (value['compatibilityMode'] === undefined ||
      value['compatibilityMode'] === 'excel-modern' ||
      value['compatibilityMode'] === 'odf-1.4') &&
    (value['dateSystem'] === undefined || value['dateSystem'] === '1900' || value['dateSystem'] === '1904') &&
    (value['iterate'] === undefined || value['iterate'] === null || typeof value['iterate'] === 'boolean') &&
    (value['iterateCount'] === undefined || value['iterateCount'] === null || isPositiveJsonInteger(value['iterateCount'])) &&
    (value['iterateDelta'] === undefined ||
      value['iterateDelta'] === null ||
      (typeof value['iterateDelta'] === 'string' && Number.isFinite(Number(value['iterateDelta'])))) &&
    (value['fullPrecision'] === undefined || value['fullPrecision'] === null || typeof value['fullPrecision'] === 'boolean') &&
    (value['fullCalcOnLoad'] === undefined || value['fullCalcOnLoad'] === null || typeof value['fullCalcOnLoad'] === 'boolean') &&
    (value['concurrentCalc'] === undefined || value['concurrentCalc'] === null || typeof value['concurrentCalc'] === 'boolean')
  )
}

function isStringArray(value: unknown): value is string[] {
  return isDenseArray(value) && value.every((item) => typeof item === 'string')
}

function isSimpleDate(value: unknown): value is { year: number; month: number; day: number } {
  return (
    isRecord(value) &&
    isJsonInteger(value['year']) &&
    isJsonInteger(value['month']) &&
    value['month'] >= 1 &&
    value['month'] <= 12 &&
    isJsonInteger(value['day']) &&
    value['day'] >= 1 &&
    value['day'] <= 31
  )
}

function isPersistableConfigEntry(key: string, entry: unknown): boolean {
  switch (key) {
    case 'accentSensitive':
    case 'caseSensitive':
    case 'evaluateNullToZero':
    case 'ignorePunctuation':
    case 'leapYear1900':
    case 'matchWholeCell':
    case 'smartRounding':
    case 'useArrayArithmetic':
    case 'useColumnIndex':
    case 'useStats':
    case 'useRegularExpressions':
    case 'useWildcards':
      return typeof entry === 'boolean'
    case 'caseFirst':
      return entry === 'upper' || entry === 'lower' || entry === 'false'
    case 'calculationSettings':
      return isCalculationSettings(entry)
    case 'chooseAddressMappingPolicy':
      return isChooseAddressMappingPolicy(entry)
    case 'context':
      return isWorkPaperContextValue(entry)
    case 'currencySymbol':
    case 'dateFormats':
    case 'timeFormats':
      return isStringArray(entry)
    case 'decimalSeparator':
      return entry === '.' || entry === ','
    case 'evaluationTimeoutMs':
      return isJsonNumber(entry) && entry >= 0
    case 'functionArgSeparator':
    case 'language':
    case 'licenseKey':
    case 'localeLang':
      return typeof entry === 'string'
    case 'ignoreWhiteSpace':
      return entry === 'standard' || entry === 'any'
    case 'arrayColumnSeparator':
      return entry === ',' || entry === ';'
    case 'arrayRowSeparator':
      return entry === ';' || entry === '|'
    case 'maxRows':
    case 'maxColumns':
      return isPositiveJsonInteger(entry)
    case 'nullDate':
      return isSimpleDate(entry)
    case 'nullYear':
    case 'precisionRounding':
      return isJsonInteger(entry)
    case 'precisionEpsilon':
      return isJsonNumber(entry)
    case 'thousandSeparator':
      return entry === '' || entry === ',' || entry === '.'
    case 'undoLimit':
      return isNonNegativeJsonInteger(entry)
    default:
      return false
  }
}

function isPersistableWorkPaperConfig(value: unknown): value is PersistableWorkPaperConfig {
  if (!isRecord(value)) {
    return false
  }
  return Object.entries(value).every(([key, entry]) => {
    return (PERSISTABLE_WORK_PAPER_CONFIG_KEYS as readonly string[]).includes(key) && isPersistableConfigEntry(key, entry)
  })
}

function isPersistedWorkPaperNamedExpression(value: unknown): value is PersistedWorkPaperNamedExpression {
  return (
    isRecord(value) &&
    isNonEmptyString(value['name']) &&
    isRawCellContent(value['expression']) &&
    (value['scopeSheetName'] === undefined || isNonEmptyString(value['scopeSheetName'])) &&
    (value['options'] === undefined || isNamedExpressionOptions(value['options']))
  )
}

function isPersistedWorkPaperSheet(value: unknown): value is PersistedWorkPaperSheet {
  return isRecord(value) && isNonEmptyString(value['name']) && isWorkPaperSheet(value['content'])
}

function hasUniqueSheetNames(sheets: readonly PersistedWorkPaperSheet[]): boolean {
  const names = new Set<string>()
  for (const sheet of sheets) {
    if (names.has(sheet.name)) {
      return false
    }
    names.add(sheet.name)
  }
  return true
}

function namedExpressionKey(expression: PersistedWorkPaperNamedExpression): string {
  return `${expression.scopeSheetName ?? '<global>'}\u0000${expression.name}`
}

function hasValidNamedExpressionScopes(
  expressions: readonly PersistedWorkPaperNamedExpression[],
  sheetNames: ReadonlySet<string>,
): boolean {
  const namesByScope = new Set<string>()
  for (const expression of expressions) {
    if (expression.scopeSheetName !== undefined && !sheetNames.has(expression.scopeSheetName)) {
      return false
    }
    const key = namedExpressionKey(expression)
    if (namesByScope.has(key)) {
      return false
    }
    namesByScope.add(key)
  }
  return true
}

/**
 * Checks whether a value matches the persisted WorkPaper document format.
 */
export function isPersistedWorkPaperDocument(value: unknown): value is PersistedWorkPaperDocument {
  if (
    !isRecord(value) ||
    value['format'] !== WORK_PAPER_DOCUMENT_FORMAT ||
    !isDenseArray(value['sheets']) ||
    !value['sheets'].every((sheet) => isPersistedWorkPaperSheet(sheet)) ||
    !isDenseArray(value['namedExpressions']) ||
    !value['namedExpressions'].every((expression) => isPersistedWorkPaperNamedExpression(expression)) ||
    (value['config'] !== undefined && !isPersistableWorkPaperConfig(value['config']))
  ) {
    return false
  }

  const sheets = value['sheets']
  const namedExpressions = value['namedExpressions']
  if (!hasUniqueSheetNames(sheets)) {
    return false
  }

  return hasValidNamedExpressionScopes(namedExpressions, new Set(sheets.map((sheet) => sheet.name)))
}

function assertPersistedWorkPaperDocument(value: unknown): asserts value is PersistedWorkPaperDocument {
  if (!isPersistedWorkPaperDocument(value)) {
    throw new WorkPaperPersistenceError('Invalid persisted WorkPaper document')
  }
}

function setPersistableWorkPaperConfigValue<Key extends PersistableWorkPaperConfigKey>(
  target: PersistableWorkPaperConfig,
  key: Key,
  value: PersistableWorkPaperConfig[Key],
): void {
  target[key] = structuredClone(value)
}

/**
 * Clones the documented JSON-safe subset of WorkPaper configuration values.
 */
export function pickPersistableWorkPaperConfig(config: WorkPaperConfig): PersistableWorkPaperConfig {
  const picked: PersistableWorkPaperConfig = {}
  for (const key of PERSISTABLE_WORK_PAPER_CONFIG_KEYS) {
    const value = config[key]
    if (value !== undefined) {
      setPersistableWorkPaperConfigValue(picked, key, value)
    }
  }
  return picked
}

/**
 * Exports sheets, named expressions, and optional config from a WorkPaper.
 */
export function exportWorkPaperDocument(workbook: WorkPaper, options: { includeConfig?: boolean } = {}): PersistedWorkPaperDocument {
  const { includeConfig = true } = options
  const sheets = workbook.getSheetNames().map((name) => {
    const sheetId = workbook.getSheetId(name)
    if (sheetId === undefined) {
      throw new WorkPaperPersistenceError(`Missing sheet id for ${name}`)
    }
    return {
      name,
      content: serializeSheetContentForDocument(workbook, sheetId),
    } satisfies PersistedWorkPaperSheet
  })
  const namedExpressions = workbook.getAllNamedExpressionsSerialized().map((expression) => serializeNamedExpression(workbook, expression))
  const document: PersistedWorkPaperDocument = {
    format: WORK_PAPER_DOCUMENT_FORMAT,
    sheets,
    namedExpressions,
  }
  if (includeConfig) {
    document.config = pickPersistableWorkPaperConfig(workbook.getConfig())
  }
  return document
}

function serializeSheetContentForDocument(workbook: WorkPaper, sheetId: number): RawCellContent[][] {
  return workbook
    .getSheetSerialized(sheetId)
    .map((row, rowIndex) =>
      row.map((cellContent, colIndex) =>
        cellContent !== null &&
        !isFormulaContent(cellContent) &&
        workbook.isCellPartOfArray({ sheet: sheetId, row: rowIndex, col: colIndex })
          ? null
          : cellContent,
      ),
    )
}

/**
 * Creates a WorkPaper instance from a validated persisted document.
 */
export function createWorkPaperFromDocument(document: PersistedWorkPaperDocument): WorkPaper {
  assertPersistedWorkPaperDocument(document)
  const sheetEntries = document.sheets.map((sheet) => [sheet.name, sheet.content] as const)
  const sheetIdsByName = new Map(sheetEntries.map(([name], index) => [name, index + 1] as const))
  const namedExpressions = document.namedExpressions.map((expression) => deserializeNamedExpression(expression, sheetIdsByName))
  const workbook = WorkPaper.buildFromSheetEntries(sheetEntries, document.config ?? {}, namedExpressions)
  workbook.clearUndoStack()
  workbook.clearRedoStack()
  return workbook
}

/**
 * Serializes a validated WorkPaper document to JSON.
 */
export function serializeWorkPaperDocument(document: PersistedWorkPaperDocument): string {
  assertPersistedWorkPaperDocument(document)
  return JSON.stringify(document)
}

/**
 * Parses and validates a WorkPaper document from JSON.
 */
export function parseWorkPaperDocument(json: string): PersistedWorkPaperDocument {
  let parsed: unknown
  try {
    parsed = JSON.parse(json) as unknown
  } catch (error) {
    throw new WorkPaperPersistenceError('Invalid WorkPaper document JSON', error)
  }
  assertPersistedWorkPaperDocument(parsed)
  return parsed
}

function serializeNamedExpression(workbook: WorkPaper, expression: SerializedWorkPaperNamedExpression): PersistedWorkPaperNamedExpression {
  const persisted: PersistedWorkPaperNamedExpression = {
    name: expression.name,
    expression: expression.expression,
  }
  if (expression.options !== undefined) {
    persisted.options = structuredClone(expression.options)
  }
  if (expression.scope === undefined) {
    return persisted
  }
  const scopeSheetName = workbook.getSheetName(expression.scope)
  if (!scopeSheetName) {
    throw new WorkPaperPersistenceError(`Missing scope sheet for named expression ${expression.name}`)
  }
  persisted.scopeSheetName = scopeSheetName
  return persisted
}

function deserializeNamedExpression(
  expression: PersistedWorkPaperNamedExpression,
  sheetIdsByName: ReadonlyMap<string, number>,
): SerializedWorkPaperNamedExpression {
  const serialized: SerializedWorkPaperNamedExpression = {
    name: expression.name,
    expression: expression.expression,
  }
  if (expression.options !== undefined) {
    serialized.options = structuredClone(expression.options)
  }
  if (expression.scopeSheetName === undefined) {
    return serialized
  }
  const scope = sheetIdsByName.get(expression.scopeSheetName)
  if (scope === undefined) {
    throw new WorkPaperPersistenceError(`Missing scoped sheet ${expression.scopeSheetName}`)
  }
  serialized.scope = scope
  return serialized
}
