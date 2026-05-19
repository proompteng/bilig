import { MAX_COLS, MAX_ROWS, type WorkbookCalculationSettingsSnapshot } from '@bilig/protocol'
import { normalizeWorkbookCalculationSettings } from '@bilig/core/headless-runtime'
import {
  WorkPaperConfigValueTooBigError,
  WorkPaperConfigValueTooSmallError,
  WorkPaperExpectedOneOfValuesError,
  WorkPaperExpectedValueOfTypeError,
} from './work-paper-errors.js'
import type {
  WorkPaperCalculationSettings,
  WorkPaperConfig,
  WorkPaperFunctionPluginDefinition,
  WorkPaperLicenseKeyValidityState,
} from './work-paper-types.js'

export const DEFAULT_CONFIG: Readonly<WorkPaperConfig> = Object.freeze({
  accentSensitive: false,
  caseSensitive: false,
  caseFirst: 'false',
  calculationSettings: undefined,
  chooseAddressMappingPolicy: undefined,
  context: undefined,
  currencySymbol: ['$'],
  dateFormats: [],
  functionArgSeparator: ',',
  decimalSeparator: '.',
  evaluateNullToZero: true,
  evaluationTimeoutMs: undefined,
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
  useColumnIndex: true,
  useStats: true,
  undoLimit: 100,
  useRegularExpressions: true,
  useWildcards: true,
})

export const WORKPAPER_CONFIG_KEYS = [
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

export const WORKPAPER_PUBLIC_ERROR_NAMES = new Set([
  'WorkPaperConfigValueTooBigError',
  'WorkPaperConfigValueTooSmallError',
  'WorkPaperEvaluationSuspendedError',
  'WorkPaperEvaluationTimeoutError',
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

export function clonePluginDefinition(plugin: WorkPaperFunctionPluginDefinition): WorkPaperFunctionPluginDefinition {
  const cloned: WorkPaperFunctionPluginDefinition = {
    ...plugin,
    implementedFunctions: Object.fromEntries(
      Object.entries(plugin.implementedFunctions).map(([name, metadata]) => [name, { ...metadata }]),
    ),
  }
  if (plugin.aliases !== undefined) {
    cloned.aliases = { ...plugin.aliases }
  }
  if (plugin.functions !== undefined) {
    cloned.functions = { ...plugin.functions }
  }
  return cloned
}

export function cloneConfig(config: WorkPaperConfig): WorkPaperConfig {
  const cloned: WorkPaperConfig = {
    ...config,
  }
  if (config.calculationSettings !== undefined) {
    cloned.calculationSettings = { ...config.calculationSettings }
  }
  if (config.chooseAddressMappingPolicy !== undefined) {
    cloned.chooseAddressMappingPolicy = { ...config.chooseAddressMappingPolicy }
  }
  if (config.context !== undefined) {
    cloned.context = structuredClone(config.context)
  }
  if (config.currencySymbol !== undefined) {
    cloned.currencySymbol = [...config.currencySymbol]
  }
  if (config.dateFormats !== undefined) {
    cloned.dateFormats = [...config.dateFormats]
  }
  if (config.functionPlugins !== undefined) {
    cloned.functionPlugins = config.functionPlugins.map((plugin) => clonePluginDefinition(plugin))
  }
  if (config.nullDate !== undefined) {
    cloned.nullDate = { ...config.nullDate }
  }
  if (config.timeFormats !== undefined) {
    cloned.timeFormats = [...config.timeFormats]
  }
  return cloned
}

export function normalizeConfiguredWorkPaperCalculationSettings(
  settings: WorkPaperCalculationSettings | undefined,
  base?: WorkbookCalculationSettingsSnapshot,
): WorkbookCalculationSettingsSnapshot | undefined {
  if (settings === undefined) {
    return undefined
  }
  return normalizeWorkbookCalculationSettings(settings, base)
}

function validateWorkPaperCalculationSettings(settings: WorkPaperCalculationSettings): void {
  if (settings.mode !== undefined && settings.mode !== 'automatic' && settings.mode !== 'manual') {
    throw new WorkPaperExpectedOneOfValuesError('"automatic", "manual"', 'calculationSettings.mode')
  }
  if (
    settings.compatibilityMode !== undefined &&
    settings.compatibilityMode !== 'excel-modern' &&
    settings.compatibilityMode !== 'odf-1.4'
  ) {
    throw new WorkPaperExpectedOneOfValuesError('"excel-modern", "odf-1.4"', 'calculationSettings.compatibilityMode')
  }
  if (settings.dateSystem !== undefined && settings.dateSystem !== '1900' && settings.dateSystem !== '1904') {
    throw new WorkPaperExpectedOneOfValuesError('"1900", "1904"', 'calculationSettings.dateSystem')
  }
  if (settings.iterate !== undefined && settings.iterate !== null && typeof settings.iterate !== 'boolean') {
    throw new WorkPaperExpectedValueOfTypeError('boolean', 'calculationSettings.iterate')
  }
  if (settings.iterateCount !== undefined && settings.iterateCount !== null) {
    if (!Number.isSafeInteger(settings.iterateCount)) {
      throw new WorkPaperExpectedValueOfTypeError('safe integer', 'calculationSettings.iterateCount')
    }
    if (settings.iterateCount < 1) {
      throw new WorkPaperConfigValueTooSmallError('calculationSettings.iterateCount', 1)
    }
  }
  if (settings.iterateDelta !== undefined && settings.iterateDelta !== null) {
    if (typeof settings.iterateDelta !== 'string' || !Number.isFinite(Number(settings.iterateDelta))) {
      throw new WorkPaperExpectedValueOfTypeError('finite numeric string', 'calculationSettings.iterateDelta')
    }
  }
  if (settings.fullPrecision !== undefined && settings.fullPrecision !== null && typeof settings.fullPrecision !== 'boolean') {
    throw new WorkPaperExpectedValueOfTypeError('boolean', 'calculationSettings.fullPrecision')
  }
  if (settings.fullCalcOnLoad !== undefined && settings.fullCalcOnLoad !== null && typeof settings.fullCalcOnLoad !== 'boolean') {
    throw new WorkPaperExpectedValueOfTypeError('boolean', 'calculationSettings.fullCalcOnLoad')
  }
  if (settings.concurrentCalc !== undefined && settings.concurrentCalc !== null && typeof settings.concurrentCalc !== 'boolean') {
    throw new WorkPaperExpectedValueOfTypeError('boolean', 'calculationSettings.concurrentCalc')
  }
}

export function checkWorkPaperLicenseKeyValidity(licenseKey: string | undefined): WorkPaperLicenseKeyValidityState {
  if (!licenseKey || licenseKey.trim().length === 0) {
    return 'missing'
  }
  if (licenseKey === 'internal' || licenseKey === 'gpl-v3' || licenseKey === 'internal-use-in-handsontable') {
    return 'valid'
  }
  return 'invalid'
}

export function validateWorkPaperConfig(config: WorkPaperConfig): void {
  if (config.maxRows !== undefined && (!Number.isInteger(config.maxRows) || config.maxRows < 1)) {
    throw new WorkPaperConfigValueTooSmallError('maxRows', 1)
  }
  if (config.maxColumns !== undefined && (!Number.isInteger(config.maxColumns) || config.maxColumns < 1)) {
    throw new WorkPaperConfigValueTooSmallError('maxColumns', 1)
  }
  if (config.evaluationTimeoutMs !== undefined && (!Number.isFinite(config.evaluationTimeoutMs) || config.evaluationTimeoutMs < 0)) {
    throw new WorkPaperConfigValueTooSmallError('evaluationTimeoutMs', 0)
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
  if (config.calculationSettings !== undefined) {
    if (
      typeof config.calculationSettings !== 'object' ||
      config.calculationSettings === null ||
      Array.isArray(config.calculationSettings)
    ) {
      throw new WorkPaperExpectedValueOfTypeError('object', 'calculationSettings')
    }
    validateWorkPaperCalculationSettings(config.calculationSettings)
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

export function functionPluginIds(config: WorkPaperConfig): string[] {
  return (config.functionPlugins ?? []).map((plugin) => plugin.id).toSorted()
}

export function canReuseWorkPaperSnapshotRebuild(currentConfig: WorkPaperConfig, nextConfig: WorkPaperConfig): boolean {
  if (currentConfig.language !== nextConfig.language) {
    return false
  }
  const currentPluginIds = functionPluginIds(currentConfig)
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

export function canApplyRuntimeOnlyWorkPaperConfigUpdate(changedKeys: readonly (keyof WorkPaperConfig)[]): boolean {
  return changedKeys.every((key) => key === 'useColumnIndex' || key === 'useStats' || key === 'evaluationTimeoutMs')
}
