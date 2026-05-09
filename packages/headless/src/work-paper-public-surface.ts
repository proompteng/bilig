import type { SpreadsheetEngine } from '@bilig/core'
import { formatCellDisplayValue } from '@bilig/protocol'
import type { CellValue, WorkbookCalculationSettingsSnapshot } from '@bilig/protocol'
import { checkWorkPaperLicenseKeyValidity, cloneConfig, DEFAULT_CONFIG } from './work-paper-config.js'
import { numberToWorkPaperDate, numberToWorkPaperDateTime, numberToWorkPaperTime } from './work-paper-date-time.js'
import type { WorkPaperEmitter } from './work-paper-emitter.js'
import {
  calculateWorkPaperFormula,
  getWorkPaperNamedExpressionsFromFormula,
  normalizeWorkPaperFormula,
  validateWorkPaperFormula,
} from './work-paper-formula-analysis.js'
import { collectWorkPaperFormulaDiagnostics } from './work-paper-formula-diagnostics.js'
import {
  getCapturedWorkPaperFunctionPlugin,
  getCapturedWorkPaperFunctionPlugins,
  listCapturedWorkPaperFunctionNames,
  type InternalFunctionBinding,
} from './work-paper-function-registry.js'
import {
  getRegisteredWorkPaperFunctionPluginById,
  getRegisteredWorkPaperFunctionPluginsById,
  readRegisteredWorkPaperLanguage,
} from './work-paper-static-registry.js'
import {
  DEFAULT_WORKPAPER_CONFIG,
  getAllWorkPaperStaticFunctionPlugins,
  getRegisteredWorkPaperStaticFunctionNames,
  getRegisteredWorkPaperStaticLanguageCodes,
  getWorkPaperStaticFunctionPlugin,
  getWorkPaperStaticLanguage,
  registerWorkPaperStaticFunction,
  registerWorkPaperStaticFunctionPlugin,
  registerWorkPaperStaticLanguage,
  unregisterAllWorkPaperStaticFunctions,
  unregisterWorkPaperStaticFunction,
  unregisterWorkPaperStaticFunctionPlugin,
  unregisterWorkPaperStaticLanguage,
  workPaperLanguages,
  WORKPAPER_BUILD_DATE,
  WORKPAPER_RELEASE_DATE,
  WORKPAPER_VERSION,
} from './work-paper-static-api.js'
import type {
  RawCellContent,
  SerializedWorkPaperNamedExpression,
  WorkPaperAddressMappingAdapter,
  WorkPaperArrayMappingAdapter,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperFormulaDiagnostic,
  WorkPaperColumnSearchAdapter,
  WorkPaperConfig,
  WorkPaperCalculationSettings,
  WorkPaperDateTime,
  WorkPaperDependencyGraphAdapter,
  WorkPaperEvaluatorAdapter,
  WorkPaperFunctionPluginDefinition,
  WorkPaperFunctionTranslationsPackage,
  WorkPaperGraphAdapter,
  WorkPaperLanguagePackage,
  WorkPaperLazilyTransformingAstServiceAdapter,
  WorkPaperLicenseKeyValidityState,
  WorkPaperRangeMappingAdapter,
  WorkPaperSheet,
  WorkPaperSheetMappingAdapter,
  WorkPaperEngineCounters,
  WorkPaperDetailedListener,
  WorkPaperEventName,
  WorkPaperInternals,
  WorkPaperListener,
} from './work-paper-types.js'
import { WorkPaperCapabilitySurface } from './work-paper-capability-surface.js'

export abstract class WorkPaperPublicSurface extends WorkPaperCapabilitySurface {
  static version = WORKPAPER_VERSION
  static buildDate = WORKPAPER_BUILD_DATE
  static releaseDate = WORKPAPER_RELEASE_DATE
  static readonly languages = workPaperLanguages
  static readonly defaultConfig: WorkPaperConfig = DEFAULT_WORKPAPER_CONFIG

  protected abstract readonly emitter: WorkPaperEmitter
  protected abstract readonly functionSnapshot: Map<string, InternalFunctionBinding>
  protected abstract readonly functionAliasLookup: Map<string, InternalFunctionBinding>
  protected abstract readonly internals: WorkPaperInternals
  protected abstract engine: SpreadsheetEngine
  protected abstract config: WorkPaperConfig

  protected abstract assertNotDisposed(): void
  protected abstract applyCalculationSettings(settings: WorkPaperCalculationSettings): void
  protected abstract createScratchWorkbook(config: WorkPaperConfig): {
    readonly engine: SpreadsheetEngine
    readonly registerNamedExpression: (expression: SerializedWorkPaperNamedExpression) => void
    readonly requireSheetId: (sheetName: string) => number
    readonly replaceSheetContent: (sheetId: number, sheet: WorkPaperSheet) => void
    readonly clearHistoryStacks: () => void
    readonly applyRawContent: (address: WorkPaperCellAddress, content: RawCellContent) => void
    readonly getRangeValues: (range: WorkPaperCellRange) => CellValue[][]
    readonly getCellValue: (address: WorkPaperCellAddress) => CellValue
    readonly dispose: () => void
  }
  protected abstract messageOf(error: unknown, fallback: string): string

  static getLanguage(languageCode: string): WorkPaperLanguagePackage {
    return getWorkPaperStaticLanguage(languageCode)
  }

  static registerLanguage(languageCode: string, languagePackage: WorkPaperLanguagePackage): void {
    registerWorkPaperStaticLanguage(languageCode, languagePackage)
  }

  static unregisterLanguage(languageCode: string): void {
    unregisterWorkPaperStaticLanguage(languageCode)
  }

  static getRegisteredLanguagesCodes(): string[] {
    return getRegisteredWorkPaperStaticLanguageCodes()
  }

  static registerFunctionPlugin(plugin: WorkPaperFunctionPluginDefinition, translations?: WorkPaperFunctionTranslationsPackage): void {
    registerWorkPaperStaticFunctionPlugin(plugin, translations)
  }

  static unregisterFunctionPlugin(plugin: WorkPaperFunctionPluginDefinition | string): void {
    unregisterWorkPaperStaticFunctionPlugin(plugin)
  }

  static registerFunction(
    functionId: string,
    plugin: WorkPaperFunctionPluginDefinition,
    translations?: WorkPaperFunctionTranslationsPackage,
  ): void {
    registerWorkPaperStaticFunction(functionId, plugin, translations)
  }

  static unregisterFunction(functionId: string): void {
    unregisterWorkPaperStaticFunction(functionId)
  }

  static unregisterAllFunctions(): void {
    unregisterAllWorkPaperStaticFunctions()
  }

  static getRegisteredFunctionNames(languageCode?: string): string[] {
    return getRegisteredWorkPaperStaticFunctionNames(languageCode)
  }

  static getFunctionPlugin(functionId: string): WorkPaperFunctionPluginDefinition | undefined {
    return getWorkPaperStaticFunctionPlugin(functionId)
  }

  static getAllFunctionPlugins(): WorkPaperFunctionPluginDefinition[] {
    return getAllWorkPaperStaticFunctionPlugins()
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

  getCalculationSettings(): WorkbookCalculationSettingsSnapshot {
    this.assertNotDisposed()
    return structuredClone(this.engine.getCalculationSettings())
  }

  setCalculationSettings(settings: WorkPaperCalculationSettings): void {
    this.assertNotDisposed()
    this.applyCalculationSettings(settings)
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

  getPerformanceCounters(): WorkPaperEngineCounters {
    this.assertNotDisposed()
    const counterAwareEngine = this.engine as SpreadsheetEngine & {
      getPerformanceCounters(): WorkPaperEngineCounters
    }
    return structuredClone(counterAwareEngine.getPerformanceCounters())
  }

  resetPerformanceCounters(): void {
    this.assertNotDisposed()
    const counterAwareEngine = this.engine as SpreadsheetEngine & {
      resetPerformanceCounters(): void
    }
    counterAwareEngine.resetPerformanceCounters()
  }

  normalizeFormula(formula: string): string {
    return normalizeWorkPaperFormula(formula, { messageOf: (error, fallback) => this.messageOf(error, fallback) })
  }

  calculateFormula(formula: string, scope?: number): CellValue | CellValue[][] {
    return calculateWorkPaperFormula({
      createWorkbook: (config) => this.createScratchWorkbook(config),
      config: this.getConfig(),
      serializedSheets: this.getAllSheetsSerialized(),
      namedExpressions: this.getAllNamedExpressionsSerialized(),
      formula,
      ...(scope !== undefined ? { scope } : {}),
      messageOf: (error, fallback) => this.messageOf(error, fallback),
    })
  }

  getCellDisplayValue(address: WorkPaperCellAddress): string {
    return formatCellDisplayValue(this.getCellValue(address), this.getCellValueFormat(address))
  }

  getCellFormulaDiagnostics(address: WorkPaperCellAddress): WorkPaperFormulaDiagnostic[] {
    return collectWorkPaperFormulaDiagnostics(address, {
      getCellValue: (target) => this.getCellValue(target),
      getCellValueFormat: (target) => this.getCellValueFormat(target),
      getCellFormula: (target) => this.getCellFormula(target),
      getRangeValues: (range) => this.getRangeValues(range),
      getSheetId: (sheetName) => this.sheetMapping.getSheetId(sheetName),
      getSheetName: (sheetId) => this.sheetMapping.getSheetName(sheetId),
      simpleCellAddressToString: (target, options) => this.simpleCellAddressToString(target, options),
      simpleCellRangeToString: (range, options) => this.simpleCellRangeToString(range, options),
    })
  }

  getNamedExpressionsFromFormula(formula: string): string[] {
    return getWorkPaperNamedExpressionsFromFormula(formula, { messageOf: (error, fallback) => this.messageOf(error, fallback) })
  }

  validateFormula(formula: string): boolean {
    return validateWorkPaperFormula(formula)
  }

  getRegisteredFunctionNames(languageCode?: string): string[] {
    const code = languageCode ?? this.config.language ?? DEFAULT_CONFIG.language ?? 'enGB'
    return listCapturedWorkPaperFunctionNames({
      functionSnapshot: this.functionSnapshot.values(),
      language: readRegisteredWorkPaperLanguage(code),
    })
  }

  getFunctionPlugin(functionId: string): WorkPaperFunctionPluginDefinition | undefined {
    return getCapturedWorkPaperFunctionPlugin({
      functionId,
      functionAliasLookup: this.functionAliasLookup,
      getPluginById: getRegisteredWorkPaperFunctionPluginById,
    })
  }

  getAllFunctionPlugins(): WorkPaperFunctionPluginDefinition[] {
    return getCapturedWorkPaperFunctionPlugins({
      functionSnapshot: this.functionSnapshot.values(),
      getPluginsById: getRegisteredWorkPaperFunctionPluginsById,
    })
  }

  numberToDateTime(value: number): WorkPaperDateTime | undefined {
    return numberToWorkPaperDateTime(value)
  }

  numberToDate(value: number): Omit<WorkPaperDateTime, 'hours' | 'minutes' | 'seconds'> | undefined {
    return numberToWorkPaperDate(value)
  }

  numberToTime(value: number): Pick<WorkPaperDateTime, 'hours' | 'minutes' | 'seconds'> | undefined {
    return numberToWorkPaperTime(value)
  }
}
