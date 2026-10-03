import { cloneConfig, DEFAULT_CONFIG } from './work-paper-config.js'
import {
  getAllRegisteredWorkPaperFunctionPlugins,
  getRegisteredWorkPaperFunctionNames,
  getRegisteredWorkPaperFunctionPlugin,
  getRegisteredWorkPaperLanguage,
  getRegisteredWorkPaperLanguageCodes,
  registerWorkPaperFunction,
  registerWorkPaperFunctionPlugin,
  registerWorkPaperLanguage,
  unregisterAllWorkPaperFunctions,
  unregisterWorkPaperFunction,
  unregisterWorkPaperFunctionPlugin,
  unregisterWorkPaperLanguage,
} from './work-paper-static-registry.js'
import type {
  WorkPaperConfig,
  WorkPaperFunctionPluginDefinition,
  WorkPaperFunctionTranslationsPackage,
  WorkPaperLanguagePackage,
} from './work-paper-types.js'
export { WORKPAPER_VERSION } from './work-paper-version.js'

export const WORKPAPER_BUILD_DATE = '2026-04-10'
export const WORKPAPER_RELEASE_DATE = '2026-04-10'
export const workPaperLanguages: Record<string, WorkPaperLanguagePackage> = {}
export const DEFAULT_WORKPAPER_CONFIG: WorkPaperConfig = cloneConfig(DEFAULT_CONFIG)

export function getWorkPaperStaticLanguage(languageCode: string): WorkPaperLanguagePackage {
  return getRegisteredWorkPaperLanguage(languageCode)
}

export function registerWorkPaperStaticLanguage(languageCode: string, languagePackage: WorkPaperLanguagePackage): void {
  registerWorkPaperLanguage(workPaperLanguages, languageCode, languagePackage)
}

export function unregisterWorkPaperStaticLanguage(languageCode: string): void {
  unregisterWorkPaperLanguage(workPaperLanguages, languageCode)
}

export function getRegisteredWorkPaperStaticLanguageCodes(): string[] {
  return getRegisteredWorkPaperLanguageCodes()
}

export function registerWorkPaperStaticFunctionPlugin(
  plugin: WorkPaperFunctionPluginDefinition,
  translations?: WorkPaperFunctionTranslationsPackage,
): void {
  registerWorkPaperFunctionPlugin(workPaperLanguages, plugin, translations)
}

export function unregisterWorkPaperStaticFunctionPlugin(plugin: WorkPaperFunctionPluginDefinition | string): void {
  unregisterWorkPaperFunctionPlugin(plugin)
}

export function registerWorkPaperStaticFunction(
  functionId: string,
  plugin: WorkPaperFunctionPluginDefinition,
  translations?: WorkPaperFunctionTranslationsPackage,
): void {
  registerWorkPaperFunction(workPaperLanguages, functionId, plugin, translations)
}

export function unregisterWorkPaperStaticFunction(functionId: string): void {
  unregisterWorkPaperFunction(functionId)
}

export function unregisterAllWorkPaperStaticFunctions(): void {
  unregisterAllWorkPaperFunctions()
}

export function getRegisteredWorkPaperStaticFunctionNames(languageCode?: string): string[] {
  return getRegisteredWorkPaperFunctionNames(languageCode)
}

export function getWorkPaperStaticFunctionPlugin(functionId: string): WorkPaperFunctionPluginDefinition | undefined {
  return getRegisteredWorkPaperFunctionPlugin(functionId)
}

export function getAllWorkPaperStaticFunctionPlugins(): WorkPaperFunctionPluginDefinition[] {
  return getAllRegisteredWorkPaperFunctionPlugins()
}
