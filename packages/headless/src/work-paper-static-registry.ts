import { ErrorCode, type CellValue } from '@bilig/protocol'
import { installExternalFunctionAdapter } from '@bilig/formula/external-function-adapter'
import {
  WorkPaperFunctionPluginValidationError,
  WorkPaperLanguageAlreadyRegisteredError,
  WorkPaperLanguageNotRegisteredError,
} from './work-paper-errors.js'
import { clonePluginDefinition } from './work-paper-config.js'
import { compareSheetNames } from './work-paper-sheet-inspection.js'
import { errorValue, scalarFromResult } from './work-paper-runtime-helpers.js'
import type { WorkPaperFunctionImplementation } from './work-paper-function-registry.js'
import type {
  WorkPaperFunctionPluginDefinition,
  WorkPaperFunctionTranslationsPackage,
  WorkPaperLanguagePackage,
} from './work-paper-types.js'

export const workPaperLanguageRegistry = new Map<string, WorkPaperLanguagePackage>()
export const workPaperFunctionPluginRegistry = new Map<string, WorkPaperFunctionPluginDefinition>()
export const workPaperGlobalCustomFunctions = new Map<string, WorkPaperFunctionImplementation>()

let customAdapterInstalled = false

export function ensureWorkPaperCustomAdapterInstalled(): void {
  if (customAdapterInstalled) {
    return
  }
  installExternalFunctionAdapter({
    surface: 'host',
    resolveFunction(name) {
      const implementation = workPaperGlobalCustomFunctions.get(name.trim().toUpperCase())
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

export function hasRegisteredWorkPaperFunctionPlugins(): boolean {
  return workPaperFunctionPluginRegistry.size > 0
}

export function getRegisteredWorkPaperLanguage(languageCode: string): WorkPaperLanguagePackage {
  const language = workPaperLanguageRegistry.get(languageCode)
  if (!language) {
    throw new WorkPaperLanguageNotRegisteredError(languageCode)
  }
  return structuredClone(language)
}

export function readRegisteredWorkPaperLanguage(languageCode: string): WorkPaperLanguagePackage | undefined {
  return workPaperLanguageRegistry.get(languageCode)
}

export function registerWorkPaperLanguage(
  languages: Record<string, WorkPaperLanguagePackage>,
  languageCode: string,
  languagePackage: WorkPaperLanguagePackage,
): void {
  if (workPaperLanguageRegistry.has(languageCode)) {
    throw new WorkPaperLanguageAlreadyRegisteredError(languageCode)
  }
  workPaperLanguageRegistry.set(languageCode, structuredClone(languagePackage))
  languages[languageCode] = structuredClone(languagePackage)
}

export function unregisterWorkPaperLanguage(languages: Record<string, WorkPaperLanguagePackage>, languageCode: string): void {
  if (!workPaperLanguageRegistry.delete(languageCode)) {
    throw new WorkPaperLanguageNotRegisteredError(languageCode)
  }
  delete languages[languageCode]
}

export function getRegisteredWorkPaperLanguageCodes(): string[] {
  return [...workPaperLanguageRegistry.keys()].toSorted(compareSheetNames)
}

export function registerWorkPaperFunctionPlugin(
  languages: Record<string, WorkPaperLanguagePackage>,
  plugin: WorkPaperFunctionPluginDefinition,
  translations?: WorkPaperFunctionTranslationsPackage,
): void {
  workPaperFunctionPluginRegistry.set(plugin.id, clonePluginDefinition(plugin))
  if (translations) {
    loadWorkPaperFunctionTranslations(languages, translations)
  }
}

export function unregisterWorkPaperFunctionPlugin(plugin: WorkPaperFunctionPluginDefinition | string): void {
  const pluginId = typeof plugin === 'string' ? plugin : plugin.id
  workPaperFunctionPluginRegistry.delete(pluginId)
}

export function registerWorkPaperFunction(
  languages: Record<string, WorkPaperLanguagePackage>,
  functionId: string,
  plugin: WorkPaperFunctionPluginDefinition,
  translations?: WorkPaperFunctionTranslationsPackage,
): void {
  const existing = workPaperFunctionPluginRegistry.get(plugin.id)
  const nextPlugin = clonePluginDefinition(existing ?? plugin)
  if (!nextPlugin.implementedFunctions[functionId]) {
    throw WorkPaperFunctionPluginValidationError.functionNotDeclaredInPlugin(functionId, plugin.id)
  }
  workPaperFunctionPluginRegistry.set(nextPlugin.id, nextPlugin)
  if (translations) {
    loadWorkPaperFunctionTranslations(languages, translations)
  }
}

export function unregisterWorkPaperFunction(functionId: string): void {
  const normalized = functionId.trim().toUpperCase()
  workPaperFunctionPluginRegistry.forEach((plugin, pluginId) => {
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
    workPaperFunctionPluginRegistry.set(pluginId, nextPlugin)
  })
}

export function unregisterAllWorkPaperFunctions(): void {
  workPaperFunctionPluginRegistry.clear()
}

export function getRegisteredWorkPaperFunctionNames(languageCode?: string): string[] {
  const normalized = languageCode ?? 'enGB'
  const language = workPaperLanguageRegistry.get(normalized)
  const functions = [...workPaperFunctionPluginRegistry.values()].flatMap((plugin) => Object.keys(plugin.implementedFunctions))
  if (!language?.functions) {
    return functions.toSorted(compareSheetNames)
  }
  return functions.map((name) => language.functions?.[name] ?? name).toSorted(compareSheetNames)
}

export function getRegisteredWorkPaperFunctionPlugin(functionId: string): WorkPaperFunctionPluginDefinition | undefined {
  const normalized = functionId.trim().toUpperCase()
  const plugin = [...workPaperFunctionPluginRegistry.values()].find(
    (candidate) => candidate.implementedFunctions[normalized] !== undefined || candidate.aliases?.[normalized] !== undefined,
  )
  return plugin ? clonePluginDefinition(plugin) : undefined
}

export function getAllRegisteredWorkPaperFunctionPlugins(): WorkPaperFunctionPluginDefinition[] {
  return [...workPaperFunctionPluginRegistry.values()].map((plugin) => clonePluginDefinition(plugin))
}

export function getRegisteredWorkPaperFunctionPluginById(pluginId: string): WorkPaperFunctionPluginDefinition | undefined {
  const plugin = workPaperFunctionPluginRegistry.get(pluginId)
  return plugin ? clonePluginDefinition(plugin) : undefined
}

export function getRegisteredWorkPaperFunctionPluginsById(pluginIds: Iterable<string>): WorkPaperFunctionPluginDefinition[] {
  return [...pluginIds]
    .map((pluginId) => getRegisteredWorkPaperFunctionPluginById(pluginId))
    .filter((plugin): plugin is WorkPaperFunctionPluginDefinition => plugin !== undefined)
}

function loadWorkPaperFunctionTranslations(
  languages: Record<string, WorkPaperLanguagePackage>,
  translations: WorkPaperFunctionTranslationsPackage,
): void {
  Object.entries(translations).forEach(([languageCode, functionTranslations]) => {
    const existing = workPaperLanguageRegistry.get(languageCode)
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
    workPaperLanguageRegistry.set(languageCode, nextLanguage)
    languages[languageCode] = structuredClone(nextLanguage)
  })
}
