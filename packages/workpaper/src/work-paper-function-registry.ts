import type { CellValue } from '@bilig/protocol'
import type { EvaluationResult } from '@bilig/formula'
import { compareSheetNames } from './work-paper-sheet-inspection.js'
import type { WorkPaperFunctionPluginDefinition, WorkPaperLanguagePackage } from './work-paper-types.js'

export type WorkPaperFunctionImplementation = (...args: CellValue[]) => EvaluationResult | CellValue | undefined

export interface InternalFunctionBinding {
  pluginId: string
  publicName: string
  internalName: string
  implementation?: WorkPaperFunctionImplementation
}

export function listCapturedWorkPaperFunctionNames(args: {
  readonly functionSnapshot: Iterable<InternalFunctionBinding>
  readonly language: WorkPaperLanguagePackage | undefined
}): string[] {
  const functions = [...args.functionSnapshot]
    .filter((binding) => binding.publicName === binding.publicName.toUpperCase())
    .map((binding) => binding.publicName)
    .toSorted(compareSheetNames)
  if (!args.language?.functions) {
    return functions
  }
  return functions.map((name) => args.language?.functions?.[name] ?? name)
}

export function getCapturedWorkPaperFunctionPlugin(args: {
  readonly functionId: string
  readonly functionAliasLookup: ReadonlyMap<string, InternalFunctionBinding>
  readonly getPluginById: (pluginId: string) => WorkPaperFunctionPluginDefinition | undefined
}): WorkPaperFunctionPluginDefinition | undefined {
  const binding = args.functionAliasLookup.get(args.functionId.trim().toUpperCase())
  return binding ? args.getPluginById(binding.pluginId) : undefined
}

export function getCapturedWorkPaperFunctionPlugins(args: {
  readonly functionSnapshot: Iterable<InternalFunctionBinding>
  readonly getPluginsById: (pluginIds: Iterable<string>) => WorkPaperFunctionPluginDefinition[]
}): WorkPaperFunctionPluginDefinition[] {
  const pluginIds = new Set([...args.functionSnapshot].map((binding) => binding.pluginId))
  return args.getPluginsById(pluginIds)
}

export function captureWorkPaperFunctionRegistry(args: {
  readonly workbookId: number
  readonly configFunctionPlugins: readonly WorkPaperFunctionPluginDefinition[] | undefined
  readonly plugins: Iterable<WorkPaperFunctionPluginDefinition>
  readonly functionSnapshot: Map<string, InternalFunctionBinding>
  readonly functionAliasLookup: Map<string, InternalFunctionBinding>
  readonly internalFunctionLookup: Map<string, InternalFunctionBinding>
  readonly globalCustomFunctions: Map<string, WorkPaperFunctionImplementation>
}): void {
  const allowedPluginIds =
    args.configFunctionPlugins && args.configFunctionPlugins.length > 0
      ? new Set(args.configFunctionPlugins.map((plugin) => plugin.id))
      : undefined
  for (const plugin of args.plugins) {
    if (allowedPluginIds && !allowedPluginIds.has(plugin.id)) {
      continue
    }
    Object.keys(plugin.implementedFunctions).forEach((functionId) => {
      const normalized = functionId.trim().toUpperCase()
      const internalName = `__BILIG_WORKPAPER_FN_${args.workbookId}_${normalized}`
      const implementation = plugin.functions?.[normalized]
      const binding: InternalFunctionBinding = {
        pluginId: plugin.id,
        publicName: normalized,
        internalName,
      }
      if (implementation !== undefined) {
        binding.implementation = implementation
      }
      args.functionSnapshot.set(normalized, binding)
      args.functionAliasLookup.set(normalized, binding)
      args.internalFunctionLookup.set(internalName, binding)
      if (implementation) {
        args.globalCustomFunctions.set(internalName, implementation)
      }
    })
    Object.entries(plugin.aliases ?? {}).forEach(([alias, target]) => {
      const binding = args.functionSnapshot.get(target.trim().toUpperCase())
      if (!binding) {
        return
      }
      args.functionAliasLookup.set(alias.trim().toUpperCase(), binding)
    })
  }
}

export function clearWorkPaperFunctionBindings(args: {
  readonly functionSnapshot: Map<string, InternalFunctionBinding>
  readonly functionAliasLookup: Map<string, InternalFunctionBinding>
  readonly internalFunctionLookup: Map<string, InternalFunctionBinding>
  readonly globalCustomFunctions: Map<string, WorkPaperFunctionImplementation>
  readonly preserveInternalFunctionLookup?: boolean
}): void {
  args.internalFunctionLookup.forEach((_binding, internalName) => {
    args.globalCustomFunctions.delete(internalName)
  })
  args.functionSnapshot.clear()
  args.functionAliasLookup.clear()
  if (args.preserveInternalFunctionLookup !== true) {
    args.internalFunctionLookup.clear()
  }
}
