import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import {
  captureWorkPaperFunctionRegistry,
  clearWorkPaperFunctionBindings,
  getCapturedWorkPaperFunctionPlugin,
  getCapturedWorkPaperFunctionPlugins,
  listCapturedWorkPaperFunctionNames,
  type InternalFunctionBinding,
  type WorkPaperFunctionImplementation,
} from '../work-paper-function-registry.js'
import type { WorkPaperFunctionPluginDefinition } from '../work-paper-types.js'

function plugin(id: string, functions: Record<string, WorkPaperFunctionImplementation>): WorkPaperFunctionPluginDefinition {
  return {
    id,
    implementedFunctions: Object.fromEntries(Object.keys(functions).map((name) => [name, { method: name }])),
    functions,
    aliases: { Alias: Object.keys(functions)[0] ?? 'MISSING' },
  }
}

describe('work paper function registry helpers', () => {
  it('captures allowed plugin functions with aliases and internal custom names', () => {
    const functionSnapshot = new Map<string, InternalFunctionBinding>()
    const functionAliasLookup = new Map<string, InternalFunctionBinding>()
    const internalFunctionLookup = new Map<string, InternalFunctionBinding>()
    const globalCustomFunctions = new Map<string, WorkPaperFunctionImplementation>()
    const addOne: WorkPaperFunctionImplementation = () => ({ tag: ValueTag.Number, value: 2 })
    const ignored: WorkPaperFunctionImplementation = () => ({ tag: ValueTag.Number, value: 99 })

    captureWorkPaperFunctionRegistry({
      workbookId: 42,
      configFunctionPlugins: [plugin('math', { ADDONE: addOne })],
      plugins: [plugin('math', { ADDONE: addOne }), plugin('ignored', { IGNORED: ignored })],
      functionSnapshot,
      functionAliasLookup,
      internalFunctionLookup,
      globalCustomFunctions,
    })

    expect([...functionSnapshot.keys()]).toEqual(['ADDONE'])
    expect(functionAliasLookup.get('ALIAS')?.publicName).toBe('ADDONE')
    expect(internalFunctionLookup.get('__BILIG_WORKPAPER_FN_42_ADDONE')?.pluginId).toBe('math')
    expect(globalCustomFunctions.get('__BILIG_WORKPAPER_FN_42_ADDONE')).toBe(addOne)
    expect(globalCustomFunctions.has('__BILIG_WORKPAPER_FN_42_IGNORED')).toBe(false)
  })

  it('clears captured bindings and removes global custom functions', () => {
    const functionSnapshot = new Map<string, InternalFunctionBinding>()
    const functionAliasLookup = new Map<string, InternalFunctionBinding>()
    const internalFunctionLookup = new Map<string, InternalFunctionBinding>()
    const globalCustomFunctions = new Map<string, WorkPaperFunctionImplementation>()
    const addOne: WorkPaperFunctionImplementation = () => ({ tag: ValueTag.Number, value: 2 })
    captureWorkPaperFunctionRegistry({
      workbookId: 42,
      configFunctionPlugins: undefined,
      plugins: [plugin('math', { ADDONE: addOne })],
      functionSnapshot,
      functionAliasLookup,
      internalFunctionLookup,
      globalCustomFunctions,
    })

    clearWorkPaperFunctionBindings({ functionSnapshot, functionAliasLookup, internalFunctionLookup, globalCustomFunctions })

    expect(functionSnapshot.size).toBe(0)
    expect(functionAliasLookup.size).toBe(0)
    expect(internalFunctionLookup.size).toBe(0)
    expect(globalCustomFunctions.size).toBe(0)
  })

  it('can preserve internal names for public formula restoration while removing active implementations', () => {
    const functionSnapshot = new Map<string, InternalFunctionBinding>()
    const functionAliasLookup = new Map<string, InternalFunctionBinding>()
    const internalFunctionLookup = new Map<string, InternalFunctionBinding>()
    const globalCustomFunctions = new Map<string, WorkPaperFunctionImplementation>()
    const addOne: WorkPaperFunctionImplementation = () => ({ tag: ValueTag.Number, value: 2 })
    captureWorkPaperFunctionRegistry({
      workbookId: 42,
      configFunctionPlugins: undefined,
      plugins: [plugin('math', { ADDONE: addOne })],
      functionSnapshot,
      functionAliasLookup,
      internalFunctionLookup,
      globalCustomFunctions,
    })

    clearWorkPaperFunctionBindings({
      functionSnapshot,
      functionAliasLookup,
      internalFunctionLookup,
      globalCustomFunctions,
      preserveInternalFunctionLookup: true,
    })

    expect(functionSnapshot.size).toBe(0)
    expect(functionAliasLookup.size).toBe(0)
    expect(internalFunctionLookup.get('__BILIG_WORKPAPER_FN_42_ADDONE')?.publicName).toBe('ADDONE')
    expect(globalCustomFunctions.size).toBe(0)
  })

  it('lists captured function names with optional language translations', () => {
    const functionSnapshot = new Map<string, InternalFunctionBinding>([
      ['BETA', { pluginId: 'math', publicName: 'BETA', internalName: '__BETA' }],
      ['ALPHA', { pluginId: 'math', publicName: 'ALPHA', internalName: '__ALPHA' }],
    ])

    expect(listCapturedWorkPaperFunctionNames({ functionSnapshot: functionSnapshot.values(), language: undefined })).toEqual([
      'ALPHA',
      'BETA',
    ])
    expect(
      listCapturedWorkPaperFunctionNames({
        functionSnapshot: functionSnapshot.values(),
        language: { functions: { ALPHA: 'ALFA' } },
      }),
    ).toEqual(['ALFA', 'BETA'])
  })

  it('resolves captured function plugins through public callbacks', () => {
    const functionSnapshot = new Map<string, InternalFunctionBinding>([
      ['ADDONE', { pluginId: 'math', publicName: 'ADDONE', internalName: '__ADDONE' }],
      ['LOOKUPX', { pluginId: 'lookup', publicName: 'LOOKUPX', internalName: '__LOOKUPX' }],
    ])
    const functionAliasLookup = new Map<string, InternalFunctionBinding>([
      ['ALIAS', { pluginId: 'math', publicName: 'ADDONE', internalName: '__ADDONE' }],
    ])
    const mathPlugin = plugin('math', {})
    const lookupPlugin = plugin('lookup', {})

    expect(
      getCapturedWorkPaperFunctionPlugin({
        functionId: 'alias',
        functionAliasLookup,
        getPluginById: (pluginId) => (pluginId === 'math' ? mathPlugin : undefined),
      }),
    ).toBe(mathPlugin)
    expect(
      getCapturedWorkPaperFunctionPlugin({
        functionId: 'missing',
        functionAliasLookup,
        getPluginById: () => lookupPlugin,
      }),
    ).toBeUndefined()
    expect(
      getCapturedWorkPaperFunctionPlugins({
        functionSnapshot: functionSnapshot.values(),
        getPluginsById: (pluginIds) => [...pluginIds].map((pluginId) => (pluginId === 'math' ? mathPlugin : lookupPlugin)),
      }).map((registeredPlugin) => registeredPlugin.id),
    ).toEqual(['math', 'lookup'])
  })
})
