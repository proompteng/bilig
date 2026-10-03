import { describe, expect, it } from 'vitest'
import { MAX_ROWS, ValueTag } from '@bilig/protocol'
import {
  canApplyRuntimeOnlyWorkPaperConfigUpdate,
  canReuseWorkPaperSnapshotRebuild,
  checkWorkPaperLicenseKeyValidity,
  cloneConfig,
  clonePluginDefinition,
  DEFAULT_CONFIG,
  functionPluginIds,
  normalizeConfiguredWorkPaperCalculationSettings,
  validateWorkPaperConfig,
  WORKPAPER_CONFIG_KEYS,
  WORKPAPER_PUBLIC_ERROR_NAMES,
} from '../work-paper-config.js'
import {
  WorkPaperConfigValueTooBigError,
  WorkPaperConfigValueTooSmallError,
  WorkPaperExpectedOneOfValuesError,
  WorkPaperExpectedValueOfTypeError,
} from '../work-paper-errors.js'
import type { WorkPaperConfig, WorkPaperFunctionPluginDefinition } from '../work-paper-types.js'

function plugin(id: string): WorkPaperFunctionPluginDefinition {
  return {
    id,
    implementedFunctions: {
      TEST: {
        method: 'test',
        parameters: [{ argumentType: 'NUMBER' }],
      },
    },
    aliases: { ALIAS: 'TEST' },
    functions: {
      TEST: () => ({ tag: ValueTag.Number, value: 1 }),
    },
  }
}

describe('work paper config helpers', () => {
  it('deep-clones mutable config fields and plugin metadata', () => {
    const originalPlugin = plugin('plugin-b')
    const config = cloneConfig({
      ...DEFAULT_CONFIG,
      calculationSettings: { iterate: true, iterateCount: 32, iterateDelta: '0.01' },
      chooseAddressMappingPolicy: { mode: 'dense' },
      context: { nested: { value: 1 } },
      currencySymbol: ['$', 'EUR'],
      dateFormats: ['MM/DD/YYYY'],
      functionPlugins: [originalPlugin],
      nullDate: { year: 1900, month: 1, day: 1 },
      timeFormats: ['HH:mm'],
    })

    expect(config).not.toBe(DEFAULT_CONFIG)
    expect(config.calculationSettings).toEqual({ iterate: true, iterateCount: 32, iterateDelta: '0.01' })
    expect(config.chooseAddressMappingPolicy).toEqual({ mode: 'dense' })
    expect(config.calculationSettings).not.toBe(DEFAULT_CONFIG.calculationSettings)
    expect(config.chooseAddressMappingPolicy).not.toBe(DEFAULT_CONFIG.chooseAddressMappingPolicy)
    expect(config.context).toEqual({ nested: { value: 1 } })
    expect(config.currencySymbol).toEqual(['$', 'EUR'])
    expect(config.dateFormats).toEqual(['MM/DD/YYYY'])
    expect(config.nullDate).toEqual({ year: 1900, month: 1, day: 1 })
    expect(config.timeFormats).toEqual(['HH:mm'])
    expect(config.functionPlugins?.[0]).toEqual(originalPlugin)
    expect(config.functionPlugins?.[0]).not.toBe(originalPlugin)
    expect(config.functionPlugins?.[0]?.implementedFunctions.TEST).not.toBe(originalPlugin.implementedFunctions.TEST)
  })

  it('clones plugin definitions independently', () => {
    const originalPlugin = plugin('plugin-a')
    const cloned = clonePluginDefinition(originalPlugin)

    expect(cloned).toEqual(originalPlugin)
    expect(cloned).not.toBe(originalPlugin)
    expect(cloned.aliases).not.toBe(originalPlugin.aliases)
    expect(cloned.functions).not.toBe(originalPlugin.functions)
    expect(cloned.implementedFunctions.TEST).not.toBe(originalPlugin.implementedFunctions.TEST)
  })

  it('validates bounds, enums, callbacks, and structured-cloneable context', () => {
    const invalidDecimalConfig: WorkPaperConfig = {}
    Reflect.set(invalidDecimalConfig, 'decimalSeparator', ':')
    const invalidCallbackConfig: WorkPaperConfig = {}
    Reflect.set(invalidCallbackConfig, 'parseDateTime', 'nope')
    const invalidCalculationSettingsConfig: WorkPaperConfig = {
      calculationSettings: {
        iterateCount: 0,
      },
    }

    expect(() => validateWorkPaperConfig({ maxRows: 0 })).toThrow(WorkPaperConfigValueTooSmallError)
    expect(() => validateWorkPaperConfig({ maxRows: MAX_ROWS + 1 })).toThrow(WorkPaperConfigValueTooBigError)
    expect(() => validateWorkPaperConfig(invalidDecimalConfig)).toThrow(WorkPaperExpectedOneOfValuesError)
    expect(() => validateWorkPaperConfig(invalidCallbackConfig)).toThrow(WorkPaperExpectedValueOfTypeError)
    expect(() => validateWorkPaperConfig(invalidCalculationSettingsConfig)).toThrow(WorkPaperConfigValueTooSmallError)
    expect(() => validateWorkPaperConfig({ context: () => undefined })).toThrow(WorkPaperExpectedValueOfTypeError)
    expect(() =>
      validateWorkPaperConfig({
        calculationSettings: {
          mode: 'manual',
          compatibilityMode: 'odf-1.4',
          dateSystem: '1904',
          iterate: true,
          iterateCount: 50,
          iterateDelta: '0.001',
          fullPrecision: false,
          fullCalcOnLoad: true,
          calcOnSave: true,
          calcCompleted: false,
          concurrentCalc: false,
        },
      }),
    ).not.toThrow()
    expect(() => validateWorkPaperConfig({ chooseAddressMappingPolicy: { mode: 'sparse' } })).not.toThrow()
  })

  it('normalizes configured calculation settings against workbook defaults', () => {
    expect(normalizeConfiguredWorkPaperCalculationSettings({ iterate: true })).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      iterate: true,
    })
    expect(normalizeConfiguredWorkPaperCalculationSettings({ calcOnSave: true, calcCompleted: false })).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      calcOnSave: true,
      calcCompleted: false,
    })
    expect(
      normalizeConfiguredWorkPaperCalculationSettings(
        { iterateCount: 250, iterateDelta: '0.0001' },
        { mode: 'manual', compatibilityMode: 'odf-1.4', iterate: true },
      ),
    ).toEqual({
      mode: 'manual',
      compatibilityMode: 'odf-1.4',
      iterate: true,
      iterateCount: 250,
      iterateDelta: '0.0001',
    })
  })

  it('reports license states and sorted plugin ids', () => {
    expect(checkWorkPaperLicenseKeyValidity(undefined)).toBe('missing')
    expect(checkWorkPaperLicenseKeyValidity('')).toBe('missing')
    expect(checkWorkPaperLicenseKeyValidity('internal')).toBe('valid')
    expect(checkWorkPaperLicenseKeyValidity('custom')).toBe('invalid')

    expect(functionPluginIds({ functionPlugins: [plugin('z'), plugin('a')] })).toEqual(['a', 'z'])
  })

  it('detects snapshot-rebuild reuse and runtime-only config changes', () => {
    expect(
      canReuseWorkPaperSnapshotRebuild(
        { language: 'enGB', functionPlugins: [plugin('z'), plugin('a')] },
        { language: 'enGB', functionPlugins: [plugin('a'), plugin('z')] },
      ),
    ).toBe(true)
    expect(canReuseWorkPaperSnapshotRebuild({ language: 'enGB' }, { language: 'deDE' })).toBe(false)
    expect(canReuseWorkPaperSnapshotRebuild({ language: 'enGB', functionPlugins: [plugin('a')] }, { language: 'enGB' })).toBe(false)
    expect(canApplyRuntimeOnlyWorkPaperConfigUpdate(['useColumnIndex', 'useStats'])).toBe(true)
    expect(canApplyRuntimeOnlyWorkPaperConfigUpdate(['useColumnIndex', 'language'])).toBe(false)
  })

  it('keeps runtime config and public error allowlists visible', () => {
    expect(DEFAULT_CONFIG.useColumnIndex).toBe(true)
    expect(WORKPAPER_CONFIG_KEYS).toContain('calculationSettings')
    expect(WORKPAPER_CONFIG_KEYS).toContain('functionPlugins')
    expect(WORKPAPER_CONFIG_KEYS).toContain('maxRows')
    expect(WORKPAPER_PUBLIC_ERROR_NAMES.has('WorkPaperOperationError')).toBe(true)
    expect(WORKPAPER_PUBLIC_ERROR_NAMES.has('InternalImplementationError')).toBe(false)
  })
})
