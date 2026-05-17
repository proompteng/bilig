import { describe, expect, it, vi } from 'vitest'
import { compileFormula } from '@bilig/formula'
import { createFormulaFamilyStore } from '../formula/formula-family-store.js'
import { createFormulaBindingFamilyIndexController } from '../engine/services/formula-binding-family-index-controller.js'
import type { RuntimeFormula } from '../engine/runtime-state.js'

function deferredRun(overrides: { cellIndices: number[]; start: number; lastIndex: number }) {
  return {
    sheetId: 1,
    templateId: 7,
    shapeKey: 'relative-add',
    axis: 'row' as const,
    fixedIndex: 4,
    step: 1,
    ordered: true,
    ...overrides,
  }
}

function structuralSourceTransform() {
  return {
    ownerSheetName: 'Summary',
    targetSheetName: 'Data',
    transform: { kind: 'insert' as const, axis: 'row' as const, start: 2, count: 1 },
    preservesValue: true,
  }
}

function runtimeFormulaFixture(): RuntimeFormula {
  const compiled = compileFormula('A1+1')
  return {
    cellIndex: 99,
    formulaSlotId: 0,
    planId: 1,
    templateId: undefined,
    source: 'A1+1',
    compiled,
    plan: { id: 1, source: 'A1+1', compiled },
    dependencyIndices: new Uint32Array(),
    dependencyEntities: { ptr: 0, len: 0, cap: 0 },
    rangeDependencies: new Uint32Array(),
    graphRangeDependencies: new Uint32Array(),
    runtimeProgram: compiled.program,
    constants: compiled.constants,
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: compiled.program.length,
    constNumberOffset: 0,
    constNumberLength: compiled.constants.length,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: undefined,
    directScalar: undefined,
    directCriteria: undefined,
  }
}

describe('formula binding family index controller', () => {
  it('replays deferred runs with structural source transforms on first family lookup', () => {
    const formulaFamilies = createFormulaFamilyStore()
    const shapeKeyCache = new Map<string, string>()
    const registerFormulaFamilyInStoreNow = vi.fn()
    const controller = createFormulaBindingFamilyIndexController({
      formulaFamilies,
      formulaFamilyShapeKeyCache: shapeKeyCache,
      registerFormulaFamilyInStoreNow,
      countFormulaSheetMembersNow: (sheetId) => (sheetId === 1 ? 3 : 0),
      rebuildFormulaFamilyIndexNow: vi.fn(),
    })
    const transform = structuralSourceTransform()

    controller.deferRunsNow([
      deferredRun({ cellIndices: [10, 11], start: 0, lastIndex: 1 }),
      deferredRun({ cellIndices: [12], start: 2, lastIndex: 2 }),
    ])

    expect(controller.canUseNow()).toBe(true)
    expect(controller.isReadyNow()).toBe(false)
    expect(controller.tryDeferStructuralSourceTransformsNow(1, transform, () => true)).toBe(3)
    expect(controller.hasStructuralSourceTransformsNow()).toBe(true)

    expect(controller.getStructuralSourceTransformNow(10)).toEqual(transform)

    expect(controller.isReadyNow()).toBe(true)
    expect(registerFormulaFamilyInStoreNow).not.toHaveBeenCalled()
    expect(formulaFamilies.getStats()).toEqual({ familyCount: 1, runCount: 1, memberCount: 3 })
    expect(controller.consumeStructuralSourceTransformsNow()).toEqual([{ cellIndices: [10, 11, 12], transform }])
    expect(controller.hasStructuralSourceTransformsNow()).toBe(false)
  })

  it('invalidates deferred replay state when later mutations make the run snapshot untrusted', () => {
    const formulaFamilies = createFormulaFamilyStore()
    const shapeKeyCache = new Map<string, string>([['stale', 'key']])
    const registerFormulaFamilyInStoreNow = vi.fn()
    const controller = createFormulaBindingFamilyIndexController({
      formulaFamilies,
      formulaFamilyShapeKeyCache: shapeKeyCache,
      registerFormulaFamilyInStoreNow,
      countFormulaSheetMembersNow: () => 1,
      rebuildFormulaFamilyIndexNow: vi.fn(),
    })

    controller.deferRunsNow([deferredRun({ cellIndices: [20], start: 0, lastIndex: 0 })])
    expect(controller.tryDeferStructuralSourceTransformsNow(1, structuralSourceTransform(), () => true)).toBe(1)

    controller.registerFormulaFamilyNow(99, runtimeFormulaFixture())

    expect(controller.isReadyNow()).toBe(false)
    expect(controller.canUseNow()).toBe(false)
    expect(controller.tryDeferStructuralSourceTransformsNow(1, structuralSourceTransform(), () => true)).toBeUndefined()
    expect(controller.hasStructuralSourceTransformsNow()).toBe(false)
    expect(shapeKeyCache.size).toBe(0)
    expect(registerFormulaFamilyInStoreNow).not.toHaveBeenCalled()
  })
})
