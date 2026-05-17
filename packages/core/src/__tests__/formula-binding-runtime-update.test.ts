import { describe, expect, it } from 'vitest'
import { compileFormula } from '@bilig/formula'
import { applyFormulaRuntimePlanFields } from '../engine/services/formula-binding-runtime-update.js'
import type { RuntimeFormula } from '../engine/runtime-state.js'

function runtimeFormulaFixture(): RuntimeFormula {
  const existingCompiled = compileFormula('A1+1')
  return {
    cellIndex: 7,
    formulaSlotId: 3,
    planId: 12,
    templateId: 44,
    source: 'Old!A1+1',
    compiled: existingCompiled,
    plan: { id: 12, source: 'Old!A1+1', compiled: existingCompiled, templateId: 44, refCount: 1 },
    dependencyIndices: new Uint32Array([1]),
    dependencyEntities: { ptr: 0, len: 0, cap: 0 },
    rangeDependencies: new Uint32Array([2]),
    graphRangeDependencies: new Uint32Array([3]),
    runtimeProgram: new Uint32Array([21, 22]),
    constants: existingCompiled.constants,
    structuralSourceTransform: {
      ownerSheetName: 'Summary',
      targetSheetName: 'Old',
      transform: { kind: 'insert', axis: 'row', start: 0, count: 1 },
      preservesValue: true,
    },
    sourceRenameTransforms: [{ oldSheetName: 'Old', newSheetName: 'New' }],
    programOffset: 0,
    programLength: 2,
    constNumberOffset: 0,
    constNumberLength: 1,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: undefined,
    directScalar: undefined,
    directCriteria: undefined,
  }
}

describe('formula binding runtime update helpers', () => {
  it('updates shared runtime plan fields and clears deferred source transforms', () => {
    const formula = runtimeFormulaFixture()
    const compiled = compileFormula('A1+2')
    const plan = { id: 99, source: 'New!A1+1', compiled, templateId: 55, refCount: 1 }
    const runtimeProgram = new Uint32Array([41, 42, 43])

    applyFormulaRuntimePlanFields(formula, {
      source: 'New!A1+1',
      plan,
      templateId: 55,
      runtimeProgram,
      programLength: runtimeProgram.length,
    })

    expect(formula).toMatchObject({
      source: 'New!A1+1',
      planId: 99,
      templateId: 55,
      plan,
      constants: compiled.constants,
      programLength: 3,
      constNumberLength: compiled.constants.length,
      structuralSourceTransform: undefined,
      sourceRenameTransforms: undefined,
    })
    expect(formula.runtimeProgram).toBe(runtimeProgram)
  })
})
