import { describe, expect, it } from 'vitest'
import { compileFormula, type CompiledFormula } from '@bilig/formula'
import type { RuntimeFormula, RuntimeStructuralFormulaSourceTransform } from '../engine/runtime-state.js'
import { getRuntimeFormulaSource, getRuntimeFormulaStructuralCompiled } from '../engine/runtime-formula-source.js'

function makeRuntimeFormula(
  source: string,
  compiled: CompiledFormula,
  structuralSourceTransform: RuntimeStructuralFormulaSourceTransform | undefined = undefined,
): RuntimeFormula {
  return {
    cellIndex: 0,
    formulaSlotId: 0,
    planId: 0,
    templateId: undefined,
    source,
    compiled,
    plan: { id: 0, source, compiled },
    dependencyIndices: new Uint32Array(0),
    dependencyEntities: { ptr: -1, len: 0, cap: 0 },
    rangeDependencies: new Uint32Array(0),
    runtimeProgram: compiled.program,
    constants: compiled.constants,
    structuralSourceTransform,
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

describe('runtime formula source helpers', () => {
  it('returns the stored source and no compiled override when no structural transform is pending', () => {
    const compiled = compileFormula('A1+1')
    const formula = makeRuntimeFormula('A1+1', compiled)

    expect(getRuntimeFormulaSource(formula)).toBe('A1+1')
    expect(getRuntimeFormulaStructuralCompiled(formula)).toBeUndefined()
  })

  it('materializes deferred structural formula source and compiled metadata', () => {
    const compiled = compileFormula('SUM(A1:A10)')
    const formula = makeRuntimeFormula('SUM(A1:A10)', compiled, {
      ownerSheetName: 'Sheet1',
      targetSheetName: 'Sheet1',
      preservesValue: true,
      transform: {
        kind: 'insert',
        axis: 'row',
        start: 5,
        count: 1,
      },
    })

    const rewrittenCompiled = getRuntimeFormulaStructuralCompiled(formula)

    expect(getRuntimeFormulaSource(formula)).toBe('SUM(A1:A11)')
    expect(rewrittenCompiled).toBeDefined()
    expect(rewrittenCompiled?.deps).toEqual(['A1:A11'])
    expect(rewrittenCompiled?.program).toBe(compiled.program)
  })
})
