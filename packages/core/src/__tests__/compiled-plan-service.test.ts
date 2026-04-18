import { describe, expect, it } from 'vitest'
import { compileFormula, compileFormulaAst, parseFormula } from '@bilig/formula'
import { createEngineCompiledPlanService } from '../engine/services/compiled-plan-service.js'

describe('EngineCompiledPlanService', () => {
  it('reuses one immutable compiled plan for the same compiled object identity and releases it by refcount', () => {
    const service = createEngineCompiledPlanService()
    const shared = compileFormula('1+2')

    const first = service.intern('1+2', shared)
    const second = service.intern('1+2', shared)
    const different = service.intern('2+3', compileFormula('2+3'))

    expect(second.id).toBe(first.id)
    expect(second.compiled).toBe(first.compiled)
    expect(different.id).not.toBe(first.id)

    service.release(first.id)
    expect(service.get(first.id)).toBeDefined()

    service.release(second.id)
    expect(service.get(first.id)).toBeUndefined()
    expect(service.get(different.id)).toBeDefined()
  })

  it('does not merge distinct compiled plans that happen to share a source string', () => {
    const service = createEngineCompiledPlanService()

    const first = service.intern('Shared', compileFormulaAst('Shared', parseFormula('1+2')))
    const second = service.intern('Shared', compileFormulaAst('Shared', parseFormula('2+3')))

    expect(second.id).not.toBe(first.id)
    expect([...second.compiled.constants]).not.toEqual([...first.compiled.constants])
  })

  it('interns a new plan when replace targets a missing plan id', () => {
    const service = createEngineCompiledPlanService()
    const compiled = compileFormula('3+4')

    const record = service.replace(999_999, '3+4', compiled)

    expect(record.source).toBe('3+4')
    expect(record.compiled).toBe(compiled)
    expect(service.get(record.id)).toBe(record)
  })

  it('treats release on a missing plan id as a no-op', () => {
    const service = createEngineCompiledPlanService()
    const record = service.intern('5+6', compileFormula('5+6'))

    service.release(record.id + 1)

    expect(service.get(record.id)).toBe(record)
  })

  it('tracks template ownership on compiled plans and clears all plan records', () => {
    const service = createEngineCompiledPlanService()
    const compiled = compileFormula('A1+B1')

    const record = service.intern('A1+B1', compiled, 7)
    expect(record.templateId).toBe(7)

    service.clear()

    expect(service.get(record.id)).toBeUndefined()
  })
})
