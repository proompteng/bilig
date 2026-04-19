import { describe, expect, it } from 'vitest'
import { FormulaMode } from '@bilig/protocol'
import { createEngineFormulaTemplateNormalizationService } from '../engine/services/formula-template-normalization-service.js'
import { createEngineCounters } from '../perf/engine-counters.js'

describe('EngineFormulaTemplateNormalizationService', () => {
  it('compiles repeated row-shifted templates once and translates later instances', () => {
    const service = createEngineFormulaTemplateNormalizationService()

    const first = service.compileForCell('A1+B1', 0, 2)
    const second = service.compileForCell('A2+B2', 1, 2)
    const third = service.compileForCell('A3+B3', 2, 2)

    expect(first.deps).toEqual(['A1', 'B1'])
    expect(second.deps).toEqual(['A2', 'B2'])
    expect(third.deps).toEqual(['A3', 'B3'])
    expect(second.ast).toBe(first.ast)
    expect(third.ast).toBe(first.ast)
    expect(second.astMatchesSource).toBe(false)
    expect(third.astMatchesSource).toBe(false)
  })

  it('keeps distinct template families separate', () => {
    const service = createEngineFormulaTemplateNormalizationService()

    const add = service.compileForCell('A1+B1', 0, 2)
    const multiply = service.compileForCell('A2*B2', 1, 2)

    expect(multiply.ast).not.toBe(add.ast)
    expect(multiply.source).toBe('A2*B2')
    expect(multiply.deps).toEqual(['A2', 'B2'])
  })

  it('clears cached families between initialization batches', () => {
    const service = createEngineFormulaTemplateNormalizationService()

    const first = service.resolveForCell('A1+B1', 0, 2)
    service.reset()
    const second = service.resolveForCell('A2+B2', 1, 2)

    expect(second.compiled.ast).not.toBe(first.compiled.ast)
    expect(second.compiled.deps).toEqual(['A2', 'B2'])
  })

  it('reuses an existing family when the recent-column cache was displaced', () => {
    const service = createEngineFormulaTemplateNormalizationService()

    const first = service.compileForCell('A1+B1', 0, 2)
    service.compileForCell('A1*B1', 0, 2)
    const reused = service.compileForCell('A1+B1', 0, 2)

    expect(reused.ast).toBe(first.ast)
    expect(reused.source).toBe('A1+B1')
    expect(reused.deps).toEqual(['A1', 'B1'])
  })

  it('translates an existing family after a recent-column cache miss', () => {
    const service = createEngineFormulaTemplateNormalizationService()

    const first = service.compileForCell('A1+B1', 0, 2)
    service.compileForCell('A1*B1', 0, 2)
    const translated = service.compileForCell('A2+B2', 1, 2)

    expect(translated.ast).toBe(first.ast)
    expect(translated.astMatchesSource).toBe(false)
    expect(translated.deps).toEqual(['A2', 'B2'])
  })

  it('tracks parse work only when a new template family is compiled', () => {
    const counters = createEngineCounters()
    const service = createEngineFormulaTemplateNormalizationService({ counters })

    service.compileForCell('A1+B1', 0, 2)
    service.compileForCell('A2+B2', 1, 2)
    service.compileForCell('A1*B1', 0, 2)

    expect(counters.formulasParsed).toBe(2)
  })

  it('reuses the same compiled owner for identical invariant sources across cells', () => {
    const service = createEngineFormulaTemplateNormalizationService()

    const first = service.resolveForCell('1+2', 0, 0)
    const second = service.resolveForCell('1+2', 0, 1)

    expect(second.templateId).toBe(first.templateId)
    expect(second.compiled).toBe(first.compiled)
    expect(second.compiled.source).toBe('1+2')
  })

  it('keeps template families as runtime owners across transient cache clears', () => {
    const service = createEngineFormulaTemplateNormalizationService()

    const first = service.resolveForCell('A1+B1', 0, 2)
    service.clear()
    const second = service.resolveForCell('A2+B2', 1, 2)

    expect(second.templateId).toBe(first.templateId)
    expect(service.listTemplates()).toHaveLength(1)
  })

  it('fast-compiles simple direct aggregate families and translates reusable shifted windows', () => {
    const service = createEngineFormulaTemplateNormalizationService()

    const first = service.resolveForCell('SUM(A1:A3)', 0, 1)
    const second = service.resolveForCell('SUM(A2:A4)', 1, 1)

    expect(first.compiled.directAggregateCandidate?.aggregateKind).toBe('sum')
    expect(first.compiled.mode).toBe(FormulaMode.WasmFastPath)
    expect(second.templateId).toBe(first.templateId)
    expect(second.compiled.deps).toEqual(['A2:A4'])
    expect(second.compiled.symbolicRanges).toEqual(['A2:A4'])
    expect(second.compiled.astMatchesSource).toBe(false)
  })

  it('collapses anchored prefix aggregate families into one template owner', () => {
    const counters = createEngineCounters()
    const service = createEngineFormulaTemplateNormalizationService({ counters })

    const first = service.resolveForCell('SUM(A1:A1)', 0, 5)
    const second = service.resolveForCell('SUM(A1:A2)', 1, 5)
    const third = service.resolveForCell('SUM(A1:A3)', 2, 5)

    expect(second.templateId).toBe(first.templateId)
    expect(third.templateId).toBe(first.templateId)
    expect(service.listTemplates()).toHaveLength(1)
    expect(counters.formulasParsed).toBe(1)
    expect(second.compiled.directAggregateCandidate?.aggregateKind).toBe('sum')
    expect(second.compiled.deps).toEqual(['A1:A2'])
    expect(third.compiled.deps).toEqual(['A1:A3'])
    expect(second.compiled.astMatchesSource).toBe(true)
    expect(third.compiled.astMatchesSource).toBe(true)
  })

  it('returns undefined for unknown template ids and exposes hydrated template snapshots', () => {
    const service = createEngineFormulaTemplateNormalizationService()
    const first = service.resolveForCell('A1+B1', 0, 0)

    expect(service.resolveByTemplateId(999_999, 'A1+B1', 0, 0)).toBeUndefined()
    expect(service.listTemplates()).toEqual([
      expect.objectContaining({
        id: first.templateId,
        templateKey: first.templateKey,
        baseSource: 'A1+B1',
      }),
    ])

    service.reset()
    service.hydrateTemplates([
      {
        id: first.templateId,
        templateKey: first.templateKey,
        baseSource: 'A1+B1',
        baseRow: 0,
        baseCol: 0,
        compiled: first.compiled,
      },
    ])

    const hydrated = service.resolveByTemplateId(first.templateId, 'A1+B1', 0, 0)
    expect(hydrated?.compiled).toBe(first.compiled)
  })
})
