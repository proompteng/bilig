import { describe, expect, it } from 'vitest'
import { captureModelScenario, importModelBackup, parseModelDocument } from '../models/model-document.js'
import { createModel } from '../models/model-templates.js'
import { calculateModel } from '../models/model-calculation.js'

describe('model workspace proof', () => {
  it('recalculates a changed assumption and verifies portable WorkPaper restore', () => {
    const model = createModel('contribution')
    const before = calculateModel(model)
    const edited = {
      ...model,
      inputs: model.inputs.map((input) =>
        input.id === 'units' ? { id: input.id, label: input.label, format: input.format, value: 150 } : input,
      ),
    }
    const after = calculateModel(edited)
    expect(before.current[0]).toMatchObject({ id: 'revenue', value: 10000 })
    expect(after.current[0]).toMatchObject({ id: 'revenue', value: 15000 })
    expect(after.restoreVerified).toBe(true)
    expect(after.workpaperJson).toContain('=Inputs!B2*Inputs!B3')
  })

  it('freezes scenario formulas as well as inputs', () => {
    const baseline = captureModelScenario(createModel('contribution'), 'Baseline')
    const changed = {
      ...baseline,
      outputs: baseline.outputs.map((output) => ({ id: output.id, label: output.label, format: output.format, formula: '=1' })),
    }
    const result = calculateModel(changed)
    expect(result.current[0]).toMatchObject({ value: 1 })
    expect(result.scenarios[0]?.values[0]).toMatchObject({ value: 10000 })
  })

  it('preserves error results instead of claiming successful calculation', () => {
    const model = createModel('custom')
    const result = calculateModel({ ...model, outputs: [{ id: 'bad', label: 'Invalid', formula: '=1/0', format: 'number' }] })
    expect(result.current[0]).toMatchObject({ kind: 'error', text: '#DIV/0!' })
  })

  it('imports backups as independent models with identical calculated results', () => {
    const source = captureModelScenario(createModel('project'), 'Approved')
    const imported = importModelBackup(JSON.stringify(source))
    expect(imported.id).not.toBe(source.id)
    expect(imported.scenarios).toEqual(source.scenarios)
    expect(calculateModel(imported).current).toEqual(calculateModel(source).current)
  })

  it.each([
    { format: 'future' },
    { inputs: [{ id: 'bad', label: 'Bad', value: Infinity, format: 'number' }] },
    { outputs: [{ id: 'bad', label: 'Bad', formula: 'not a formula', format: 'number' }] },
    { revision: -1 },
    { scenarios: Array.from({ length: 13 }, () => ({})) },
  ])('rejects malformed data at the persistence boundary: %j', (invalid) => {
    expect(() => parseModelDocument({ ...createModel('custom'), ...invalid })).toThrow()
  })
})
