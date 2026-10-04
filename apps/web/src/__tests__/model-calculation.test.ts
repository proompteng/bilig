import { describe, expect, it } from 'vitest'
import { captureModelScenario, importModelBackup, parseModelDocument } from '../models/model-document.js'
import { createModel } from '../models/model-templates.js'
import { calculateModel } from '../models/model-calculation.js'
import { createWorkPaperFromDocument, parseWorkPaperDocument } from '@bilig/workpaper/browser'

describe('model workspace proof', () => {
  it('recalculates five what-if steps and leaves the model and portable baseline unchanged', () => {
    const model = createModel('contribution')
    const before = structuredClone(model)
    const result = calculateModel(model, { inputId: 'units', step: 10 })
    expect(result.sensitivity?.rows.map((row) => row.inputValue)).toEqual([80, 90, 100, 110, 120])
    expect(result.sensitivity?.rows.map((row) => row.values.find((value) => value.id === 'profit'))).toEqual(
      [1800, 2400, 3000, 3600, 4200].map((value) => expect.objectContaining({ kind: 'number', value })),
    )
    expect(model).toEqual(before)
    expect(result.current[0]).toMatchObject({ value: 10000 })
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(result.workpaperJson))
    try {
      const sheet = restored.getSheetId('Results')!
      expect(restored.getCellValue({ sheet, row: 1, col: 1 })).toMatchObject({ value: 10000 })
    } finally {
      restored.dispose()
    }
  })

  it('handles zero and negative assumptions and preserves formula errors in what-if results', () => {
    const model = createModel('custom')
    const result = calculateModel(
      { ...model, inputs: [{ ...model.inputs[0], value: 0 }], outputs: [{ ...model.outputs[0], formula: '=1/Inputs!B2' }] },
      { inputId: 'input-1', step: 1 },
    )
    expect(result.sensitivity?.rows.map((row) => row.inputValue)).toEqual([-2, -1, 0, 1, 2])
    expect(result.sensitivity?.rows[2]?.values[0]).toMatchObject({ kind: 'error', text: '#DIV/0!' })
    expect(result.sensitivity?.rows[0]?.values[0]).toMatchObject({ value: -0.5 })
  })

  it.each([0, -1, Infinity, NaN])('rejects an invalid what-if step %s', (step) => {
    expect(() => calculateModel(createModel('contribution'), { inputId: 'units', step })).toThrow('Step size')
  })

  it('rejects a missing assumption instead of displaying results for the wrong input', () => {
    expect(() => calculateModel(createModel('contribution'), { inputId: 'missing', step: 1 })).toThrow('assumption')
  })

  it('keeps the exact baseline value instead of rounding large assumptions during analysis', () => {
    const model = createModel('custom')
    const value = 123456789012345
    const result = calculateModel({ ...model, inputs: [{ ...model.inputs[0], value }] }, { inputId: 'input-1', step: 1 })
    expect(result.sensitivity?.rows[2]?.inputValue).toBe(value)
    expect(result.sensitivity?.rows[2]?.values).toEqual(result.current)
  })

  it('recalculates a changed assumption and verifies portable WorkPaper restore', () => {
    const model = createModel('contribution')
    const before = calculateModel(model, null)
    const edited = {
      ...model,
      inputs: model.inputs.map((input) =>
        input.id === 'units' ? { id: input.id, label: input.label, format: input.format, value: 150 } : input,
      ),
    }
    const after = calculateModel(edited, null)
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
    const result = calculateModel(changed, null)
    expect(result.current[0]).toMatchObject({ value: 1 })
    expect(result.scenarios[0]?.values[0]).toMatchObject({ value: 10000 })
  })

  it('preserves error results instead of claiming successful calculation', () => {
    const model = createModel('custom')
    const result = calculateModel({ ...model, outputs: [{ id: 'bad', label: 'Invalid', formula: '=1/0', format: 'number' }] }, null)
    expect(result.current[0]).toMatchObject({ kind: 'error', text: '#DIV/0!' })
  })

  it('imports backups as independent models with identical calculated results', () => {
    const source = captureModelScenario(createModel('project'), 'Approved')
    const imported = importModelBackup(JSON.stringify(source))
    expect(imported.id).not.toBe(source.id)
    expect(imported.scenarios).toEqual(source.scenarios)
    expect(calculateModel(imported, null).current).toEqual(calculateModel(source, null).current)
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
