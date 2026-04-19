import { describe, expect, it } from 'vitest'
import { createTemplateBank } from '../formula/template-bank.js'

describe('TemplateBank', () => {
  it('returns stored snapshots by id and reports missing ids as undefined', () => {
    const bank = createTemplateBank()
    const resolved = bank.resolve('1+2', 0, 0)

    expect(bank.get(resolved.templateId)).toEqual(
      expect.objectContaining({
        id: resolved.templateId,
        templateKey: resolved.templateKey,
        baseSource: '1+2',
      }),
    )
    expect(bank.resolveById(999_999, '1+2', 0, 0)).toBeUndefined()
  })

  it('reuses an existing template record when the family key repeats', () => {
    const bank = createTemplateBank()

    const first = bank.resolve('A1+B1', 0, 0)
    const second = bank.resolve('A1+B1', 0, 0)

    expect(second.templateId).toBe(first.templateId)
    expect(bank.get(first.templateId)?.compiled).toBe(first.compiled)
  })
})
