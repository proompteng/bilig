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

  it('rejects stale template ids when the current source belongs to a different template family', () => {
    const bank = createTemplateBank()

    const prefix = bank.resolve('SUM(A1:A2)', 1, 3)

    expect(bank.resolveById(prefix.templateId, 'SUM(A1:A1)', 1, 3)).toBeUndefined()
    expect(bank.resolveById(prefix.templateId, 'SUM(A1:A3)', 2, 3)).toEqual(
      expect.objectContaining({
        templateId: prefix.templateId,
      }),
    )
  })
})
