import { describe, expect, it } from 'vitest'
import {
  tryMatchInitialSimpleRowRelativeBinaryTemplate,
  tryMatchInitialSimpleRowRelativeBinaryTemplateShape,
} from '../formula/initial-simple-direct-scalar-template.js'
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

  it('reuses simple row-relative binary template families without changing semantics', () => {
    const bank = createTemplateBank()

    const first = bank.resolve('A1+B1', 0, 2)
    const second = bank.resolve('A2+B2', 1, 2)
    const multiplied = bank.resolve('E2*2', 1, 5)
    const translatedMultiply = bank.resolve('E3*2', 2, 5)

    expect(second.templateId).toBe(first.templateId)
    expect(translatedMultiply.templateId).toBe(multiplied.templateId)
    expect(second.compiled.symbolicRefs).toEqual(['A2', 'B2'])
    expect(second.compiled.parsedDeps).toEqual([
      { kind: 'cell', address: 'A2', row: 1, col: 0, rowAbsolute: false, colAbsolute: false },
      { kind: 'cell', address: 'B2', row: 1, col: 1, rowAbsolute: false, colAbsolute: false },
    ])
    expect(translatedMultiply.compiled.symbolicRefs).toEqual(['E3'])
    expect(translatedMultiply.compiled.parsedDeps).toEqual([
      { kind: 'cell', address: 'E3', row: 2, col: 4, rowAbsolute: false, colAbsolute: false },
    ])
  })

  it('matches cached simple row-relative binary template shapes without reparsing the family', () => {
    const anchor = tryMatchInitialSimpleRowRelativeBinaryTemplate('A1+B1', 0, 2)

    expect(anchor).toBeDefined()
    const translated = tryMatchInitialSimpleRowRelativeBinaryTemplateShape('a2+b2', 1, 2, anchor!)

    expect(translated).toEqual(
      expect.objectContaining({
        templateKey: anchor!.templateKey,
        symbolicRefs: ['A2', 'B2'],
        parsedDeps: [
          { kind: 'cell', address: 'A2', row: 1, col: 0, rowAbsolute: false, colAbsolute: false },
          { kind: 'cell', address: 'B2', row: 1, col: 1, rowAbsolute: false, colAbsolute: false },
        ],
        parsedSymbolicRefs: [
          { address: 'A2', row: 1, col: 0, rowAbsolute: false, colAbsolute: false },
          { address: 'B2', row: 1, col: 1, rowAbsolute: false, colAbsolute: false },
        ],
      }),
    )
    expect(tryMatchInitialSimpleRowRelativeBinaryTemplateShape('A20+B20', 1, 2, anchor!)).toBeUndefined()
  })

  it('matches cached simple row-relative number template shapes exactly', () => {
    const anchor = tryMatchInitialSimpleRowRelativeBinaryTemplate('E2*2', 1, 5)

    expect(anchor).toBeDefined()
    const translated = tryMatchInitialSimpleRowRelativeBinaryTemplateShape('E3*2', 2, 5, anchor!)

    expect(translated).toEqual(
      expect.objectContaining({
        templateKey: anchor!.templateKey,
        symbolicRefs: ['E3'],
        parsedDeps: [{ kind: 'cell', address: 'E3', row: 2, col: 4, rowAbsolute: false, colAbsolute: false }],
        parsedSymbolicRefs: [{ address: 'E3', row: 2, col: 4, rowAbsolute: false, colAbsolute: false }],
      }),
    )
    expect(tryMatchInitialSimpleRowRelativeBinaryTemplateShape('E3*20', 2, 5, anchor!)).toBeUndefined()
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
