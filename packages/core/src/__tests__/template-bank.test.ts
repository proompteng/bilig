import { describe, expect, it } from 'vitest'
import {
  tryMatchInitialSimpleRowRelativeBinaryTemplate,
  tryMatchInitialSimpleRowRelativeBinaryTemplateShape,
} from '../formula/initial-simple-direct-scalar-template.js'
import { createTemplateBank } from '../formula/template-bank.js'
import { createEngineCounters } from '../perf/engine-counters.js'

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

  it('translates trusted simple template ids without source-specific recompilation', () => {
    const bank = createTemplateBank()

    const scalar = bank.resolve('E2*2', 1, 5)
    const trustedScalar = bank.resolveTrustedById(scalar.templateId, 'E3*2', 2, 5)
    const aggregate = bank.resolve('SUM(A2:A4)', 5, 3)
    const trustedAggregate = bank.resolveTrustedById(aggregate.templateId, 'SUM(A3:A5)', 6, 3)
    const prefixAggregate = bank.resolve('SUM(A1:A1)', 0, 4)
    const trustedPrefixAggregate = bank.resolveTrustedById(prefixAggregate.templateId, 'SUM(A1:A3)', 2, 4)

    expect(trustedScalar?.templateId).toBe(scalar.templateId)
    expect(trustedScalar?.compiled.astMatchesSource).toBe(false)
    expect(trustedScalar?.compiled.symbolicRefs).toEqual(['E3'])
    expect(trustedScalar?.compiled.parsedSymbolicRefs).toEqual([{ address: 'E3', row: 2, col: 4, rowAbsolute: false, colAbsolute: false }])
    expect(trustedAggregate?.templateId).toBe(aggregate.templateId)
    expect(trustedAggregate?.compiled.astMatchesSource).toBe(false)
    expect(trustedAggregate?.compiled.symbolicRanges).toEqual(['A3:A5'])
    expect(trustedAggregate?.compiled.parsedSymbolicRanges).toEqual([
      {
        kind: 'range',
        refKind: 'cells',
        address: 'A3:A5',
        startAddress: 'A3',
        endAddress: 'A5',
        startRow: 2,
        endRow: 4,
        startCol: 0,
        endCol: 0,
      },
    ])
    expect(trustedPrefixAggregate?.templateId).toBe(prefixAggregate.templateId)
    expect(trustedPrefixAggregate?.compiled.astMatchesSource).toBe(false)
    expect(trustedPrefixAggregate?.compiled.symbolicRanges).toEqual(['A1:A3'])
    expect(trustedPrefixAggregate?.compiled.parsedSymbolicRanges).toEqual([
      {
        kind: 'range',
        refKind: 'cells',
        address: 'A1:A3',
        startAddress: 'A1',
        endAddress: 'A3',
        startRow: 0,
        endRow: 2,
        startCol: 0,
        endCol: 0,
      },
    ])
  })

  it('reuses row-relative rectangular aggregate template families without reparsing every row', () => {
    const counters = createEngineCounters()
    const bank = createTemplateBank({ counters })

    const first = bank.resolve('SUM(A1:F1)', 0, 6)
    const second = bank.resolve('SUM(A2:F2)', 1, 6)
    const third = bank.resolve('SUM(A3:F3)', 2, 6)

    expect(second.templateId).toBe(first.templateId)
    expect(third.templateId).toBe(first.templateId)
    expect(counters.formulasParsed).toBe(1)
    expect(second.compiled.symbolicRanges).toEqual(['A2:F2'])
    expect(second.compiled.parsedSymbolicRanges).toEqual([
      {
        kind: 'range',
        refKind: 'cells',
        address: 'A2:F2',
        startAddress: 'A2',
        endAddress: 'F2',
        startRow: 1,
        endRow: 1,
        startCol: 0,
        endCol: 5,
      },
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

  it('does not reuse simple template families for unsafe row references', () => {
    const bank = createTemplateBank()
    const unsafeRow = String(Number.MAX_SAFE_INTEGER + 1)

    const scalar = bank.resolve('A1+1', 0, 0)
    const aggregate = bank.resolve('SUM(A1:A1)', 0, 0)

    expect(bank.resolveById(scalar.templateId, `A${unsafeRow}+1`, Number.MAX_SAFE_INTEGER, 0)).toBeUndefined()
    expect(bank.resolveById(aggregate.templateId, `SUM(A1:A${unsafeRow})`, Number.MAX_SAFE_INTEGER, 0)).toBeUndefined()
  })
})
