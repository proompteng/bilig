import { describe, expect, it } from 'vitest'
import { WorkPaperParseError } from '../work-paper-errors.js'
import { restorePublicWorkPaperFormula, rewriteWorkPaperFormulaForStorage } from '../work-paper-formula-rewrite.js'
import { makeNamedExpressionKey } from '../work-paper-runtime-helpers.js'

describe('work paper formula rewrite helpers', () => {
  it('rewrites scoped named expressions and function aliases for storage', () => {
    const namedExpressions = new Map([
      [
        makeNamedExpressionKey('Rate'),
        {
          publicName: 'Rate',
          internalName: '__BILIG_RATE',
        },
      ],
      [
        makeNamedExpressionKey('Tax', 2),
        {
          publicName: 'Tax',
          internalName: '__BILIG_S2_TAX',
          scope: 2,
        },
      ],
    ])
    const functionAliasLookup = new Map([
      [
        'DOUBLE',
        {
          publicName: 'DOUBLE',
          internalName: '__BILIG_DOUBLE',
        },
      ],
    ])

    const rewritten = rewriteWorkPaperFormulaForStorage({
      formula: '=Tax + Rate + DOUBLE(1)',
      ownerSheetId: 2,
      namedExpressions,
      functionAliasLookup,
      messageOf: (error, fallback) => (error instanceof Error ? error.message : fallback),
    })

    expect(rewritten).toContain('__BILIG_S2_TAX')
    expect(rewritten).toContain('__BILIG_RATE')
    expect(rewritten).toContain('__BILIG_DOUBLE')
  })

  it('restores sheet-scoped names and internal function names for public formulas', () => {
    const namedExpressions = new Map([
      [
        makeNamedExpressionKey('Tax', 2),
        {
          publicName: 'Tax',
          internalName: '__BILIG_S2_TAX',
          scope: 2,
        },
      ],
    ])
    const internalFunctionLookup = new Map([
      [
        '__BILIG_DOUBLE',
        {
          publicName: 'DOUBLE',
          internalName: '__BILIG_DOUBLE',
        },
      ],
    ])

    const restored = restorePublicWorkPaperFormula({
      formula: '__BILIG_S2_TAX + __BILIG_DOUBLE(1)',
      ownerSheetId: 2,
      namedExpressions,
      internalFunctionLookup,
    })

    expect(restored).toContain('Tax')
    expect(restored).toContain('DOUBLE')
  })

  it('wraps storage parse failures in the public parse error type', () => {
    expect(() =>
      rewriteWorkPaperFormulaForStorage({
        formula: '=',
        ownerSheetId: 1,
        namedExpressions: new Map(),
        functionAliasLookup: new Map([['DOUBLE', { publicName: 'DOUBLE', internalName: '__BILIG_DOUBLE' }]]),
        messageOf: () => 'bad formula',
      }),
    ).toThrow(WorkPaperParseError)
  })
})
