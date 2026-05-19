import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { restorePublicWorkPaperFormula, rewriteWorkPaperFormulaForStorage } from '../work-paper-formula-rewrite.js'
import { makeNamedExpressionKey } from '../work-paper-runtime-helpers.js'

describe('work paper formula rewrite fuzz', () => {
  it('should roundtrip workbook-scoped named expressions through storage names', async () => {
    await runProperty({
      suite: 'headless/work-paper-formula-rewrite/workbook-name-roundtrip',
      arbitrary: fc.record({
        suffix: safeFormulaNameSuffixArbitrary,
        ownerSheetId: fc.integer({ min: 1, max: 32 }),
      }),
      predicate: async ({ suffix, ownerSheetId }) => {
        const publicName = `Rate${suffix}`
        const internalName = `__BILIG_${publicName.toUpperCase()}`
        const namedExpressions = new Map([
          [
            makeNamedExpressionKey(publicName),
            {
              publicName,
              internalName,
            },
          ],
        ])

        const stored = rewriteWorkPaperFormulaForStorage({
          formula: `=${publicName} + 1`,
          ownerSheetId,
          namedExpressions,
          functionAliasLookup: new Map(),
          messageOf: (error, fallback) => (error instanceof Error ? error.message : fallback),
        })
        const restored = restorePublicWorkPaperFormula({
          formula: stored,
          ownerSheetId,
          namedExpressions,
          internalFunctionLookup: new Map(),
        })

        expect(stored).toContain(internalName)
        expect(restored).toContain(publicName)
        expect(restored).not.toContain(internalName)
      },
      parameters: { numRuns: 120 },
    })
  })
})

const safeFormulaNameSuffixArbitrary = fc
  .array(fc.constantFrom('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'), {
    minLength: 1,
    maxLength: 8,
  })
  .map((chars) => chars.join(''))
