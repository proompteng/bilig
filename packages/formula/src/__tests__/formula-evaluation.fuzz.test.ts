import { describe, expect, it } from 'vitest'
import { evaluateAst } from '../js-evaluator.js'
import { parseFormula } from '../parser.js'
import { serializeFormula } from '../translation.js'
import { evaluableFormulaArbitrary, evaluationContext } from './formula-fuzz-helpers.js'
import { runProperty } from '@bilig/test-fuzz'

describe('formula evaluation fuzz', () => {
  it('keeps JS evaluation stable across canonicalization for coercion-heavy formulas', async () => {
    await runProperty({
      suite: 'formula/evaluation/canonicalization-stability',
      arbitrary: evaluableFormulaArbitrary,
      predicate: (formula) => {
        const canonical = serializeFormula(parseFormula(formula))
        expect(evaluateAst(parseFormula(formula), evaluationContext)).toEqual(evaluateAst(parseFormula(canonical), evaluationContext))
      },
    })
  })
})
