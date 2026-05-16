import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'

import { EngineFormulaBindingError } from '../engine/errors.js'
import { formulaBindingEffect } from '../engine/services/formula-binding-effect.js'

describe('formulaBindingEffect', () => {
  it('returns successful binding operation values unchanged', () => {
    expect(Effect.runSync(formulaBindingEffect('fallback', () => 42))).toBe(42)
  })

  it('wraps thrown failures in EngineFormulaBindingError with the original cause', () => {
    const cause = new Error('parse failed')
    const result = Effect.runSync(
      Effect.either(
        formulaBindingEffect('fallback', () => {
          throw cause
        }),
      ),
    )

    expect(result._tag).toBe('Left')
    expect(result.left).toBeInstanceOf(EngineFormulaBindingError)
    expect(result.left.message).toBe('parse failed')
    expect(result.left.cause).toBe(cause)
  })

  it('uses the fallback message when the thrown value has no useful message', () => {
    const result = Effect.runSync(
      Effect.either(
        formulaBindingEffect('fallback', () => {
          throw new Error('')
        }),
      ),
    )

    expect(result._tag).toBe('Left')
    expect(result.left).toBeInstanceOf(EngineFormulaBindingError)
    expect(result.left.message).toBe('fallback')
  })
})
