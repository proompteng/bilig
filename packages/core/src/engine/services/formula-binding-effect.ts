import { Effect } from 'effect'
import { EngineFormulaBindingError } from '../errors.js'
import { formulaBindingErrorMessage } from './formula-binding-plan-helpers.js'

export function formulaBindingEffect<T>(message: string, operation: () => T): Effect.Effect<T, EngineFormulaBindingError> {
  return Effect.try({
    try: operation,
    catch: (cause) =>
      new EngineFormulaBindingError({
        message: formulaBindingErrorMessage(message, cause),
        cause,
      }),
  })
}
