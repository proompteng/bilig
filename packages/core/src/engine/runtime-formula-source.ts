import { rewriteCompiledFormulaForStructuralTransform, rewriteFormulaForStructuralTransform, type CompiledFormula } from '@bilig/formula'
import type { RuntimeFormula } from './runtime-state.js'

export function getRuntimeFormulaSource(formula: RuntimeFormula): string {
  const deferred = formula.structuralSourceTransform
  if (!deferred) {
    return formula.source
  }
  return rewriteFormulaForStructuralTransform(formula.source, deferred.ownerSheetName, deferred.targetSheetName, deferred.transform)
}

export function getRuntimeFormulaStructuralCompiled(formula: RuntimeFormula): CompiledFormula | undefined {
  const deferred = formula.structuralSourceTransform
  if (!deferred) {
    return undefined
  }
  return rewriteCompiledFormulaForStructuralTransform(
    formula.compiled,
    deferred.ownerSheetName,
    deferred.targetSheetName,
    deferred.transform,
  ).compiled
}
