import { rewriteCompiledFormulaForStructuralTransform, rewriteFormulaForStructuralTransform, type CompiledFormula } from '@bilig/formula'
import type { RuntimeFormula, RuntimeStructuralFormulaSourceTransform } from './runtime-state.js'

export function getRuntimeFormulaSource(
  formula: RuntimeFormula,
  inheritedStructuralSourceTransform?: RuntimeStructuralFormulaSourceTransform,
): string {
  const deferred = formula.structuralSourceTransform ?? inheritedStructuralSourceTransform
  if (!deferred) {
    return formula.source
  }
  return rewriteFormulaForStructuralTransform(formula.source, deferred.ownerSheetName, deferred.targetSheetName, deferred.transform)
}

export function getRuntimeFormulaStructuralCompiled(
  formula: RuntimeFormula,
  inheritedStructuralSourceTransform?: RuntimeStructuralFormulaSourceTransform,
): CompiledFormula | undefined {
  const deferred = formula.structuralSourceTransform ?? inheritedStructuralSourceTransform
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
