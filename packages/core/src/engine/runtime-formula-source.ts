import {
  renameFormulaSheetReferences,
  rewriteCompiledFormulaForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  type CompiledFormula,
} from '@bilig/formula'
import type { RuntimeFormula, RuntimeStructuralFormulaSourceTransform } from './runtime-state.js'

function getRuntimeFormulaBaseSource(formula: RuntimeFormula): string {
  let source = formula.source
  formula.sourceRenameTransforms?.forEach((transform) => {
    source = renameFormulaSheetReferences(source, transform.oldSheetName, transform.newSheetName)
  })
  return source
}

export function getRuntimeFormulaSource(
  formula: RuntimeFormula,
  inheritedStructuralSourceTransform?: RuntimeStructuralFormulaSourceTransform,
): string {
  const source = getRuntimeFormulaBaseSource(formula)
  const deferred = formula.structuralSourceTransform ?? inheritedStructuralSourceTransform
  if (!deferred) {
    return source
  }
  return rewriteFormulaForStructuralTransform(source, deferred.ownerSheetName, deferred.targetSheetName, deferred.transform)
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
