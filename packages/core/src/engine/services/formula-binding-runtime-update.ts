import type { CompiledPlanRecord, RuntimeFormula } from '../runtime-state.js'

export interface FormulaRuntimePlanFieldUpdate {
  readonly source: string
  readonly plan: CompiledPlanRecord
  readonly templateId: number | undefined
  readonly programLength: number
  readonly runtimeProgram?: Uint32Array
}

export function applyFormulaRuntimePlanFields(formula: RuntimeFormula, update: FormulaRuntimePlanFieldUpdate): void {
  formula.source = update.source
  formula.structuralSourceTransform = undefined
  formula.sourceRenameTransforms = undefined
  delete formula.preserveCachedValueOnFullRecalc
  formula.planId = update.plan.id
  formula.templateId = update.templateId
  formula.compiled = update.plan.compiled
  formula.plan = update.plan
  if (update.runtimeProgram !== undefined) {
    formula.runtimeProgram = update.runtimeProgram
  }
  formula.constants = update.plan.compiled.constants
  formula.programLength = update.programLength
  formula.constNumberLength = update.plan.compiled.constants.length
}
