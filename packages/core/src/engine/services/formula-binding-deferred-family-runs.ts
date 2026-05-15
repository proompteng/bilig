import type { FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'
import { materializeDeferredFormulaFamilyRunMembers, type DeferredInitialFormulaFamilyRun } from './formula-initialization-family-runs.js'

export function registerDeferredFormulaFamilyIndexRunsNow(args: {
  readonly formulaFamilies: CreateEngineFormulaBindingServiceArgs['formulaFamilies']
  readonly formulaFamilyShapeKeyCache: FormulaBindingFamilyShapeKeyCache
  readonly runs: readonly DeferredInitialFormulaFamilyRun[]
}): void {
  args.formulaFamilies.clear()
  args.formulaFamilyShapeKeyCache.clear()
  args.runs.forEach((run) => {
    const step = run.cellIndices.length <= 1 ? 1 : run.step
    if (
      run.ordered &&
      step > 0 &&
      args.formulaFamilies.registerFreshUniformRun({
        sheetId: run.sheetId,
        templateId: run.templateId,
        shapeKey: run.shapeKey,
        axis: run.axis,
        fixedIndex: run.fixedIndex,
        start: run.start,
        step,
        cellIndices: run.cellIndices,
      })
    ) {
      return
    }
    args.formulaFamilies.registerFormulaRun({
      sheetId: run.sheetId,
      templateId: run.templateId,
      shapeKey: run.shapeKey,
      members: materializeDeferredFormulaFamilyRunMembers(run),
    })
  })
}
