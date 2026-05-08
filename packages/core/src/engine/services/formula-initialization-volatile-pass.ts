import type { U32 } from '../runtime-state.js'

export function recalculateFreshVolatileFormulasAfterInitialMaterialization(
  args: {
    readonly beginMutationCollection: () => void
    readonly markVolatileFormulasChanged: (count: number) => number
    readonly getChangedFormulaBuffer: () => U32
    readonly recalculate: (changedRoots: readonly number[] | U32, kernelSyncRoots?: readonly number[] | U32) => U32
    readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32
  },
  recalculated: U32,
): U32 {
  args.beginMutationCollection()
  const volatileFormulaChangedCount = args.markVolatileFormulasChanged(0)
  if (volatileFormulaChangedCount === 0) {
    return recalculated
  }
  const volatileRoots = args.getChangedFormulaBuffer().subarray(0, volatileFormulaChangedCount)
  return args.reconcilePivotOutputs(args.recalculate(volatileRoots, volatileRoots), false)
}
