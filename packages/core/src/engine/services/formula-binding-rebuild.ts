import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

export function rebuildAllFormulaBindingsNow(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly clearFormulaBookkeepingNow: () => void
  readonly bindFormulaNow: (cellIndex: number, ownerSheetName: string, source: string) => boolean
  readonly invalidateFormulaNow: (cellIndex: number) => void
  readonly isCellIndexMappedNow: (cellIndex: number) => boolean
  readonly pruneOrphanedDependencyCells: (cellIndices: readonly number[]) => void
}): number[] {
  const pending = [...args.serviceArgs.state.formulas.entries()].map(([cellIndex, formula]) => ({
    cellIndex,
    source: formula.source,
    dependencyIndices: [...formula.dependencyIndices],
    planId: formula.planId,
  }))
  pending.forEach(({ planId }) => {
    args.serviceArgs.compiledPlans.release(planId)
  })
  args.serviceArgs.state.formulas.clear()
  args.serviceArgs.formulaInstances.clear()
  args.serviceArgs.state.ranges.reset()
  args.serviceArgs.edgeArena.reset()
  args.serviceArgs.programArena.reset()
  args.serviceArgs.constantArena.reset()
  args.serviceArgs.rangeListArena.reset()
  args.serviceArgs.reverseState.reverseCellEdges.length = 0
  args.serviceArgs.reverseState.reverseRangeEdges.length = 0
  args.serviceArgs.reverseState.reverseDefinedNameEdges.clear()
  args.serviceArgs.reverseState.reverseTableEdges.clear()
  args.serviceArgs.reverseState.reverseSpillEdges.clear()
  args.serviceArgs.reverseState.reverseAggregateColumnEdges.clear()
  args.serviceArgs.reverseState.reverseExactLookupColumnEdges.clear()
  args.serviceArgs.reverseState.reverseSortedLookupColumnEdges.clear()
  args.clearFormulaBookkeepingNow()
  args.serviceArgs.regionGraph.reset()

  const activeCellIndices: number[] = []
  pending.forEach(({ cellIndex, source }) => {
    if (!args.isCellIndexMappedNow(cellIndex)) {
      args.serviceArgs.state.workbook.pruneCellIfEmpty(cellIndex)
      return
    }
    const ownerSheetName = args.serviceArgs.state.workbook.getSheetNameById(args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]!)
    if (!ownerSheetName || !args.serviceArgs.state.workbook.getSheet(ownerSheetName)) {
      return
    }
    try {
      args.bindFormulaNow(cellIndex, ownerSheetName, source)
    } catch {
      args.invalidateFormulaNow(cellIndex)
    }
    activeCellIndices.push(cellIndex)
  })
  args.pruneOrphanedDependencyCells(pending.flatMap(({ dependencyIndices }) => dependencyIndices))
  return activeCellIndices
}
