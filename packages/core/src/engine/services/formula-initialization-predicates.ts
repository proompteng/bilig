import type { RuntimeFormula } from '../runtime-state.js'

export function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export function canEvaluateInitialDirectRuntimeFormula(formula: RuntimeFormula | undefined): boolean {
  return (
    formula !== undefined &&
    !formula.compiled.volatile &&
    !formula.compiled.producesSpill &&
    (formula.directAggregate !== undefined ||
      formula.directCriteria !== undefined ||
      formula.directLookup !== undefined ||
      formula.directScalar !== undefined)
  )
}

export function hasPendingFormulaDependency(
  formula: RuntimeFormula,
  pendingFormulaCells: Uint8Array,
  getRangeMembers?: (rangeIndex: number) => Uint32Array,
): boolean {
  const dependencies = formula.dependencyIndices
  for (let index = 0; index < dependencies.length; index += 1) {
    if ((pendingFormulaCells[dependencies[index]!] ?? 0) !== 0) {
      return true
    }
  }
  if (getRangeMembers === undefined) {
    return false
  }
  const ranges = formula.graphRangeDependencies
  for (let rangeIndexCursor = 0; rangeIndexCursor < ranges.length; rangeIndexCursor += 1) {
    const members = getRangeMembers(ranges[rangeIndexCursor]!)
    for (let memberIndex = 0; memberIndex < members.length; memberIndex += 1) {
      if ((pendingFormulaCells[members[memberIndex]!] ?? 0) !== 0) {
        return true
      }
    }
  }
  return false
}
