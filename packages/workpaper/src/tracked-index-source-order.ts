export interface TrackedSortedIndexSource {
  readonly changedCellIndices: readonly number[] | Uint32Array
  readonly changedCellIndicesSortedDisjoint?: boolean
}

export function trackedSourceHasSortedDisjointIndices(source: TrackedSortedIndexSource): boolean {
  if (source.changedCellIndicesSortedDisjoint !== undefined) {
    return source.changedCellIndicesSortedDisjoint
  }
  let previous = -1
  for (let index = 0; index < source.changedCellIndices.length; index += 1) {
    const cellIndex = source.changedCellIndices[index]!
    if (!Number.isInteger(cellIndex) || cellIndex < 0 || cellIndex <= previous) {
      return false
    }
    previous = cellIndex
  }
  return true
}
