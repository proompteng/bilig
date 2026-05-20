export function hasQueuedFormulaDependency(
  cellIndex: number,
  queuedFormulaCells: ReadonlySet<number>,
  forEachFormulaDependencyCell: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void,
): boolean {
  if (queuedFormulaCells.size === 0) {
    return false
  }
  let hasQueuedDependency = false
  forEachFormulaDependencyCell(cellIndex, (dependencyCellIndex) => {
    hasQueuedDependency ||= queuedFormulaCells.has(dependencyCellIndex)
  })
  return hasQueuedDependency
}
