import type { RuntimeFormula } from '../runtime-state.js'

export function updateVolatileFormulaIndexEntry(
  volatileFormulaCells: Set<number> | undefined,
  cellIndex: number,
  formula: RuntimeFormula | undefined,
): void {
  if (!volatileFormulaCells) {
    return
  }
  if (formula?.compiled.volatile) {
    volatileFormulaCells.add(cellIndex)
    return
  }
  volatileFormulaCells.delete(cellIndex)
}
