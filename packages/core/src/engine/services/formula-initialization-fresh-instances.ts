import type { FormulaInstanceSnapshot } from '../../formula/formula-instance-table.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { InitialFormulaEntryRefSource } from './formula-initialization-refs.js'

export function writeDeferredFormulaInstance(
  records: FormulaInstanceSnapshot[] | undefined,
  index: number,
  prepared: {
    readonly cellIndex: number
    readonly row: number
    readonly col: number
    readonly ownerSheetName: string
  },
  formula: RuntimeFormula | undefined,
): number {
  if (!records || !formula) {
    return index
  }
  records[index] = {
    cellIndex: prepared.cellIndex,
    sheetName: prepared.ownerSheetName,
    row: prepared.row,
    col: prepared.col,
    source: formula.source,
    ...(formula.templateId !== undefined ? { templateId: formula.templateId } : {}),
  }
  return index + 1
}

export function materializeDeferredFormulaInstances(records: FormulaInstanceSnapshot[], count: number): readonly FormulaInstanceSnapshot[] {
  return count === records.length ? records : records.slice(0, count)
}

export function readAlignedFreshFormulaInstancesFromRefs<Entry>(
  refs: InitialFormulaEntryRefSource<Entry>,
): readonly FormulaInstanceSnapshot[] | undefined {
  if (!hasFreshFormulaInstances(refs)) {
    return undefined
  }
  const records = refs.freshFormulaInstances
  return records !== undefined && records.length === refs.length ? records : undefined
}

function hasFreshFormulaInstances(value: unknown): value is {
  readonly freshFormulaInstances: readonly FormulaInstanceSnapshot[] | undefined
} {
  return typeof value === 'object' && value !== null && 'freshFormulaInstances' in value
}
