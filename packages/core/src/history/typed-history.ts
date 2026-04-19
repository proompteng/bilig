import type { EngineCellMutationRef } from '../cell-mutations-at.js'

export const CELL_MUTATION_TRANSACTION_KIND = 'cell-mutations' as const

export interface CellMutationTransactionRecord {
  readonly kind: typeof CELL_MUTATION_TRANSACTION_KIND
  readonly refs: readonly EngineCellMutationRef[]
  readonly potentialNewCells?: number
}

export function isCellMutationTransactionRecord(value: unknown): value is CellMutationTransactionRecord {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const kind = Reflect.get(value, 'kind')
  const refs = Reflect.get(value, 'refs')
  return kind === CELL_MUTATION_TRANSACTION_KIND && Array.isArray(refs)
}
