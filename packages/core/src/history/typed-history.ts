import type { EngineCellMutationRef } from '../cell-mutations-at.js'

export interface CellMutationTransactionRecord {
  readonly kind: 'cell-mutations'
  readonly refs: readonly EngineCellMutationRef[]
  readonly potentialNewCells?: number
}
