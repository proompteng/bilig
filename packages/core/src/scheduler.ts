import type { CellStore } from './cell-store.js'
import { CalcChain } from './scheduler/calc-chain.js'
import { DirtyFrontier } from './scheduler/dirty-frontier.js'

type U32 = Uint32Array

export interface SchedulerReverseGraph {
  getDependents(entityId: number): U32
}

export interface SchedulerResult {
  orderedFormulaCellIndices: U32
  orderedFormulaCount: number
  rangeNodeVisits: number
}

export class RecalcScheduler {
  private readonly dirtyFrontier = new DirtyFrontier()
  private readonly calcChain = new CalcChain()

  rebuildChain(formulaCellIndices: Iterable<number> | readonly number[] | U32, cellStore: CellStore): void {
    this.calcChain.rebuild(formulaCellIndices, cellStore)
  }

  collectDirty(
    changedRoots: readonly number[] | U32,
    graph: SchedulerReverseGraph,
    cellStore: CellStore,
    formulaCellIndices: Iterable<number>,
    formulaCount: number,
    hasFormula: (cellIndex: number) => boolean,
    rangeCount: number,
  ): SchedulerResult {
    if (!this.calcChain.hasChainFor(formulaCount)) {
      this.calcChain.rebuild(formulaCellIndices, cellStore)
    }
    const dirty = this.dirtyFrontier.collectDirty(changedRoots, graph, hasFormula, cellStore.size, rangeCount)
    const ordered = this.calcChain.orderDirty(dirty.dirtyFormulaCellIndices, dirty.dirtyFormulaCount)
    return {
      orderedFormulaCellIndices: ordered.orderedFormulaCellIndices,
      orderedFormulaCount: ordered.orderedFormulaCount,
      rangeNodeVisits: dirty.rangeNodeVisits,
    }
  }
}
