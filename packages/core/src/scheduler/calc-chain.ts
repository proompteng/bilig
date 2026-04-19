import type { CellStore } from '../cell-store.js'
import { growUint32 } from '../engine-buffer-utils.js'

type U32 = Uint32Array

export interface CalcChainResult {
  orderedFormulaCellIndices: U32
  orderedFormulaCount: number
}

export class CalcChain {
  private rankCounts: U32 = new Uint32Array(128)
  private orderedDirty: U32 = new Uint32Array(128)

  orderDirty(dirtyFormulaCellIndices: U32, dirtyFormulaCount: number, cellStore: CellStore): CalcChainResult {
    if (dirtyFormulaCount === 0) {
      return {
        orderedFormulaCellIndices: this.orderedDirty,
        orderedFormulaCount: 0,
      }
    }

    let minRank = Number.MAX_SAFE_INTEGER
    let maxRank = 0
    let isAlreadyOrdered = true
    let previousRank = -1
    for (let index = 0; index < dirtyFormulaCount; index += 1) {
      const cellIndex = dirtyFormulaCellIndices[index]!
      const rank = cellStore.topoRanks[cellIndex] ?? 0
      if (rank < previousRank) {
        isAlreadyOrdered = false
      } else {
        previousRank = rank
      }
      minRank = Math.min(minRank, rank)
      maxRank = Math.max(maxRank, rank)
    }

    if (isAlreadyOrdered) {
      return {
        orderedFormulaCellIndices: dirtyFormulaCellIndices,
        orderedFormulaCount: dirtyFormulaCount,
      }
    }

    const rankSpan = maxRank - minRank + 1
    this.ensureRankCapacity(rankSpan + 1, dirtyFormulaCount + 1)
    this.rankCounts.fill(0, 0, rankSpan)
    for (let index = 0; index < dirtyFormulaCount; index += 1) {
      const cellIndex = dirtyFormulaCellIndices[index]!
      const rank = (cellStore.topoRanks[cellIndex] ?? 0) - minRank
      this.rankCounts[rank] = (this.rankCounts[rank] ?? 0) + 1
    }

    let prefix = 0
    for (let index = 0; index < rankSpan; index += 1) {
      const count = this.rankCounts[index]!
      this.rankCounts[index] = prefix
      prefix += count
    }

    for (let index = 0; index < dirtyFormulaCount; index += 1) {
      const cellIndex = dirtyFormulaCellIndices[index]!
      const rank = (cellStore.topoRanks[cellIndex] ?? 0) - minRank
      const target = this.rankCounts[rank]!
      this.orderedDirty[target] = cellIndex
      this.rankCounts[rank] = target + 1
    }

    return {
      orderedFormulaCellIndices: this.orderedDirty,
      orderedFormulaCount: dirtyFormulaCount,
    }
  }

  private ensureRankCapacity(rankCountSize: number, dirtySize: number): void {
    if (rankCountSize > this.rankCounts.length) {
      let capacity = this.rankCounts.length
      while (capacity < rankCountSize) {
        capacity *= 2
      }
      this.rankCounts = growUint32(this.rankCounts, capacity)
    }
    if (dirtySize > this.orderedDirty.length) {
      let capacity = this.orderedDirty.length
      while (capacity < dirtySize) {
        capacity *= 2
      }
      this.orderedDirty = growUint32(this.orderedDirty, capacity)
    }
  }
}
