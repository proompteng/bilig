import type { CellStore } from '../cell-store.js'
import { growUint32 } from '../engine-buffer-utils.js'

type U32 = Uint32Array

export interface CalcChainResult {
  orderedFormulaCellIndices: U32
  orderedFormulaCount: number
}

export class CalcChain {
  private rankCounts: U32 = new Uint32Array(128)
  private formulaIds: U32 = new Uint32Array(128)
  private orderedChain: U32 = new Uint32Array(128)
  private chainIndexByCell: U32 = new Uint32Array(128)
  private orderedDirty: U32 = new Uint32Array(128)
  private dirtySeen: U32 = new Uint32Array(128)
  private dirtyEpoch = 1
  private chainFormulaCount = 0

  rebuild(formulaCellIndices: Iterable<number> | readonly number[] | U32, cellStore: CellStore): void {
    this.ensureCellCapacity(cellStore.size + 1)

    for (let index = 0; index < this.chainFormulaCount; index += 1) {
      const cellIndex = this.orderedChain[index]!
      this.chainIndexByCell[cellIndex] = 0
    }

    let formulaCount = 0
    for (const cellIndex of formulaCellIndices) {
      this.ensureFormulaCapacity(formulaCount + 1)
      this.formulaIds[formulaCount] = cellIndex
      formulaCount += 1
    }

    this.chainFormulaCount = formulaCount
    if (formulaCount === 0) {
      return
    }

    let minRank = Number.MAX_SAFE_INTEGER
    let maxRank = 0
    for (let index = 0; index < formulaCount; index += 1) {
      const cellIndex = this.formulaIds[index]!
      const rank = cellStore.topoRanks[cellIndex] ?? 0
      minRank = Math.min(minRank, rank)
      maxRank = Math.max(maxRank, rank)
    }

    const rankSpan = maxRank - minRank + 1
    this.ensureRankCapacity(rankSpan, formulaCount)
    this.rankCounts.fill(0, 0, rankSpan)
    for (let index = 0; index < formulaCount; index += 1) {
      const cellIndex = this.formulaIds[index]!
      const rank = (cellStore.topoRanks[cellIndex] ?? 0) - minRank
      this.rankCounts[rank] = (this.rankCounts[rank] ?? 0) + 1
    }

    let prefix = 0
    for (let index = 0; index < rankSpan; index += 1) {
      const count = this.rankCounts[index]!
      this.rankCounts[index] = prefix
      prefix += count
    }

    for (let index = 0; index < formulaCount; index += 1) {
      const cellIndex = this.formulaIds[index]!
      const rank = (cellStore.topoRanks[cellIndex] ?? 0) - minRank
      const target = this.rankCounts[rank]!
      this.orderedChain[target] = cellIndex
      this.chainIndexByCell[cellIndex] = target + 1
      this.rankCounts[rank] = target + 1
    }
  }

  hasChainFor(formulaCount: number): boolean {
    return this.chainFormulaCount === formulaCount
  }

  coversDirty(dirtyFormulaCellIndices: U32, dirtyFormulaCount: number): boolean {
    for (let index = 0; index < dirtyFormulaCount; index += 1) {
      const cellIndex = dirtyFormulaCellIndices[index]!
      if ((this.chainIndexByCell[cellIndex] ?? 0) === 0) {
        return false
      }
    }
    return true
  }

  orderDirty(dirtyFormulaCellIndices: U32, dirtyFormulaCount: number): CalcChainResult {
    if (dirtyFormulaCount === 0) {
      return {
        orderedFormulaCellIndices: this.orderedDirty,
        orderedFormulaCount: 0,
      }
    }

    if (dirtyFormulaCount === this.chainFormulaCount) {
      return {
        orderedFormulaCellIndices: this.orderedChain,
        orderedFormulaCount: this.chainFormulaCount,
      }
    }

    this.dirtyEpoch += 1
    if (this.dirtyEpoch === 0xffff_ffff) {
      this.dirtyEpoch = 1
      this.dirtySeen.fill(0)
    }
    for (let index = 0; index < dirtyFormulaCount; index += 1) {
      const cellIndex = dirtyFormulaCellIndices[index]!
      if ((this.chainIndexByCell[cellIndex] ?? 0) !== 0) {
        this.dirtySeen[cellIndex] = this.dirtyEpoch
      }
    }

    let orderedCount = 0
    for (let index = 0; index < this.chainFormulaCount; index += 1) {
      const cellIndex = this.orderedChain[index]!
      if (this.dirtySeen[cellIndex] !== this.dirtyEpoch) {
        continue
      }
      this.orderedDirty[orderedCount] = cellIndex
      orderedCount += 1
    }

    return {
      orderedFormulaCellIndices: this.orderedDirty,
      orderedFormulaCount: orderedCount,
    }
  }

  private ensureRankCapacity(rankCountSize: number, chainSize: number): void {
    if (rankCountSize > this.rankCounts.length) {
      let capacity = this.rankCounts.length
      while (capacity < rankCountSize) {
        capacity *= 2
      }
      this.rankCounts = growUint32(this.rankCounts, capacity)
    }
    if (chainSize > this.orderedChain.length) {
      let capacity = this.orderedChain.length
      while (capacity < chainSize) {
        capacity *= 2
      }
      this.orderedChain = growUint32(this.orderedChain, capacity)
      this.orderedDirty = growUint32(this.orderedDirty, capacity)
    }
  }

  private ensureCellCapacity(size: number): void {
    if (size <= this.chainIndexByCell.length) {
      return
    }
    let capacity = this.chainIndexByCell.length
    while (capacity < size) {
      capacity *= 2
    }
    this.chainIndexByCell = growUint32(this.chainIndexByCell, capacity)
    this.dirtySeen = growUint32(this.dirtySeen, capacity)
  }

  private ensureFormulaCapacity(size: number): void {
    if (size <= this.formulaIds.length) {
      return
    }
    let capacity = this.formulaIds.length
    while (capacity < size) {
      capacity *= 2
    }
    this.formulaIds = growUint32(this.formulaIds, capacity)
    this.orderedChain = growUint32(this.orderedChain, capacity)
    this.orderedDirty = growUint32(this.orderedDirty, capacity)
  }
}
