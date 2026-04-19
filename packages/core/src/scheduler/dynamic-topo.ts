import type { CellStore } from '../cell-store.js'
import { makeCellEntity } from '../entity-ids.js'
import { growUint32 } from '../engine-buffer-utils.js'

type U32 = Uint32Array

export interface DynamicTopoGraph {
  readonly forEachFormulaDependencyCell: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void
  readonly collectFormulaDependents: (entityId: number) => U32
}

export interface DynamicTopoRepairResult {
  readonly repaired: boolean
  readonly orderedFormulaCellIndices: U32
  readonly orderedFormulaCount: number
}

export class DynamicTopo {
  private affectedEpoch = 1
  private affectedSeen: U32 = new Uint32Array(128)
  private affectedFormulaIds: U32 = new Uint32Array(128)
  private topoQueue: U32 = new Uint32Array(128)
  private topoIndegree: U32 = new Uint32Array(128)
  private externalPredecessorRanks: U32 = new Uint32Array(128)
  private orderedFormulaIds: U32 = new Uint32Array(128)

  repair(
    changedFormulaCellIndices: readonly number[] | U32,
    graph: DynamicTopoGraph,
    cellStore: CellStore,
    hasFormula: (cellIndex: number) => boolean,
  ): DynamicTopoRepairResult {
    this.ensureCapacity(cellStore.size + 1)
    this.affectedEpoch += 1
    if (this.affectedEpoch === 0xffff_ffff) {
      this.affectedEpoch = 1
      this.affectedSeen.fill(0)
    }

    let affectedCount = 0
    const addAffectedFormula = (cellIndex: number): void => {
      if (!hasFormula(cellIndex) || this.affectedSeen[cellIndex] === this.affectedEpoch) {
        return
      }
      this.affectedSeen[cellIndex] = this.affectedEpoch
      this.ensureFormulaCapacity(affectedCount + 1)
      this.affectedFormulaIds[affectedCount] = cellIndex
      affectedCount += 1
    }

    for (let index = 0; index < changedFormulaCellIndices.length; index += 1) {
      addAffectedFormula(changedFormulaCellIndices[index]!)
    }
    for (let index = 0; index < affectedCount; index += 1) {
      const formulaCellIndex = this.affectedFormulaIds[index]!
      const dependents = graph.collectFormulaDependents(makeCellEntity(formulaCellIndex))
      for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
        addAffectedFormula(dependents[dependentIndex]!)
      }
    }

    if (affectedCount === 0) {
      return {
        repaired: true,
        orderedFormulaCellIndices: this.orderedFormulaIds,
        orderedFormulaCount: 0,
      }
    }

    let queueLength = 0
    for (let index = 0; index < affectedCount; index += 1) {
      const formulaCellIndex = this.affectedFormulaIds[index]!
      this.topoIndegree[formulaCellIndex] = 0
      this.externalPredecessorRanks[formulaCellIndex] = 0
      graph.forEachFormulaDependencyCell(formulaCellIndex, (dependencyCellIndex) => {
        if (!hasFormula(dependencyCellIndex)) {
          return
        }
        if (this.affectedSeen[dependencyCellIndex] === this.affectedEpoch) {
          this.topoIndegree[formulaCellIndex] = (this.topoIndegree[formulaCellIndex] ?? 0) + 1
          return
        }
        const externalFloor = (cellStore.topoRanks[dependencyCellIndex] ?? 0) + 1
        if (externalFloor > (this.externalPredecessorRanks[formulaCellIndex] ?? 0)) {
          this.externalPredecessorRanks[formulaCellIndex] = externalFloor
        }
      })
      if ((this.topoIndegree[formulaCellIndex] ?? 0) === 0) {
        this.ensureQueueCapacity(queueLength + 1)
        this.topoQueue[queueLength] = formulaCellIndex
        queueLength += 1
      }
    }

    let orderedCount = 0
    for (let queueIndex = 0; queueIndex < queueLength; queueIndex += 1) {
      const formulaCellIndex = this.topoQueue[queueIndex]!
      this.ensureOrderedCapacity(orderedCount + 1)
      this.orderedFormulaIds[orderedCount] = formulaCellIndex
      orderedCount += 1

      const dependents = graph.collectFormulaDependents(makeCellEntity(formulaCellIndex))
      for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
        const dependentCellIndex = dependents[dependentIndex]!
        if (this.affectedSeen[dependentCellIndex] !== this.affectedEpoch) {
          continue
        }
        const next = (this.topoIndegree[dependentCellIndex] ?? 0) - 1
        this.topoIndegree[dependentCellIndex] = next
        if (next === 0) {
          this.ensureQueueCapacity(queueLength + 1)
          this.topoQueue[queueLength] = dependentCellIndex
          queueLength += 1
        }
      }
    }

    if (orderedCount !== affectedCount) {
      return {
        repaired: false,
        orderedFormulaCellIndices: this.orderedFormulaIds,
        orderedFormulaCount: orderedCount,
      }
    }

    let nextRank = 0
    for (let index = 0; index < orderedCount; index += 1) {
      const formulaCellIndex = this.orderedFormulaIds[index]!
      nextRank = Math.max(nextRank, this.externalPredecessorRanks[formulaCellIndex] ?? 0)
      cellStore.topoRanks[formulaCellIndex] = nextRank
      nextRank += 1
    }

    return {
      repaired: true,
      orderedFormulaCellIndices: this.orderedFormulaIds,
      orderedFormulaCount: orderedCount,
    }
  }

  private ensureCapacity(size: number): void {
    if (size <= this.affectedSeen.length) {
      return
    }
    let capacity = this.affectedSeen.length
    while (capacity < size) {
      capacity *= 2
    }
    this.affectedSeen = growUint32(this.affectedSeen, capacity)
    this.topoIndegree = growUint32(this.topoIndegree, capacity)
    this.externalPredecessorRanks = growUint32(this.externalPredecessorRanks, capacity)
  }

  private ensureFormulaCapacity(size: number): void {
    if (size <= this.affectedFormulaIds.length) {
      return
    }
    let capacity = this.affectedFormulaIds.length
    while (capacity < size) {
      capacity *= 2
    }
    this.affectedFormulaIds = growUint32(this.affectedFormulaIds, capacity)
  }

  private ensureQueueCapacity(size: number): void {
    if (size <= this.topoQueue.length) {
      return
    }
    let capacity = this.topoQueue.length
    while (capacity < size) {
      capacity *= 2
    }
    this.topoQueue = growUint32(this.topoQueue, capacity)
  }

  private ensureOrderedCapacity(size: number): void {
    if (size <= this.orderedFormulaIds.length) {
      return
    }
    let capacity = this.orderedFormulaIds.length
    while (capacity < size) {
      capacity *= 2
    }
    this.orderedFormulaIds = growUint32(this.orderedFormulaIds, capacity)
  }
}
