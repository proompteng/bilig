import { entityPayload, isExactLookupColumnEntity, isRangeEntity, isSortedLookupColumnEntity, makeCellEntity } from '../entity-ids.js'
import { growUint32 } from '../engine-buffer-utils.js'

type U32 = Uint32Array

export interface DirtyFrontierReverseGraph {
  getDependents(entityId: number): U32
}

export interface DirtyFrontierResult {
  dirtyFormulaCellIndices: U32
  dirtyFormulaCount: number
  rangeNodeVisits: number
}

export class DirtyFrontier {
  private cellEpoch = 1
  private rangeEpoch = 1
  private exactLookupEpoch = 1
  private sortedLookupEpoch = 1
  private visitedCells: U32 = new Uint32Array(64)
  private visitedRanges: U32 = new Uint32Array(64)
  private visitedExactLookupColumns = new Map<number, number>()
  private visitedSortedLookupColumns = new Map<number, number>()
  private entityQueue: U32 = new Uint32Array(128)
  private dirtyFormulaIds: U32 = new Uint32Array(128)

  collectDirty(
    changedRoots: readonly number[] | U32,
    graph: DirtyFrontierReverseGraph,
    hasFormula: (cellIndex: number) => boolean,
    cellCount: number,
    rangeCount: number,
  ): DirtyFrontierResult {
    this.ensureCellCapacity(cellCount + 1)
    this.ensureRangeCapacity(rangeCount + 1)
    this.cellEpoch += 1
    this.rangeEpoch += 1
    this.exactLookupEpoch += 1
    this.sortedLookupEpoch += 1
    if (this.exactLookupEpoch === 0xffff_ffff) {
      this.exactLookupEpoch = 1
      this.visitedExactLookupColumns.clear()
    }
    if (this.sortedLookupEpoch === 0xffff_ffff) {
      this.sortedLookupEpoch = 1
      this.visitedSortedLookupColumns.clear()
    }

    let queueLength = 0
    let dirtyLength = 0
    let rangeNodeVisits = 0

    for (let index = 0; index < changedRoots.length; index += 1) {
      const cellIndex = changedRoots[index]!
      const entityId = makeCellEntity(cellIndex)
      if (this.visitedCells[cellIndex] === this.cellEpoch) {
        continue
      }
      this.ensureEntityQueueCapacity(queueLength + 1)
      this.visitedCells[cellIndex] = this.cellEpoch
      this.entityQueue[queueLength] = entityId
      queueLength += 1
      if (!hasFormula(cellIndex)) {
        continue
      }
      this.ensureDirtyCapacity(dirtyLength + 1)
      this.dirtyFormulaIds[dirtyLength] = cellIndex
      dirtyLength += 1
    }

    for (let queueIndex = 0; queueIndex < queueLength; queueIndex += 1) {
      const entityId = this.entityQueue[queueIndex]!
      const dependents = graph.getDependents(entityId)
      if (isRangeEntity(entityId)) {
        rangeNodeVisits += 1
      }
      for (let depIndex = 0; depIndex < dependents.length; depIndex += 1) {
        const dependent = dependents[depIndex]!
        if (isRangeEntity(dependent)) {
          const rangeIndex = entityPayload(dependent)
          if (this.visitedRanges[rangeIndex] === this.rangeEpoch) {
            continue
          }
          this.visitedRanges[rangeIndex] = this.rangeEpoch
          this.ensureEntityQueueCapacity(queueLength + 1)
          this.entityQueue[queueLength] = dependent
          queueLength += 1
          continue
        }
        if (isExactLookupColumnEntity(dependent)) {
          const lookupColumnPayload = entityPayload(dependent)
          if (this.visitedExactLookupColumns.get(lookupColumnPayload) === this.exactLookupEpoch) {
            continue
          }
          this.visitedExactLookupColumns.set(lookupColumnPayload, this.exactLookupEpoch)
          this.ensureEntityQueueCapacity(queueLength + 1)
          this.entityQueue[queueLength] = dependent
          queueLength += 1
          continue
        }
        if (isSortedLookupColumnEntity(dependent)) {
          const lookupColumnPayload = entityPayload(dependent)
          if (this.visitedSortedLookupColumns.get(lookupColumnPayload) === this.sortedLookupEpoch) {
            continue
          }
          this.visitedSortedLookupColumns.set(lookupColumnPayload, this.sortedLookupEpoch)
          this.ensureEntityQueueCapacity(queueLength + 1)
          this.entityQueue[queueLength] = dependent
          queueLength += 1
          continue
        }

        const cellIndex = entityPayload(dependent)
        if (this.visitedCells[cellIndex] === this.cellEpoch) {
          continue
        }
        this.visitedCells[cellIndex] = this.cellEpoch
        this.ensureEntityQueueCapacity(queueLength + 1)
        this.entityQueue[queueLength] = dependent
        queueLength += 1
        if (!hasFormula(cellIndex)) {
          continue
        }
        this.ensureDirtyCapacity(dirtyLength + 1)
        this.dirtyFormulaIds[dirtyLength] = cellIndex
        dirtyLength += 1
      }
    }

    return {
      dirtyFormulaCellIndices: this.dirtyFormulaIds,
      dirtyFormulaCount: dirtyLength,
      rangeNodeVisits,
    }
  }

  private ensureCellCapacity(size: number): void {
    if (size <= this.visitedCells.length) {
      return
    }
    let capacity = this.visitedCells.length
    while (capacity < size) {
      capacity *= 2
    }
    this.visitedCells = growUint32(this.visitedCells, capacity)
    this.entityQueue = growUint32(this.entityQueue, capacity)
    this.dirtyFormulaIds = growUint32(this.dirtyFormulaIds, capacity)
  }

  private ensureRangeCapacity(size: number): void {
    if (size <= this.visitedRanges.length) {
      return
    }
    let capacity = this.visitedRanges.length
    while (capacity < size) {
      capacity *= 2
    }
    this.visitedRanges = growUint32(this.visitedRanges, capacity)
  }

  private ensureEntityQueueCapacity(size: number): void {
    if (size <= this.entityQueue.length) {
      return
    }
    let capacity = this.entityQueue.length
    while (capacity < size) {
      capacity *= 2
    }
    this.entityQueue = growUint32(this.entityQueue, capacity)
  }

  private ensureDirtyCapacity(size: number): void {
    if (size <= this.dirtyFormulaIds.length) {
      return
    }
    let capacity = this.dirtyFormulaIds.length
    while (capacity < size) {
      capacity *= 2
    }
    this.dirtyFormulaIds = growUint32(this.dirtyFormulaIds, capacity)
  }
}
