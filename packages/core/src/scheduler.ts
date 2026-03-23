import type { CellStore } from "./cell-store.js";
import { entityPayload, isRangeEntity, makeCellEntity } from "./entity-ids.js";

type U32 = Uint32Array;

export interface SchedulerReverseGraph {
  getDependents(entityId: number): U32;
}

export interface SchedulerResult {
  orderedFormulaCellIndices: U32;
  orderedFormulaCount: number;
  rangeNodeVisits: number;
}

export class RecalcScheduler {
  private cellEpoch = 1;
  private rangeEpoch = 1;
  private visitedCells: U32 = new Uint32Array(64);
  private visitedRanges: U32 = new Uint32Array(64);
  private entityQueue: U32 = new Uint32Array(128);
  private dirtyFormulaIds: U32 = new Uint32Array(128);
  private rankCounts: U32 = new Uint32Array(128);
  private orderedDirty: U32 = new Uint32Array(128);

  collectDirty(
    changedRoots: readonly number[] | U32,
    graph: SchedulerReverseGraph,
    cellStore: CellStore,
    hasFormula: (cellIndex: number) => boolean,
    rangeCount: number,
  ): SchedulerResult {
    this.ensureCellCapacity(cellStore.size + 1);
    this.ensureRangeCapacity(rangeCount + 1);
    this.cellEpoch += 1;
    this.rangeEpoch += 1;

    let queueLength = 0;
    let dirtyLength = 0;
    let rangeNodeVisits = 0;
    let minRank = Number.MAX_SAFE_INTEGER;
    let maxRank = 0;

    for (let index = 0; index < changedRoots.length; index += 1) {
      const cellIndex = changedRoots[index]!;
      const entityId = makeCellEntity(cellIndex);
      if (this.visitedCells[cellIndex] === this.cellEpoch) {
        continue;
      }
      this.visitedCells[cellIndex] = this.cellEpoch;
      this.entityQueue[queueLength] = entityId;
      queueLength += 1;
      if (hasFormula(cellIndex)) {
        this.dirtyFormulaIds[dirtyLength] = cellIndex;
        dirtyLength += 1;
        const rank = cellStore.topoRanks[cellIndex] ?? 0;
        minRank = Math.min(minRank, rank);
        maxRank = Math.max(maxRank, rank);
      }
    }

    for (let queueIndex = 0; queueIndex < queueLength; queueIndex += 1) {
      const entityId = this.entityQueue[queueIndex]!;
      const dependents = graph.getDependents(entityId);
      if (isRangeEntity(entityId)) {
        rangeNodeVisits += 1;
      }
      for (let depIndex = 0; depIndex < dependents.length; depIndex += 1) {
        const dependent = dependents[depIndex]!;
        if (isRangeEntity(dependent)) {
          const rangeIndex = entityPayload(dependent);
          if (this.visitedRanges[rangeIndex] === this.rangeEpoch) {
            continue;
          }
          this.visitedRanges[rangeIndex] = this.rangeEpoch;
          this.entityQueue[queueLength] = dependent;
          queueLength += 1;
          continue;
        }

        const cellIndex = entityPayload(dependent);
        if (this.visitedCells[cellIndex] === this.cellEpoch) {
          continue;
        }
        this.visitedCells[cellIndex] = this.cellEpoch;
        this.entityQueue[queueLength] = dependent;
        queueLength += 1;
        if (!hasFormula(cellIndex)) {
          continue;
        }
        this.dirtyFormulaIds[dirtyLength] = cellIndex;
        dirtyLength += 1;
        const rank = cellStore.topoRanks[cellIndex] ?? 0;
        minRank = Math.min(minRank, rank);
        maxRank = Math.max(maxRank, rank);
      }
    }

    if (dirtyLength === 0) {
      return {
        orderedFormulaCellIndices: this.orderedDirty,
        orderedFormulaCount: 0,
        rangeNodeVisits,
      };
    }

    const rankSpan = maxRank - minRank + 1;
    this.ensureRankCapacity(rankSpan + 1, dirtyLength + 1);
    this.rankCounts.fill(0, 0, rankSpan);
    for (let index = 0; index < dirtyLength; index += 1) {
      const cellIndex = this.dirtyFormulaIds[index]!;
      const rank = (cellStore.topoRanks[cellIndex] ?? 0) - minRank;
      const nextCount = (this.rankCounts[rank] ?? 0) + 1;
      this.rankCounts[rank] = nextCount;
    }

    let prefix = 0;
    for (let index = 0; index < rankSpan; index += 1) {
      const count = this.rankCounts[index]!;
      this.rankCounts[index] = prefix;
      prefix += count;
    }

    for (let index = 0; index < dirtyLength; index += 1) {
      const cellIndex = this.dirtyFormulaIds[index]!;
      const rank = (cellStore.topoRanks[cellIndex] ?? 0) - minRank;
      const target = this.rankCounts[rank]!;
      this.orderedDirty[target] = cellIndex;
      this.rankCounts[rank] = target + 1;
    }

    return {
      orderedFormulaCellIndices: this.orderedDirty,
      orderedFormulaCount: dirtyLength,
      rangeNodeVisits,
    };
  }

  private ensureCellCapacity(size: number): void {
    if (size <= this.visitedCells.length) {
      return;
    }
    let capacity = this.visitedCells.length;
    while (capacity < size) {
      capacity *= 2;
    }
    this.visitedCells = grow(this.visitedCells, capacity);
    this.entityQueue = grow(this.entityQueue, capacity);
    this.dirtyFormulaIds = grow(this.dirtyFormulaIds, capacity);
  }

  private ensureRangeCapacity(size: number): void {
    if (size <= this.visitedRanges.length) {
      return;
    }
    let capacity = this.visitedRanges.length;
    while (capacity < size) {
      capacity *= 2;
    }
    this.visitedRanges = grow(this.visitedRanges, capacity);
  }

  private ensureRankCapacity(rankCountSize: number, dirtySize: number): void {
    if (rankCountSize > this.rankCounts.length) {
      let capacity = this.rankCounts.length;
      while (capacity < rankCountSize) {
        capacity *= 2;
      }
      this.rankCounts = grow(this.rankCounts, capacity);
    }
    if (dirtySize > this.orderedDirty.length) {
      let capacity = this.orderedDirty.length;
      while (capacity < dirtySize) {
        capacity *= 2;
      }
      this.orderedDirty = grow(this.orderedDirty, capacity);
    }
  }
}

function grow(buffer: U32, capacity: number): U32 {
  const next = new Uint32Array(capacity);
  next.set(buffer);
  return next as U32;
}
