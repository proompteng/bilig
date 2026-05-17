import { growUint32 } from '../../engine-buffer-utils.js'
import type { FormulaTable } from '../../formula-table.js'
import type { U32 } from '../runtime-state.js'

export interface MutationSupportChangeSetFormulaRecord {
  readonly cellIndex: number
  readonly compiled: {
    readonly volatile: boolean
  }
}

function advanceEpoch(current: number, setEpoch: (next: number) => void, seen: U32): number {
  if (current >= 0xffff_fffe) {
    setEpoch(1)
    seen.fill(0)
    return 1
  }
  const next = current + 1
  setEpoch(next)
  return next
}

export interface MutationSupportChangeSetTracker {
  readonly beginMutationCollectionNow: () => void
  readonly markInputChangedNow: (cellIndex: number, count: number) => number
  readonly markFormulaChangedNow: (cellIndex: number, count: number) => number
  readonly markExplicitChangedNow: (cellIndex: number, count: number) => number
  readonly markVolatileFormulasChangedNow: (count: number) => number
  readonly composeMutationRootsNow: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly composeEventChangesNow: (recalculated: U32, explicitChangedCount: number) => U32
  readonly composeDisjointEventChangesNow: (recalculated: U32, explicitChangedCount: number) => U32
  readonly unionChangedSetsNow: (...sets: Array<readonly number[] | U32>) => U32
  readonly composeChangedRootsAndOrderedNow: (changedRoots: readonly number[] | U32, ordered: U32, orderedCount: number) => U32
  readonly getChangedInputBufferNow: () => U32
  readonly pushMaterializedCell: (cellIndex: number) => void
  readonly resetMaterializedCellScratchNow: (expectedSize: number) => void
}

export function createMutationSupportChangeSetTracker(args: {
  readonly formulas: FormulaTable<MutationSupportChangeSetFormulaRecord>
  readonly getVolatileFormulaCellIndices?: () => Iterable<number>
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly getCellStoreSize: () => number
  readonly getChangedInputEpoch: () => number
  readonly setChangedInputEpoch: (next: number) => void
  readonly getChangedInputSeen: () => U32
  readonly getChangedInputBuffer: () => U32
  readonly getChangedFormulaEpoch: () => number
  readonly setChangedFormulaEpoch: (next: number) => void
  readonly getChangedFormulaSeen: () => U32
  readonly getChangedFormulaBuffer: () => U32
  readonly getChangedUnionEpoch: () => number
  readonly setChangedUnionEpoch: (next: number) => void
  readonly getChangedUnionSeen: () => U32
  readonly getChangedUnion: () => U32
  readonly getMutationRoots: () => U32
  readonly getMaterializedCellCount: () => number
  readonly setMaterializedCellCount: (next: number) => void
  readonly getMaterializedCells: () => U32
  readonly setMaterializedCells: (next: U32) => void
  readonly getExplicitChangedEpoch: () => number
  readonly setExplicitChangedEpoch: (next: number) => void
  readonly getExplicitChangedSeen: () => U32
  readonly getExplicitChangedBuffer: () => U32
}): MutationSupportChangeSetTracker {
  const pushMaterializedCell = (cellIndex: number): void => {
    const nextCount = args.getMaterializedCellCount() + 1
    if (nextCount > args.getMaterializedCells().length) {
      args.setMaterializedCells(growUint32(args.getMaterializedCells(), nextCount))
    }
    args.getMaterializedCells()[args.getMaterializedCellCount()] = cellIndex
    args.setMaterializedCellCount(nextCount)
  }

  const resetMaterializedCellScratchNow = (expectedSize: number): void => {
    args.setMaterializedCellCount(0)
    if (expectedSize > args.getMaterializedCells().length) {
      args.setMaterializedCells(growUint32(args.getMaterializedCells(), expectedSize))
    }
  }

  const beginMutationCollectionNow = (): void => {
    advanceEpoch(args.getChangedInputEpoch(), args.setChangedInputEpoch, args.getChangedInputSeen())
    advanceEpoch(args.getChangedFormulaEpoch(), args.setChangedFormulaEpoch, args.getChangedFormulaSeen())
    advanceEpoch(args.getExplicitChangedEpoch(), args.setExplicitChangedEpoch, args.getExplicitChangedSeen())
    args.ensureRecalcScratchCapacity(args.getCellStoreSize() + 1)
  }

  const markInputChangedNow = (cellIndex: number, count: number): number => {
    if (args.getChangedInputSeen()[cellIndex] === args.getChangedInputEpoch()) {
      return count
    }
    args.getChangedInputSeen()[cellIndex] = args.getChangedInputEpoch()
    args.getChangedInputBuffer()[count] = cellIndex
    return count + 1
  }

  const markFormulaChangedNow = (cellIndex: number, count: number): number => {
    if (args.getChangedFormulaSeen()[cellIndex] === args.getChangedFormulaEpoch()) {
      return count
    }
    args.getChangedFormulaSeen()[cellIndex] = args.getChangedFormulaEpoch()
    args.getChangedFormulaBuffer()[count] = cellIndex
    return count + 1
  }

  const markExplicitChangedNow = (cellIndex: number, count: number): number => {
    if (args.getExplicitChangedSeen()[cellIndex] === args.getExplicitChangedEpoch()) {
      return count
    }
    args.getExplicitChangedSeen()[cellIndex] = args.getExplicitChangedEpoch()
    args.getExplicitChangedBuffer()[count] = cellIndex
    return count + 1
  }

  const markVolatileFormulasChangedNow = (count: number): number => {
    const volatileFormulaCellIndices = args.getVolatileFormulaCellIndices?.()
    if (volatileFormulaCellIndices) {
      for (const cellIndex of volatileFormulaCellIndices) {
        const formula = args.formulas.get(cellIndex)
        if (formula?.compiled.volatile) {
          count = markFormulaChangedNow(cellIndex, count)
        }
      }
      return count
    }
    args.formulas.forEach((formula, cellIndex) => {
      if (formula.compiled.volatile) {
        count = markFormulaChangedNow(cellIndex, count)
      }
    })
    return count
  }

  const composeMutationRootsNow = (changedInputCount: number, formulaChangedCount: number): U32 => {
    const total = changedInputCount + formulaChangedCount
    args.ensureRecalcScratchCapacity(total + 1)
    for (let index = 0; index < changedInputCount; index += 1) {
      args.getMutationRoots()[index] = args.getChangedInputBuffer()[index]!
    }
    for (let index = 0; index < formulaChangedCount; index += 1) {
      args.getMutationRoots()[changedInputCount + index] = args.getChangedFormulaBuffer()[index]!
    }
    return args.getMutationRoots().subarray(0, total)
  }

  const composeDisjointEventChangesNow = (recalculated: U32, explicitChangedCount: number): U32 => {
    if (explicitChangedCount === 0) {
      return recalculated
    }
    const total = explicitChangedCount + recalculated.length
    args.ensureRecalcScratchCapacity(total + 1)
    const changed = args.getChangedUnion()
    for (let index = 0; index < explicitChangedCount; index += 1) {
      changed[index] = args.getExplicitChangedBuffer()[index]!
    }
    changed.set(recalculated, explicitChangedCount)
    return changed.subarray(0, total)
  }

  const composeEventChangesNow = (recalculated: U32, explicitChangedCount: number): U32 => {
    advanceEpoch(args.getChangedUnionEpoch(), args.setChangedUnionEpoch, args.getChangedUnionSeen())
    if (explicitChangedCount === 0) {
      return recalculated
    }
    if (explicitChangedCount === 1 && recalculated.length === 0) {
      args.getChangedUnion()[0] = args.getExplicitChangedBuffer()[0]!
      return args.getChangedUnion().subarray(0, 1)
    }
    if (explicitChangedCount === 1 && recalculated.length === 1) {
      const explicitCellIndex = args.getExplicitChangedBuffer()[0]!
      const recalculatedCellIndex = recalculated[0]!
      args.getChangedUnion()[0] = explicitCellIndex
      if (explicitCellIndex === recalculatedCellIndex) {
        return args.getChangedUnion().subarray(0, 1)
      }
      args.getChangedUnion()[1] = recalculatedCellIndex
      return args.getChangedUnion().subarray(0, 2)
    }
    if (explicitChangedCount === 1 && recalculated.length === 2) {
      const explicitCellIndex = args.getExplicitChangedBuffer()[0]!
      const firstRecalculated = recalculated[0]!
      const secondRecalculated = recalculated[1]!
      args.getChangedUnion()[0] = explicitCellIndex
      if (firstRecalculated === explicitCellIndex) {
        if (secondRecalculated === explicitCellIndex) {
          return args.getChangedUnion().subarray(0, 1)
        }
        args.getChangedUnion()[1] = secondRecalculated
        return args.getChangedUnion().subarray(0, 2)
      }
      if (secondRecalculated === explicitCellIndex) {
        args.getChangedUnion()[1] = firstRecalculated
        return args.getChangedUnion().subarray(0, 2)
      }
      args.getChangedUnion()[1] = firstRecalculated
      if (firstRecalculated === secondRecalculated) {
        return args.getChangedUnion().subarray(0, 2)
      }
      args.getChangedUnion()[2] = secondRecalculated
      return args.getChangedUnion().subarray(0, 3)
    }
    let changedCount = 0
    for (let index = 0; index < explicitChangedCount; index += 1) {
      const cellIndex = args.getExplicitChangedBuffer()[index]!
      if (args.getChangedUnionSeen()[cellIndex] === args.getChangedUnionEpoch()) {
        continue
      }
      args.getChangedUnionSeen()[cellIndex] = args.getChangedUnionEpoch()
      args.getChangedUnion()[changedCount] = cellIndex
      changedCount += 1
    }
    for (let index = 0; index < recalculated.length; index += 1) {
      const cellIndex = recalculated[index]!
      if (args.getChangedUnionSeen()[cellIndex] === args.getChangedUnionEpoch()) {
        continue
      }
      args.getChangedUnionSeen()[cellIndex] = args.getChangedUnionEpoch()
      args.getChangedUnion()[changedCount] = cellIndex
      changedCount += 1
    }
    return args.getChangedUnion().subarray(0, changedCount)
  }

  const unionChangedSetsNow = (...sets: Array<readonly number[] | U32>): U32 => {
    advanceEpoch(args.getChangedUnionEpoch(), args.setChangedUnionEpoch, args.getChangedUnionSeen())
    let changedCount = 0
    for (let setIndex = 0; setIndex < sets.length; setIndex += 1) {
      const set = sets[setIndex]!
      for (let index = 0; index < set.length; index += 1) {
        const cellIndex = set[index]!
        if (args.getChangedUnionSeen()[cellIndex] === args.getChangedUnionEpoch()) {
          continue
        }
        args.getChangedUnionSeen()[cellIndex] = args.getChangedUnionEpoch()
        args.getChangedUnion()[changedCount] = cellIndex
        changedCount += 1
      }
    }
    return args.getChangedUnion().subarray(0, changedCount)
  }

  const composeChangedRootsAndOrderedNow = (changedRoots: readonly number[] | U32, ordered: U32, orderedCount: number): U32 => {
    advanceEpoch(args.getChangedUnionEpoch(), args.setChangedUnionEpoch, args.getChangedUnionSeen())
    let changedCount = 0
    for (let index = 0; index < changedRoots.length; index += 1) {
      const cellIndex = changedRoots[index]!
      if (args.getChangedUnionSeen()[cellIndex] === args.getChangedUnionEpoch()) {
        continue
      }
      args.getChangedUnionSeen()[cellIndex] = args.getChangedUnionEpoch()
      args.getChangedUnion()[changedCount] = cellIndex
      changedCount += 1
    }
    for (let index = 0; index < orderedCount; index += 1) {
      const cellIndex = ordered[index]!
      if (args.getChangedUnionSeen()[cellIndex] === args.getChangedUnionEpoch()) {
        continue
      }
      args.getChangedUnionSeen()[cellIndex] = args.getChangedUnionEpoch()
      args.getChangedUnion()[changedCount] = cellIndex
      changedCount += 1
    }
    return args.getChangedUnion().subarray(0, changedCount)
  }

  return {
    beginMutationCollectionNow,
    markInputChangedNow,
    markFormulaChangedNow,
    markExplicitChangedNow,
    markVolatileFormulasChangedNow,
    composeMutationRootsNow,
    composeEventChangesNow,
    composeDisjointEventChangesNow,
    unionChangedSetsNow,
    composeChangedRootsAndOrderedNow,
    getChangedInputBufferNow: () => args.getChangedInputBuffer(),
    pushMaterializedCell,
    resetMaterializedCellScratchNow,
  }
}
