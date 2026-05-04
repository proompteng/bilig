import { ValueTag } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

export function createOperationDirectFormulaDeltas(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'formulas' | 'counters'>
  readonly canSkipTerminalFormulaColumnVersion: (cellIndex: number) => boolean
  readonly canSkipDirectFormulaColumnVersion: (cellIndex: number) => boolean
}) {
  const applyDirectFormulaNumericDelta = (cellIndex: number, delta: number): boolean => {
    const cellStore = args.state.workbook.cellStore
    if (cellStore.tags[cellIndex] !== ValueTag.Number) {
      return false
    }
    const beforeNumber = cellStore.numbers[cellIndex] ?? 0
    const nextNumber = beforeNumber + delta
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.numbers[cellIndex] = nextNumber
    cellStore.stringIds[cellIndex] = 0
    cellStore.errors[cellIndex] = 0
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    if (cellStore.onSetValue) {
      cellStore.onSetValue(cellIndex)
    } else if (!Object.is(beforeNumber, nextNumber)) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
    return true
  }

  const applyTerminalDirectFormulaNumericDelta = (cellIndex: number, delta: number): boolean => {
    return applyTerminalDirectFormulaNumericDeltaAndReturn(cellIndex, delta) !== undefined
  }

  const applyTerminalDirectFormulaNumericDeltaAndReturn = (cellIndex: number, delta: number): number | undefined => {
    const cellStore = args.state.workbook.cellStore
    if (cellStore.tags[cellIndex] !== ValueTag.Number) {
      return undefined
    }
    const nextNumber = (cellStore.numbers[cellIndex] ?? 0) + delta
    const flags = cellStore.flags[cellIndex] ?? 0
    if ((flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
      cellStore.flags[cellIndex] = flags & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    }
    cellStore.numbers[cellIndex] = nextNumber
    if ((cellStore.stringIds[cellIndex] ?? 0) !== 0) {
      cellStore.stringIds[cellIndex] = 0
    }
    if ((cellStore.errors[cellIndex] ?? 0) !== 0) {
      cellStore.errors[cellIndex] = 0
    }
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    return nextNumber
  }

  const tryApplyDirectFormulaDeltas = (collection: DirectFormulaIndexCollection, captureChanged = true): U32 | undefined => {
    if (!collection.hasCompleteDeltas()) {
      return undefined
    }
    const cellStore = args.state.workbook.cellStore
    const changed = captureChanged ? new Uint32Array(collection.size) : EMPTY_CHANGED_CELLS
    let directAggregateDeltaApplicationCount = 0
    let directScalarDeltaApplicationCount = 0
    let canUseTerminalFormulaWrites = true
    for (let index = 0; index < collection.size; index += 1) {
      const cellIndex = collection.getCellIndexAt(index)
      if (
        ((cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0 ||
        cellStore.tags[cellIndex] !== ValueTag.Number ||
        collection.getDeltaAt(index) === undefined
      ) {
        return undefined
      }
      const formula = args.state.formulas.get(cellIndex)
      if (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) {
        directAggregateDeltaApplicationCount += 1
      }
      if (formula?.directScalar !== undefined) {
        directScalarDeltaApplicationCount += 1
      }
      if (canUseTerminalFormulaWrites && !args.canSkipTerminalFormulaColumnVersion(cellIndex)) {
        canUseTerminalFormulaWrites = false
      }
      if (captureChanged) {
        changed[index] = cellIndex
      }
    }
    const applyDeltas = (): void => {
      for (let index = 0; index < collection.size; index += 1) {
        const cellIndex = captureChanged ? changed[index]! : collection.getCellIndexAt(index)
        const applied = canUseTerminalFormulaWrites
          ? applyTerminalDirectFormulaNumericDelta(cellIndex, collection.getDeltaAt(index)!)
          : applyDirectFormulaNumericDelta(cellIndex, collection.getDeltaAt(index)!)
        if (!applied) {
          throw new Error('Failed to apply direct formula delta')
        }
      }
    }
    if (canUseTerminalFormulaWrites) {
      applyDeltas()
    } else {
      args.state.workbook.withBatchedColumnVersionUpdates(applyDeltas)
    }
    if (directAggregateDeltaApplicationCount > 0) {
      addEngineCounter(args.state.counters, 'directAggregateDeltaApplications', directAggregateDeltaApplicationCount)
    }
    if (directScalarDeltaApplicationCount > 0) {
      addEngineCounter(args.state.counters, 'directScalarDeltaApplications', directScalarDeltaApplicationCount)
    }
    return changed
  }

  const tryApplyDirectScalarDeltas = (collection: DirectFormulaIndexCollection, captureChanged = true): U32 | undefined => {
    const constantDelta = collection.getConstantScalarDelta()
    if (constantDelta === undefined && !collection.hasCompleteScalarDeltas()) {
      return undefined
    }
    const cellStore = args.state.workbook.cellStore
    const changed = captureChanged ? new Uint32Array(collection.size) : EMPTY_CHANGED_CELLS
    const hasValidatedTerminalWrites = collection.hasValidatedScalarDeltaCells()
    let canUseTerminalFormulaWrites = hasValidatedTerminalWrites
    if (!hasValidatedTerminalWrites) {
      canUseTerminalFormulaWrites = true
      for (let index = 0; index < collection.size; index += 1) {
        const cellIndex = collection.getCellIndexAt(index)
        if (((cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0 || cellStore.tags[cellIndex] !== ValueTag.Number) {
          return undefined
        }
        if (canUseTerminalFormulaWrites && !args.canSkipDirectFormulaColumnVersion(cellIndex)) {
          canUseTerminalFormulaWrites = false
        }
        if (captureChanged) {
          changed[index] = cellIndex
        }
      }
    }
    const applyDeltas = (): void => {
      for (let index = 0; index < collection.size; index += 1) {
        const cellIndex = hasValidatedTerminalWrites || !captureChanged ? collection.getCellIndexAt(index) : changed[index]!
        if (hasValidatedTerminalWrites && captureChanged) {
          changed[index] = cellIndex
        }
        const delta = constantDelta ?? collection.getScalarDeltaAt(index)
        if (delta === undefined) {
          throw new Error('Missing direct scalar delta')
        }
        if (canUseTerminalFormulaWrites) {
          if (!applyTerminalDirectFormulaNumericDelta(cellIndex, delta)) {
            throw new Error('Failed to apply direct scalar delta')
          }
        } else {
          const beforeNumber = cellStore.numbers[cellIndex] ?? 0
          const nextNumber = beforeNumber + delta
          cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
          cellStore.numbers[cellIndex] = nextNumber
          cellStore.stringIds[cellIndex] = 0
          cellStore.errors[cellIndex] = 0
          cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
          if (cellStore.onSetValue) {
            cellStore.onSetValue(cellIndex)
          } else if (!Object.is(beforeNumber, nextNumber)) {
            args.state.workbook.notifyCellValueWritten(cellIndex)
          }
        }
      }
    }
    if (canUseTerminalFormulaWrites) {
      applyDeltas()
    } else {
      args.state.workbook.withBatchedColumnVersionUpdates(applyDeltas)
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', collection.size)
    return changed
  }

  return {
    applyDirectFormulaNumericDelta,
    applyTerminalDirectFormulaNumericDelta,
    applyTerminalDirectFormulaNumericDeltaAndReturn,
    tryApplyDirectFormulaDeltas,
    tryApplyDirectScalarDeltas,
  } as const
}
