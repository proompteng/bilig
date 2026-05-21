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
  const hasActiveDirectScalarFormula = (cellIndex: number): boolean => args.state.formulas.get(cellIndex)?.directScalar !== undefined

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

  const applyDirectFormulaNumericDeltaBatch = (cellIndices: readonly number[] | U32, delta: number): boolean => {
    const cellStore = args.state.workbook.cellStore
    const tags = cellStore.tags
    for (let index = 0; index < cellIndices.length; index += 1) {
      if (tags[cellIndices[index]!] !== ValueTag.Number) {
        return false
      }
    }
    if (cellIndices.length === 0) {
      return true
    }

    const flags = cellStore.flags
    const numbers = cellStore.numbers
    const stringIds = cellStore.stringIds
    const errors = cellStore.errors
    const versions = cellStore.versions
    const sheetIds = cellStore.sheetIds
    const cols = cellStore.cols
    const formulaOutputFlags = CellFlags.SpillChild | CellFlags.PivotOutput
    const clearFormulaOutputFlags = ~formulaOutputFlags
    let sharedPhysicalSheetId = -1
    let sharedPhysicalCol = -1
    let canNotifySharedPhysicalColumn = true

    for (let index = 0; index < cellIndices.length; index += 1) {
      const cellIndex = cellIndices[index]!
      const sheetId = sheetIds[cellIndex] ?? -1
      const col = cols[cellIndex] ?? -1
      if (canNotifySharedPhysicalColumn) {
        const sheet = args.state.workbook.getSheetById(sheetId)
        if (
          sheet === undefined ||
          sheet.structureVersion !== 1 ||
          (sharedPhysicalSheetId !== -1 && sharedPhysicalSheetId !== sheetId) ||
          (sharedPhysicalCol !== -1 && sharedPhysicalCol !== col)
        ) {
          canNotifySharedPhysicalColumn = false
        } else {
          sharedPhysicalSheetId = sheetId
          sharedPhysicalCol = col
        }
      }
      const currentFlags = flags[cellIndex] ?? 0
      if ((currentFlags & formulaOutputFlags) !== 0) {
        flags[cellIndex] = currentFlags & clearFormulaOutputFlags
      }
      const beforeNumber = numbers[cellIndex] ?? 0
      const nextNumber = beforeNumber + delta
      numbers[cellIndex] = nextNumber
      if ((stringIds[cellIndex] ?? 0) !== 0) {
        stringIds[cellIndex] = 0
      }
      if ((errors[cellIndex] ?? 0) !== 0) {
        errors[cellIndex] = 0
      }
      versions[cellIndex] = (versions[cellIndex] ?? 0) + 1
      if (!canNotifySharedPhysicalColumn) {
        if (cellStore.onSetValue) {
          cellStore.onSetValue(cellIndex)
        } else if (!Object.is(beforeNumber, nextNumber)) {
          args.state.workbook.notifyCellValueWritten(cellIndex)
        }
      }
    }
    if (canNotifySharedPhysicalColumn) {
      args.state.workbook.notifyColumnsWritten(sharedPhysicalSheetId, Uint32Array.of(sharedPhysicalCol))
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
    const constantDelta = collection.getConstantDelta()
    const cellStore = args.state.workbook.cellStore
    const flags = cellStore.flags
    const tags = cellStore.tags
    const numbers = cellStore.numbers
    const stringIds = cellStore.stringIds
    const errors = cellStore.errors
    const versions = cellStore.versions
    const changed = captureChanged ? new Uint32Array(collection.size) : EMPTY_CHANGED_CELLS
    let directAggregateDeltaApplicationCount = 0
    let directScalarDeltaApplicationCount = 0
    let canUseTerminalFormulaWrites = true
    for (let index = 0; index < collection.size; index += 1) {
      const cellIndex = collection.getCellIndexAt(index)
      const formula = args.state.formulas.get(cellIndex)
      if (
        ((flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0 ||
        tags[cellIndex] !== ValueTag.Number ||
        (constantDelta === undefined && collection.getDeltaAt(index) === undefined) ||
        (formula?.directAggregate === undefined && formula?.directCriteria === undefined && formula?.directScalar === undefined)
      ) {
        return undefined
      }
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
      const formulaOutputFlags = CellFlags.SpillChild | CellFlags.PivotOutput
      const clearFormulaOutputFlags = ~formulaOutputFlags
      for (let index = 0; index < collection.size; index += 1) {
        const cellIndex = captureChanged ? changed[index]! : collection.getCellIndexAt(index)
        const delta = constantDelta ?? collection.getDeltaAt(index)!
        const beforeNumber = numbers[cellIndex] ?? 0
        const nextNumber = beforeNumber + delta
        const currentFlags = flags[cellIndex] ?? 0
        if ((currentFlags & formulaOutputFlags) !== 0) {
          flags[cellIndex] = currentFlags & clearFormulaOutputFlags
        }
        numbers[cellIndex] = nextNumber
        if ((stringIds[cellIndex] ?? 0) !== 0) {
          stringIds[cellIndex] = 0
        }
        if ((errors[cellIndex] ?? 0) !== 0) {
          errors[cellIndex] = 0
        }
        versions[cellIndex] = (versions[cellIndex] ?? 0) + 1
        if (!canUseTerminalFormulaWrites) {
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
    const hasValidatedTerminalWrites = collection.hasValidatedScalarDeltaCells()
    const hasTrustedDirectScalarFormulas = collection.hasTrustedDirectScalarDeltaCells()
    if (collection.hasCleanScalarDeltaCells()) {
      const cellIndices = collection.getCellIndicesForRead()
      if (!hasTrustedDirectScalarFormulas) {
        for (let index = 0; index < cellIndices.length; index += 1) {
          if (!hasActiveDirectScalarFormula(cellIndices[index]!)) {
            return undefined
          }
        }
      }
      const changed = captureChanged
        ? cellIndices instanceof Uint32Array
          ? cellIndices
          : Uint32Array.from(cellIndices)
        : EMPTY_CHANGED_CELLS
      const numbers = cellStore.numbers
      const versions = cellStore.versions
      for (let index = 0; index < cellIndices.length; index += 1) {
        const cellIndex = cellIndices[index]!
        numbers[cellIndex] = numbers[cellIndex]! + (constantDelta ?? collection.getScalarDeltaAt(index)!)
        versions[cellIndex] = versions[cellIndex]! + 1
      }
      addEngineCounter(args.state.counters, 'directScalarDeltaApplications', collection.size)
      return changed
    }
    if (constantDelta !== undefined && hasValidatedTerminalWrites) {
      const cellIndices = collection.getCellIndicesForRead()
      if (!hasTrustedDirectScalarFormulas) {
        for (let index = 0; index < cellIndices.length; index += 1) {
          if (!hasActiveDirectScalarFormula(cellIndices[index]!)) {
            return undefined
          }
        }
      }
      const changed = captureChanged
        ? cellIndices instanceof Uint32Array
          ? cellIndices
          : Uint32Array.from(cellIndices)
        : EMPTY_CHANGED_CELLS
      const flags = cellStore.flags
      const numbers = cellStore.numbers
      const versions = cellStore.versions
      const stringIds = cellStore.stringIds
      const errors = cellStore.errors
      const formulaOutputFlags = CellFlags.SpillChild | CellFlags.PivotOutput
      const clearFormulaOutputFlags = ~formulaOutputFlags
      for (let index = 0; index < cellIndices.length; index += 1) {
        const cellIndex = cellIndices[index]!
        const currentFlags = flags[cellIndex]!
        if ((currentFlags & formulaOutputFlags) !== 0) {
          flags[cellIndex] = currentFlags & clearFormulaOutputFlags
        }
        numbers[cellIndex] = numbers[cellIndex]! + constantDelta
        if (stringIds[cellIndex] !== 0) {
          stringIds[cellIndex] = 0
        }
        if (errors[cellIndex] !== 0) {
          errors[cellIndex] = 0
        }
        versions[cellIndex] = versions[cellIndex]! + 1
      }
      addEngineCounter(args.state.counters, 'directScalarDeltaApplications', collection.size)
      return changed
    }
    const changed = captureChanged ? new Uint32Array(collection.size) : EMPTY_CHANGED_CELLS
    let canUseTerminalFormulaWrites = hasValidatedTerminalWrites
    if (!hasValidatedTerminalWrites) {
      canUseTerminalFormulaWrites = true
      for (let index = 0; index < collection.size; index += 1) {
        const cellIndex = collection.getCellIndexAt(index)
        if (
          ((cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0 ||
          cellStore.tags[cellIndex] !== ValueTag.Number ||
          !hasActiveDirectScalarFormula(cellIndex)
        ) {
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
    applyDirectFormulaNumericDeltaBatch,
    applyTerminalDirectFormulaNumericDelta,
    applyTerminalDirectFormulaNumericDeltaAndReturn,
    tryApplyDirectFormulaDeltas,
    tryApplyDirectScalarDeltas,
  } as const
}
