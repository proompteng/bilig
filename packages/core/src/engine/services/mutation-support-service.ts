import { Effect } from 'effect'
import { ErrorCode, MAX_COLS, MAX_ROWS, ValueTag, type CellValue, type SelectionState } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import { CellFlags } from '../../cell-store.js'
import type { EdgeArena, EdgeSlice } from '../../edge-arena.js'
import { entityPayload, isRangeEntity, makeCellEntity, makeRangeEntity } from '../../entity-ids.js'
import { appendPackedCellIndex, growUint32 } from '../../engine-buffer-utils.js'
import { areCellValuesEqual, emptyValue, errorValue } from '../../engine-value-utils.js'
import type { EngineRuntimeState, SpillMaterialization, U32 } from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'

type DerivedOp = Extract<EngineOp, { kind: 'upsertSpillRange' | 'deleteSpillRange' | 'upsertPivotTable' | 'deletePivotTable' }>

function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
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

export interface EngineMutationSupportService {
  readonly beginMutationCollection: () => Effect.Effect<void, EngineMutationError>
  readonly markInputChanged: (cellIndex: number, count: number) => Effect.Effect<number, EngineMutationError>
  readonly markFormulaChanged: (cellIndex: number, count: number) => Effect.Effect<number, EngineMutationError>
  readonly markVolatileFormulasChanged: (count: number) => Effect.Effect<number, EngineMutationError>
  readonly markSpillRootsChanged: (cellIndices: readonly number[], count: number) => Effect.Effect<number, EngineMutationError>
  readonly markPivotRootsChanged: (cellIndices: readonly number[], count: number) => Effect.Effect<number, EngineMutationError>
  readonly markExplicitChanged: (cellIndex: number, count: number) => Effect.Effect<number, EngineMutationError>
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => Effect.Effect<U32, EngineMutationError>
  readonly composeEventChanges: (recalculated: U32, explicitChangedCount: number) => Effect.Effect<U32, EngineMutationError>
  readonly unionChangedSets: (...sets: Array<readonly number[] | U32>) => Effect.Effect<U32, EngineMutationError>
  readonly composeChangedRootsAndOrdered: (
    changedRoots: readonly number[] | U32,
    ordered: U32,
    orderedCount: number,
  ) => Effect.Effect<U32, EngineMutationError>
  readonly getChangedInputBuffer: () => Effect.Effect<U32, EngineMutationError>
  readonly ensureCellTracked: (sheetName: string, address: string) => Effect.Effect<number, EngineMutationError>
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => Effect.Effect<number, EngineMutationError>
  readonly clearOwnedSpill: (cellIndex: number) => Effect.Effect<number[], EngineMutationError>
  readonly materializeSpill: (
    cellIndex: number,
    arrayValue: { values: CellValue[]; rows: number; cols: number },
  ) => Effect.Effect<SpillMaterialization, EngineMutationError>
  readonly removeSheetRuntime: (
    sheetName: string,
    explicitChangedCount: number,
  ) => Effect.Effect<{ changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number }, EngineMutationError>
  readonly syncDynamicRanges: (formulaChangedCount: number) => Effect.Effect<number, EngineMutationError>
  readonly resetMaterializedCellScratch: (expectedSize: number) => Effect.Effect<void, EngineMutationError>
  readonly beginMutationCollectionNow: () => void
  readonly markInputChangedNow: (cellIndex: number, count: number) => number
  readonly markFormulaChangedNow: (cellIndex: number, count: number) => number
  readonly markVolatileFormulasChangedNow: (count: number) => number
  readonly markSpillRootsChangedNow: (cellIndices: readonly number[], count: number) => number
  readonly markPivotRootsChangedNow: (cellIndices: readonly number[], count: number) => number
  readonly markExplicitChangedNow: (cellIndex: number, count: number) => number
  readonly composeMutationRootsNow: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly composeEventChangesNow: (recalculated: U32, explicitChangedCount: number) => U32
  readonly unionChangedSetsNow: (...sets: Array<readonly number[] | U32>) => U32
  readonly composeChangedRootsAndOrderedNow: (changedRoots: readonly number[] | U32, ordered: U32, orderedCount: number) => U32
  readonly getChangedInputBufferNow: () => U32
  readonly ensureCellTrackedNow: (sheetName: string, address: string) => number
  readonly ensureCellTrackedByCoordsNow: (sheetId: number, row: number, col: number) => number
  readonly clearOwnedSpillNow: (cellIndex: number) => number[]
  readonly materializeSpillNow: (cellIndex: number, arrayValue: { values: CellValue[]; rows: number; cols: number }) => SpillMaterialization
  readonly removeSheetRuntimeNow: (
    sheetName: string,
    explicitChangedCount: number,
  ) => { changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number }
  readonly syncDynamicRangesNow: (formulaChangedCount: number) => number
  readonly resetMaterializedCellScratchNow: (expectedSize: number) => void
}

export function createEngineMutationSupportService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings' | 'formulas' | 'ranges'>
  readonly edgeArena: EdgeArena
  readonly reverseState: {
    reverseCellEdges: Array<EdgeSlice | undefined>
    reverseRangeEdges: Array<EdgeSlice | undefined>
  }
  readonly removeFormula: (cellIndex: number) => boolean
  readonly rebindFormulasForSheet: (sheetName: string, formulaChangedCount: number, candidates?: readonly number[] | U32) => number
  readonly getSelectionState: () => SelectionState
  readonly setSelection: (sheetName: string, address: string) => void
  readonly applyDerivedOp: (op: DerivedOp) => number[]
  readonly scheduleWasmProgramSync: () => void
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly collectFormulaDependents: (entityId: number) => Uint32Array
  readonly getChangedInputEpoch: () => number
  readonly setChangedInputEpoch: (next: number) => void
  readonly getChangedInputSeen: () => U32
  readonly setChangedInputSeen: (next: U32) => void
  readonly getChangedInputBuffer: () => U32
  readonly setChangedInputBuffer: (next: U32) => void
  readonly getChangedFormulaEpoch: () => number
  readonly setChangedFormulaEpoch: (next: number) => void
  readonly getChangedFormulaSeen: () => U32
  readonly setChangedFormulaSeen: (next: U32) => void
  readonly getChangedFormulaBuffer: () => U32
  readonly setChangedFormulaBuffer: (next: U32) => void
  readonly getChangedUnionEpoch: () => number
  readonly setChangedUnionEpoch: (next: number) => void
  readonly getChangedUnionSeen: () => U32
  readonly setChangedUnionSeen: (next: U32) => void
  readonly getChangedUnion: () => U32
  readonly setChangedUnion: (next: U32) => void
  readonly getMutationRoots: () => U32
  readonly setMutationRoots: (next: U32) => void
  readonly getMaterializedCellCount: () => number
  readonly setMaterializedCellCount: (next: number) => void
  readonly getMaterializedCells: () => U32
  readonly setMaterializedCells: (next: U32) => void
  readonly getExplicitChangedEpoch: () => number
  readonly setExplicitChangedEpoch: (next: number) => void
  readonly getExplicitChangedSeen: () => U32
  readonly setExplicitChangedSeen: (next: U32) => void
  readonly getExplicitChangedBuffer: () => U32
  readonly setExplicitChangedBuffer: (next: U32) => void
  readonly getImpactedFormulaEpoch: () => number
  readonly setImpactedFormulaEpoch: (next: number) => void
  readonly getImpactedFormulaSeen: () => U32
  readonly setImpactedFormulaSeen: (next: U32) => void
  readonly getImpactedFormulaBuffer: () => U32
  readonly setImpactedFormulaBuffer: (next: U32) => void
}): EngineMutationSupportService {
  const getReverseEdgeSlice = (entityId: number): EdgeSlice | undefined => {
    if (isRangeEntity(entityId)) {
      return args.reverseState.reverseRangeEdges[entityPayload(entityId)]
    }
    return args.reverseState.reverseCellEdges[entityPayload(entityId)]
  }

  const setReverseEdgeSlice = (entityId: number, slice: EdgeSlice): void => {
    const empty = slice.ptr < 0 || slice.len === 0
    if (isRangeEntity(entityId)) {
      args.reverseState.reverseRangeEdges[entityPayload(entityId)] = empty ? undefined : slice
      return
    }
    args.reverseState.reverseCellEdges[entityPayload(entityId)] = empty ? undefined : slice
  }

  const appendReverseEdge = (entityId: number, dependentEntityId: number): void => {
    const slice = getReverseEdgeSlice(entityId) ?? args.edgeArena.empty()
    setReverseEdgeSlice(entityId, args.edgeArena.appendUnique(slice, dependentEntityId))
  }

  const getEntityDependents = (entityId: number): Uint32Array =>
    args.edgeArena.readView(getReverseEdgeSlice(entityId) ?? args.edgeArena.empty())

  const pushMaterializedCell = (cellIndex: number): void => {
    const nextCount = args.getMaterializedCellCount() + 1
    if (nextCount > args.getMaterializedCells().length) {
      args.setMaterializedCells(growUint32(args.getMaterializedCells(), nextCount))
    }
    args.getMaterializedCells()[args.getMaterializedCellCount()] = cellIndex
    args.setMaterializedCellCount(nextCount)
  }

  const beginMutationCollectionNow = (): void => {
    advanceEpoch(args.getChangedInputEpoch(), args.setChangedInputEpoch, args.getChangedInputSeen())
    advanceEpoch(args.getChangedFormulaEpoch(), args.setChangedFormulaEpoch, args.getChangedFormulaSeen())
    advanceEpoch(args.getExplicitChangedEpoch(), args.setExplicitChangedEpoch, args.getExplicitChangedSeen())
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + 1)
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

  const ensureCellTrackedNow = (sheetName: string, address: string): number => {
    const ensured = args.state.workbook.ensureCellRecord(sheetName, address)
    if (ensured.created) {
      pushMaterializedCell(ensured.cellIndex)
    }
    return ensured.cellIndex
  }

  const ensureCellTrackedByCoordsNow = (sheetId: number, row: number, col: number): number => {
    const ensured = args.state.workbook.ensureCellAt(sheetId, row, col)
    if (ensured.created) {
      pushMaterializedCell(ensured.cellIndex)
    }
    return ensured.cellIndex
  }

  const clearSpillChildCell = (cellIndex: number): boolean => {
    const currentFlags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    const currentValue = args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id))
    if (currentValue.tag === ValueTag.Empty && (currentFlags & CellFlags.SpillChild) === 0) {
      return false
    }
    args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
    args.state.workbook.cellStore.flags[cellIndex] = currentFlags & ~CellFlags.SpillChild
    return true
  }

  const setSpillChildValue = (cellIndex: number, value: CellValue): boolean => {
    const currentValue = args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id))
    const currentFlags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    const nextFlags = (currentFlags & ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle)) | CellFlags.SpillChild
    if (areCellValuesEqual(currentValue, value) && currentFlags === nextFlags) {
      return false
    }
    args.state.workbook.cellStore.setValue(cellIndex, value, value.tag === ValueTag.String ? args.state.strings.intern(value.value) : 0)
    args.state.workbook.cellStore.flags[cellIndex] = nextFlags
    return true
  }

  const clearOwnedSpillNow = (cellIndex: number): number[] => {
    const sheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
    const address = args.state.workbook.getAddress(cellIndex)
    const spill = args.state.workbook.getSpill(sheetName, address)
    if (!spill) {
      return []
    }

    const owner = parseCellAddress(address, sheetName)
    const changedCellIndices: number[] = []
    for (let rowOffset = 0; rowOffset < spill.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < spill.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue
        }
        const childAddress = formatAddress(owner.row + rowOffset, owner.col + colOffset)
        const childIndex = args.state.workbook.getCellIndex(sheetName, childAddress)
        if (childIndex === undefined) {
          continue
        }
        if (clearSpillChildCell(childIndex)) {
          changedCellIndices.push(childIndex)
        }
      }
    }
    changedCellIndices.push(...args.applyDerivedOp({ kind: 'deleteSpillRange', sheetName, address }))
    return changedCellIndices
  }

  const materializeSpillNow = (
    cellIndex: number,
    arrayValue: { values: CellValue[]; rows: number; cols: number },
  ): SpillMaterialization => {
    const changedCellIndices = clearOwnedSpillNow(cellIndex)
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]!
    const sheetName = args.state.workbook.getSheetNameById(sheetId)
    const address = args.state.workbook.getAddress(cellIndex)
    const owner = parseCellAddress(address, sheetName)

    if (owner.row + arrayValue.rows > MAX_ROWS || owner.col + arrayValue.cols > MAX_COLS) {
      return { changedCellIndices, ownerValue: errorValue(ErrorCode.Spill) }
    }

    for (let rowOffset = 0; rowOffset < arrayValue.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < arrayValue.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue
        }
        const targetAddress = formatAddress(owner.row + rowOffset, owner.col + colOffset)
        const targetIndex = args.state.workbook.getCellIndex(sheetName, targetAddress)
        if (targetIndex === undefined) {
          continue
        }
        const targetValue = args.state.workbook.cellStore.getValue(targetIndex, (id) => args.state.strings.get(id))
        if (args.state.formulas.get(targetIndex) || targetValue.tag !== ValueTag.Empty) {
          return { changedCellIndices, ownerValue: errorValue(ErrorCode.Blocked) }
        }
      }
    }

    for (let rowOffset = 0; rowOffset < arrayValue.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < arrayValue.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue
        }
        const targetIndex = ensureCellTrackedByCoordsNow(sheetId, owner.row + rowOffset, owner.col + colOffset)
        const valueIndex = rowOffset * arrayValue.cols + colOffset
        const value = arrayValue.values[valueIndex] ?? emptyValue()
        if (setSpillChildValue(targetIndex, value)) {
          changedCellIndices.push(targetIndex)
        }
      }
    }

    if (arrayValue.rows > 1 || arrayValue.cols > 1) {
      changedCellIndices.push(
        ...args.applyDerivedOp({
          kind: 'upsertSpillRange',
          sheetName,
          address,
          rows: arrayValue.rows,
          cols: arrayValue.cols,
        }),
      )
    }

    return {
      changedCellIndices,
      ownerValue: arrayValue.values[0] ?? emptyValue(),
    }
  }

  const collectImpactedFormulasForCells = (cellIndices: readonly number[]): number => {
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + 1)
    advanceEpoch(args.getImpactedFormulaEpoch(), args.setImpactedFormulaEpoch, args.getImpactedFormulaSeen())

    let impactedCount = 0
    for (let cellCursor = 0; cellCursor < cellIndices.length; cellCursor += 1) {
      const cellIndex = cellIndices[cellCursor]!
      const dependents = args.collectFormulaDependents(makeCellEntity(cellIndex))
      for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
        const formulaCellIndex = dependents[dependentIndex]!
        if (args.getImpactedFormulaSeen()[formulaCellIndex] === args.getImpactedFormulaEpoch()) {
          continue
        }
        args.getImpactedFormulaSeen()[formulaCellIndex] = args.getImpactedFormulaEpoch()
        args.getImpactedFormulaBuffer()[impactedCount] = formulaCellIndex
        impactedCount += 1
      }
    }

    return impactedCount
  }

  const removeSheetRuntimeNow = (
    sheetName: string,
    explicitChangedCount: number,
  ): { changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number } => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return { changedInputCount: 0, formulaChangedCount: 0, explicitChangedCount }
    }

    const cellIndices: number[] = []
    sheet.grid.forEachCell((cellIndex) => {
      cellIndices.push(cellIndex)
    })
    const impactedCount = collectImpactedFormulasForCells(cellIndices)

    let changedInputCount = 0
    let formulaChangedCount = 0
    cellIndices.forEach((cellIndex) => {
      args.removeFormula(cellIndex)
      setReverseEdgeSlice(makeCellEntity(cellIndex), args.edgeArena.empty())
      args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.PendingDelete
      changedInputCount = markInputChangedNow(cellIndex, changedInputCount)
      explicitChangedCount = markExplicitChangedNow(cellIndex, explicitChangedCount)
    })

    args.state.workbook.deleteSheet(sheetName)
    if (args.getSelectionState().sheetName === sheetName) {
      const nextSheet = [...args.state.workbook.sheetsByName.values()].toSorted((left, right) => left.order - right.order)[0]
      args.setSelection(nextSheet?.name ?? sheetName, 'A1')
    }
    formulaChangedCount = args.rebindFormulasForSheet(
      sheetName,
      formulaChangedCount,
      args.getImpactedFormulaBuffer().subarray(0, impactedCount),
    )
    return { changedInputCount, formulaChangedCount, explicitChangedCount }
  }

  const syncDynamicRangesNow = (formulaChangedCount: number): number => {
    for (let index = 0; index < args.getMaterializedCellCount(); index += 1) {
      const cellIndex = args.getMaterializedCells()[index]!
      const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex] ?? 0
      if (sheetId === 0) {
        continue
      }
      const position = args.state.workbook.getCellPosition(cellIndex)
      const row = position?.row ?? args.state.workbook.cellStore.rows[cellIndex] ?? 0
      const col = position?.col ?? args.state.workbook.cellStore.cols[cellIndex] ?? 0
      const isFormulaCell = (args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0
      const rangeIndices = args.state.ranges.addDynamicMember(sheetId, row, col, cellIndex, isFormulaCell)
      if (rangeIndices.length > 0) {
        args.scheduleWasmProgramSync()
      }
      for (let rangeCursor = 0; rangeCursor < rangeIndices.length; rangeCursor += 1) {
        const rangeIndex = rangeIndices[rangeCursor]!
        const rangeEntity = makeRangeEntity(rangeIndex)
        appendReverseEdge(makeCellEntity(cellIndex), rangeEntity)
        const formulas = getEntityDependents(rangeEntity)
        for (let formulaCursor = 0; formulaCursor < formulas.length; formulaCursor += 1) {
          const formulaEntity = formulas[formulaCursor]!
          if (isRangeEntity(formulaEntity)) {
            continue
          }
          const formulaCellIndex = entityPayload(formulaEntity)
          const formula = args.state.formulas.get(formulaCellIndex)
          if (!formula) {
            continue
          }
          if (!isFormulaCell) {
            continue
          }
          const nextDependencyIndices = appendPackedCellIndex(formula.dependencyIndices, cellIndex)
          if (nextDependencyIndices !== formula.dependencyIndices) {
            formula.dependencyIndices = nextDependencyIndices
            formulaChangedCount = markFormulaChangedNow(formulaCellIndex, formulaChangedCount)
          }
        }
      }
    }
    return formulaChangedCount
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

  return {
    beginMutationCollection() {
      return Effect.try({
        try: () => {
          beginMutationCollectionNow()
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to begin mutation collection', cause),
            cause,
          }),
      })
    },
    markInputChanged(cellIndex, count) {
      return Effect.try({
        try: () => markInputChangedNow(cellIndex, count),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to mark input ${cellIndex} as changed`, cause),
            cause,
          }),
      })
    },
    markFormulaChanged(cellIndex, count) {
      return Effect.try({
        try: () => markFormulaChangedNow(cellIndex, count),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to mark formula ${cellIndex} as changed`, cause),
            cause,
          }),
      })
    },
    markVolatileFormulasChanged(count) {
      return Effect.try({
        try: () => {
          args.state.formulas.forEach((formula, cellIndex) => {
            if (!formula.compiled.volatile) {
              return
            }
            count = markFormulaChangedNow(cellIndex, count)
          })
          return count
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to mark volatile formulas as changed', cause),
            cause,
          }),
      })
    },
    markSpillRootsChanged(cellIndices, count) {
      return Effect.try({
        try: () => {
          for (let index = 0; index < cellIndices.length; index += 1) {
            count = markInputChangedNow(cellIndices[index]!, count)
          }
          return count
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to mark spill roots as changed', cause),
            cause,
          }),
      })
    },
    markPivotRootsChanged(cellIndices, count) {
      return Effect.try({
        try: () => {
          for (let index = 0; index < cellIndices.length; index += 1) {
            count = markInputChangedNow(cellIndices[index]!, count)
          }
          return count
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to mark pivot roots as changed', cause),
            cause,
          }),
      })
    },
    markExplicitChanged(cellIndex, count) {
      return Effect.try({
        try: () => markExplicitChangedNow(cellIndex, count),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to mark cell ${cellIndex} as explicit`, cause),
            cause,
          }),
      })
    },
    composeMutationRoots(changedInputCount, formulaChangedCount) {
      return Effect.try({
        try: () => composeMutationRootsNow(changedInputCount, formulaChangedCount),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to compose mutation roots', cause),
            cause,
          }),
      })
    },
    composeEventChanges(recalculated, explicitChangedCount) {
      return Effect.try({
        try: () => {
          advanceEpoch(args.getChangedUnionEpoch(), args.setChangedUnionEpoch, args.getChangedUnionSeen())
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
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to compose event changes', cause),
            cause,
          }),
      })
    },
    unionChangedSets(...sets) {
      return Effect.try({
        try: () => {
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
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to union changed sets', cause),
            cause,
          }),
      })
    },
    composeChangedRootsAndOrdered(changedRoots, ordered, orderedCount) {
      return Effect.try({
        try: () => {
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
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to compose changed roots and order', cause),
            cause,
          }),
      })
    },
    getChangedInputBuffer() {
      return Effect.try({
        try: () => args.getChangedInputBuffer(),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to read changed input buffer', cause),
            cause,
          }),
      })
    },
    ensureCellTracked(sheetName, address) {
      return Effect.try({
        try: () => ensureCellTrackedNow(sheetName, address),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to ensure cell ${sheetName}!${address}`, cause),
            cause,
          }),
      })
    },
    ensureCellTrackedByCoords(sheetId, row, col) {
      return Effect.try({
        try: () => ensureCellTrackedByCoordsNow(sheetId, row, col),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to ensure cell coordinates ${sheetId}:${row}:${col}`, cause),
            cause,
          }),
      })
    },
    clearOwnedSpill(cellIndex) {
      return Effect.try({
        try: () => clearOwnedSpillNow(cellIndex),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to clear spill ${cellIndex}`, cause),
            cause,
          }),
      })
    },
    materializeSpill(cellIndex, arrayValue) {
      return Effect.try({
        try: () => materializeSpillNow(cellIndex, arrayValue),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to materialize spill ${cellIndex}`, cause),
            cause,
          }),
      })
    },
    removeSheetRuntime(sheetName, explicitChangedCount) {
      return Effect.try({
        try: () => removeSheetRuntimeNow(sheetName, explicitChangedCount),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to remove sheet runtime for ${sheetName}`, cause),
            cause,
          }),
      })
    },
    syncDynamicRanges(formulaChangedCount) {
      return Effect.try({
        try: () => syncDynamicRangesNow(formulaChangedCount),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to sync dynamic ranges', cause),
            cause,
          }),
      })
    },
    resetMaterializedCellScratch(expectedSize) {
      return Effect.try({
        try: () => {
          args.setMaterializedCellCount(0)
          if (expectedSize > args.getMaterializedCells().length) {
            args.setMaterializedCells(growUint32(args.getMaterializedCells(), expectedSize))
          }
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to reset materialized cell scratch', cause),
            cause,
          }),
      })
    },
    beginMutationCollectionNow,
    markInputChangedNow,
    markFormulaChangedNow,
    markVolatileFormulasChangedNow(count) {
      args.state.formulas.forEach((formula, cellIndex) => {
        if (!formula.compiled.volatile) {
          return
        }
        count = markFormulaChangedNow(cellIndex, count)
      })
      return count
    },
    markSpillRootsChangedNow(cellIndices, count) {
      for (let index = 0; index < cellIndices.length; index += 1) {
        count = markInputChangedNow(cellIndices[index]!, count)
      }
      return count
    },
    markPivotRootsChangedNow(cellIndices, count) {
      for (let index = 0; index < cellIndices.length; index += 1) {
        count = markInputChangedNow(cellIndices[index]!, count)
      }
      return count
    },
    markExplicitChangedNow,
    composeMutationRootsNow,
    composeEventChangesNow(recalculated, explicitChangedCount) {
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
    },
    unionChangedSetsNow(...sets) {
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
    },
    composeChangedRootsAndOrderedNow(changedRoots, ordered, orderedCount) {
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
    },
    getChangedInputBufferNow: () => args.getChangedInputBuffer(),
    ensureCellTrackedNow,
    ensureCellTrackedByCoordsNow,
    clearOwnedSpillNow,
    materializeSpillNow,
    removeSheetRuntimeNow,
    syncDynamicRangesNow,
    resetMaterializedCellScratchNow(expectedSize) {
      args.setMaterializedCellCount(0)
      if (expectedSize > args.getMaterializedCells().length) {
        args.setMaterializedCells(growUint32(args.getMaterializedCells(), expectedSize))
      }
    },
  }
}
