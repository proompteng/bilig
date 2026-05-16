import type { EntityId } from '@bilig/protocol'
import type { CellStore } from '../../cell-store.js'
import { isRangeEntity, makeCellEntity, makeExactLookupColumnEntity, makeSortedLookupColumnEntity } from '../../entity-ids.js'

type U32 = Uint32Array

export interface OperationPostRecalcDirectFormulaIndexAccess {
  readonly has: (cellIndex: number) => boolean
  readonly hasCoveredDirectRangeInput: (cellIndex: number) => boolean
  readonly hasCoveredDirectFormulaInput: (cellIndex: number) => boolean
}

export interface OperationDirtyTraversalSheetAccess {
  readonly structureVersion: number
  readonly logical: {
    readonly getCellVisiblePosition: (cellIndex: number) => { readonly row: number; readonly col: number } | undefined
  }
}

export interface OperationDirtyTraversalAccess {
  readonly workbook: {
    readonly cellStore: CellStore
    readonly getSheetNameById: (sheetId: number) => string | undefined
    readonly getSheetById: (sheetId: number) => OperationDirtyTraversalSheetAccess | undefined
  }
  readonly getSingleEntityDependent: (entityId: EntityId) => number
  readonly getEntityDependents: (entityId: EntityId) => U32
  readonly collectRegionFormulaDependentsForCell: (sheetName: string, row: number, col: number) => U32
  readonly collectAffectedDirectRangeDependents: (request: {
    readonly sheetName: string
    readonly row: number
    readonly col: number
  }) => readonly number[]
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
}

export function canSkipOperationDirtyTraversalForChangedInputs(input: {
  readonly changedInputCellIndices: U32
  readonly changedInputCount: number
  readonly postRecalcDirectFormulaIndices?: OperationPostRecalcDirectFormulaIndexAccess | undefined
  readonly options?: {
    readonly lookupHandledInputCellIndices?: readonly number[]
  }
  readonly access: OperationDirtyTraversalAccess
}): boolean {
  const lookupInputCovered = (cellIndex: number): boolean => {
    const covered = input.options?.lookupHandledInputCellIndices
    if (covered === undefined) {
      return false
    }
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return true
      }
    }
    return false
  }
  const lookupDependentsArePostRecalcDirect = (sheetId: number, col: number): boolean => {
    if (input.postRecalcDirectFormulaIndices === undefined) {
      return false
    }
    const exactLookupDependents = input.access.hasTrackedExactLookupDependents(sheetId, col)
      ? input.access.getEntityDependents(makeExactLookupColumnEntity(sheetId, col))
      : new Uint32Array()
    for (let dependentIndex = 0; dependentIndex < exactLookupDependents.length; dependentIndex += 1) {
      if (!input.postRecalcDirectFormulaIndices.has(exactLookupDependents[dependentIndex]!)) {
        return false
      }
    }
    const sortedLookupDependents = input.access.hasTrackedSortedLookupDependents(sheetId, col)
      ? input.access.getEntityDependents(makeSortedLookupColumnEntity(sheetId, col))
      : new Uint32Array()
    for (let dependentIndex = 0; dependentIndex < sortedLookupDependents.length; dependentIndex += 1) {
      if (!input.postRecalcDirectFormulaIndices.has(sortedLookupDependents[dependentIndex]!)) {
        return false
      }
    }
    return true
  }
  const dependentsArePostRecalcDirect = (dependents: U32): boolean => {
    if (input.postRecalcDirectFormulaIndices === undefined) {
      return false
    }
    for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
      const dependent = dependents[dependentIndex]!
      if (isRangeEntity(dependent)) {
        return false
      }
      if (!input.postRecalcDirectFormulaIndices.has(dependent)) {
        return false
      }
    }
    return true
  }
  const rangeDependentsArePostRecalcDirect = (cellIndex: number, requireTrackedRangeDependents = false): boolean => {
    if (input.postRecalcDirectFormulaIndices === undefined) {
      return false
    }
    const cellStore = input.access.workbook.cellStore
    const sheetId = cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      return true
    }
    const sheetName = input.access.workbook.getSheetNameById(sheetId)
    if (!sheetName) {
      return false
    }
    const sheet = input.access.workbook.getSheetById(sheetId)
    let row: number | undefined
    let col: number | undefined
    if (!sheet || sheet.structureVersion === 1) {
      row = cellStore.rows[cellIndex]
      col = cellStore.cols[cellIndex]
    } else {
      const position = sheet.logical.getCellVisiblePosition(cellIndex)
      row = position?.row
      col = position?.col
    }
    if (row === undefined || col === undefined) {
      return false
    }
    if (
      (input.access.hasTrackedExactLookupDependents(sheetId, col) || input.access.hasTrackedSortedLookupDependents(sheetId, col)) &&
      !lookupInputCovered(cellIndex) &&
      !lookupDependentsArePostRecalcDirect(sheetId, col)
    ) {
      return false
    }
    if (input.postRecalcDirectFormulaIndices.hasCoveredDirectRangeInput(cellIndex)) {
      return true
    }
    if (!input.access.hasTrackedDirectRangeDependents(sheetId, col)) {
      return !requireTrackedRangeDependents
    }
    const regionDependents = input.access.collectRegionFormulaDependentsForCell(sheetName, row, col)
    for (let dependentIndex = 0; dependentIndex < regionDependents.length; dependentIndex += 1) {
      if (!input.postRecalcDirectFormulaIndices.has(regionDependents[dependentIndex]!)) {
        return false
      }
    }
    const directRangeDependents = input.access.collectAffectedDirectRangeDependents({ sheetName, row, col })
    for (let dependentIndex = 0; dependentIndex < directRangeDependents.length; dependentIndex += 1) {
      if (!input.postRecalcDirectFormulaIndices.has(directRangeDependents[dependentIndex]!)) {
        return false
      }
    }
    return true
  }
  for (let index = 0; index < input.changedInputCount; index += 1) {
    const cellIndex = input.changedInputCellIndices[index]!
    if (input.postRecalcDirectFormulaIndices?.hasCoveredDirectFormulaInput(cellIndex) === true) {
      if (!rangeDependentsArePostRecalcDirect(cellIndex)) {
        return false
      }
      continue
    }
    const singleDependent = input.access.getSingleEntityDependent(makeCellEntity(cellIndex))
    if (singleDependent === -1) {
      if (!rangeDependentsArePostRecalcDirect(cellIndex)) {
        return false
      }
      continue
    }
    if (singleDependent >= 0) {
      if (isRangeEntity(singleDependent)) {
        if (!rangeDependentsArePostRecalcDirect(cellIndex, true)) {
          return false
        }
        continue
      }
      if (input.postRecalcDirectFormulaIndices?.has(singleDependent) === true) {
        if (!rangeDependentsArePostRecalcDirect(cellIndex)) {
          return false
        }
        continue
      }
      return false
    }
    const dependents = input.access.getEntityDependents(makeCellEntity(cellIndex))
    if (dependents.length > 0 && !dependentsArePostRecalcDirect(dependents)) {
      return false
    }
    if (!rangeDependentsArePostRecalcDirect(cellIndex)) {
      return false
    }
  }
  return true
}

export function operationChangedInputsNeedRegionQueryIndices(input: {
  readonly changedInputCellIndices: U32
  readonly changedInputCount: number
  readonly postRecalcDirectFormulaIndices?: OperationPostRecalcDirectFormulaIndexAccess | undefined
  readonly access: Pick<OperationDirtyTraversalAccess, 'workbook' | 'hasTrackedDirectRangeDependents'>
}): boolean {
  const cellStore = input.access.workbook.cellStore
  for (let index = 0; index < input.changedInputCount; index += 1) {
    const cellIndex = input.changedInputCellIndices[index]!
    if (input.postRecalcDirectFormulaIndices?.hasCoveredDirectRangeInput(cellIndex) === true) {
      continue
    }
    const sheetId = cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      continue
    }
    const sheet = input.access.workbook.getSheetById(sheetId)
    const col = !sheet || sheet.structureVersion === 1 ? cellStore.cols[cellIndex] : sheet.logical.getCellVisiblePosition(cellIndex)?.col
    if (col !== undefined && input.access.hasTrackedDirectRangeDependents(sheetId, col)) {
      return true
    }
  }
  return false
}
