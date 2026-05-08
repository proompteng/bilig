import type { CellValue } from '@bilig/protocol'
import { makeExactLookupColumnEntity, makeSortedLookupColumnEntity } from '../../entity-ids.js'
import type { EngineRuntimeState, RuntimeDirectLookupDescriptor } from '../runtime-state.js'
import {
  canSkipUniformApproximateNumericTailWrite,
  canSkipUniformApproximateNumericTailWriteFromCurrentResult,
  canSkipUniformExactNumericTailWriteFromCurrentResult,
  directLookupRowBounds,
  normalizeApproximateNumericValue,
  normalizeApproximateTextValue,
  sameExactNumericValue,
} from './direct-lookup-helpers.js'
import type { OperationLookupAccess } from './operation-lookup-access.js'

type UniformLookupTailPatchTarget = Extract<
  RuntimeDirectLookupDescriptor,
  { kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric' }
>

interface LookupNumericColumnWritePlan {
  readonly handled: boolean
  readonly tailPatchTarget?: UniformLookupTailPatchTarget
}

export interface OperationLookupPlanner {
  readonly planSingleExactLookupNumericColumnWrite: (
    formulaCellIndex: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ) => LookupNumericColumnWritePlan
  readonly canSkipSingleExactLookupNumericColumnWrite: (
    formulaCellIndex: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ) => boolean
  readonly planExactLookupNumericColumnWrite: (
    sheetId: number,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ) => LookupNumericColumnWritePlan
  readonly canSkipExactLookupNumericColumnWrite: (
    sheetId: number,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ) => boolean
  readonly planSingleApproximateLookupNumericColumnWrite: (
    formulaCellIndex: number,
    sheetName: string,
    row: number,
    col: number,
    oldNumeric: number,
    newNumeric: number,
  ) => LookupNumericColumnWritePlan
  readonly canSkipSingleApproximateLookupNumericColumnWrite: (
    formulaCellIndex: number,
    sheetName: string,
    row: number,
    col: number,
    oldNumeric: number,
    newNumeric: number,
  ) => boolean
  readonly planApproximateLookupNumericColumnWrite: (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ) => LookupNumericColumnWritePlan
  readonly canSkipApproximateLookupNumericColumnWrite: (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ) => boolean
  readonly canSkipApproximateLookupNewNumericColumnWrite: (sheetId: number, col: number, row: number) => boolean
  readonly canPatchUniformLookupTailWrite: (
    directLookup: UniformLookupTailPatchTarget,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ) => boolean
  readonly patchUniformLookupTailWrites: (request: {
    sheetId: number
    col: number
    row: number
    oldNumeric: number
    newNumeric: number
    exact: boolean
    sorted: boolean
  }) => { exact: boolean; sorted: boolean }
  readonly canSkipApproximateLookupDirtyMark: (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' | 'approximate-uniform-numeric' }>,
    request: {
      sheetName: string
      row: number
      col: number
      oldValue: CellValue
      newValue: CellValue
      oldStringId?: number
      newStringId?: number
    },
  ) => boolean
}

export function createOperationLookupPlanner(args: {
  readonly state: Pick<EngineRuntimeState, 'formulas' | 'workbook' | 'strings'>
  readonly access: Pick<
    OperationLookupAccess,
    | 'readExactNumericValueForLookup'
    | 'readApproximateNumericValueForLookup'
    | 'readCellValueForLookup'
    | 'isLocallySortedNumericWrite'
    | 'isLocallySortedTextWrite'
  >
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly getEntityDependents: (entityId: number) => ArrayLike<number>
}): OperationLookupPlanner {
  const planSingleExactLookupNumericColumnWrite = (
    formulaCellIndex: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): LookupNumericColumnWritePlan => {
    const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
    if (directLookup?.kind !== 'exact' && directLookup?.kind !== 'exact-uniform-numeric') {
      return { handled: false }
    }
    if (directLookup.kind === 'exact-uniform-numeric' && directLookup.tailPatch !== undefined) {
      return { handled: false }
    }
    const { rowStart, rowEnd } = directLookupRowBounds(directLookup)
    if (row < rowStart || row > rowEnd) {
      return { handled: true }
    }
    if (
      directLookup.kind === 'exact-uniform-numeric' &&
      canSkipUniformExactNumericTailWriteFromCurrentResult(
        args.state.workbook.cellStore,
        formulaCellIndex,
        directLookup,
        row,
        oldNumeric,
        newNumeric,
      )
    ) {
      return { handled: true, tailPatchTarget: directLookup }
    }
    const operandNumeric = args.access.readExactNumericValueForLookup(directLookup.operandCellIndex)
    if (operandNumeric === undefined) {
      return { handled: false }
    }
    return { handled: !sameExactNumericValue(oldNumeric, operandNumeric) && !sameExactNumericValue(newNumeric, operandNumeric) }
  }

  const canSkipSingleExactLookupNumericColumnWrite = (
    formulaCellIndex: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    return planSingleExactLookupNumericColumnWrite(formulaCellIndex, row, oldNumeric, newNumeric).handled
  }

  const planExactLookupNumericColumnWrite = (
    sheetId: number,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): LookupNumericColumnWritePlan => {
    const lookupEntity = makeExactLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    if (singleDependent === -1) {
      return { handled: true }
    }
    if (singleDependent >= 0) {
      return planSingleExactLookupNumericColumnWrite(singleDependent, row, oldNumeric, newNumeric)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipSingleExactLookupNumericColumnWrite(dependents[index]!, row, oldNumeric, newNumeric)) {
        return { handled: false }
      }
    }
    return { handled: true }
  }

  const canSkipExactLookupNumericColumnWrite = (
    sheetId: number,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    const lookupEntity = makeExactLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      return canSkipSingleExactLookupNumericColumnWrite(singleDependent, row, oldNumeric, newNumeric)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipSingleExactLookupNumericColumnWrite(dependents[index]!, row, oldNumeric, newNumeric)) {
        return false
      }
    }
    return true
  }

  const planSingleApproximateLookupNumericColumnWrite = (
    formulaCellIndex: number,
    sheetName: string,
    row: number,
    col: number,
    oldNumeric: number,
    newNumeric: number,
  ): LookupNumericColumnWritePlan => {
    const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
    if (directLookup?.kind !== 'approximate' && directLookup?.kind !== 'approximate-uniform-numeric') {
      return { handled: false }
    }
    if (directLookup.kind === 'approximate-uniform-numeric' && directLookup.tailPatch !== undefined) {
      return { handled: false }
    }
    const { rowStart, rowEnd } = directLookupRowBounds(directLookup)
    if (row < rowStart || row > rowEnd) {
      return { handled: true }
    }
    if (
      directLookup.kind === 'approximate-uniform-numeric' &&
      canSkipUniformApproximateNumericTailWriteFromCurrentResult(
        args.state.workbook.cellStore,
        formulaCellIndex,
        directLookup,
        row,
        oldNumeric,
        newNumeric,
      )
    ) {
      return { handled: true, tailPatchTarget: directLookup }
    }
    const operandNumeric = args.access.readApproximateNumericValueForLookup(directLookup.operandCellIndex)
    if (operandNumeric === undefined) {
      return { handled: false }
    }
    const matchMode = directLookup.matchMode
    if (directLookup.kind === 'approximate-uniform-numeric') {
      if (canSkipUniformApproximateNumericTailWrite(directLookup, row, operandNumeric, oldNumeric, newNumeric)) {
        return { handled: true, tailPatchTarget: directLookup }
      }
    }
    if (matchMode === 1) {
      return {
        handled:
          oldNumeric > operandNumeric &&
          newNumeric > operandNumeric &&
          args.access.isLocallySortedNumericWrite(sheetName, row, col, rowStart, rowEnd, matchMode, newNumeric),
      }
    }
    return {
      handled:
        oldNumeric < operandNumeric &&
        newNumeric < operandNumeric &&
        args.access.isLocallySortedNumericWrite(sheetName, row, col, rowStart, rowEnd, matchMode, newNumeric),
    }
  }

  const canSkipSingleApproximateLookupNumericColumnWrite = (
    formulaCellIndex: number,
    sheetName: string,
    row: number,
    col: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    return planSingleApproximateLookupNumericColumnWrite(formulaCellIndex, sheetName, row, col, oldNumeric, newNumeric).handled
  }

  const planApproximateLookupNumericColumnWrite = (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): LookupNumericColumnWritePlan => {
    const lookupEntity = makeSortedLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    if (singleDependent === -1) {
      return { handled: true }
    }
    if (singleDependent >= 0) {
      return planSingleApproximateLookupNumericColumnWrite(singleDependent, sheetName, row, col, oldNumeric, newNumeric)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipSingleApproximateLookupNumericColumnWrite(dependents[index]!, sheetName, row, col, oldNumeric, newNumeric)) {
        return { handled: false }
      }
    }
    return { handled: true }
  }

  const canSkipApproximateLookupNumericColumnWrite = (
    sheetId: number,
    sheetName: string,
    col: number,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    const lookupEntity = makeSortedLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      return canSkipSingleApproximateLookupNumericColumnWrite(singleDependent, sheetName, row, col, oldNumeric, newNumeric)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipSingleApproximateLookupNumericColumnWrite(dependents[index]!, sheetName, row, col, oldNumeric, newNumeric)) {
        return false
      }
    }
    return true
  }

  const canSkipApproximateLookupNewNumericColumnWrite = (sheetId: number, col: number, row: number): boolean => {
    const lookupEntity = makeSortedLookupColumnEntity(sheetId, col)
    const singleDependent = args.getSingleEntityDependent(lookupEntity)
    const canSkipDependent = (formulaCellIndex: number): boolean => {
      const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
      if (directLookup?.kind !== 'approximate' && directLookup?.kind !== 'approximate-uniform-numeric') {
        return false
      }
      const { rowStart, rowEnd } = directLookupRowBounds(directLookup)
      return row < rowStart || row > rowEnd
    }
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      return canSkipDependent(singleDependent)
    }
    const dependents = args.getEntityDependents(lookupEntity)
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canSkipDependent(dependents[index]!)) {
        return false
      }
    }
    return true
  }

  const canPatchUniformLookupTailWrite = (
    directLookup: UniformLookupTailPatchTarget,
    row: number,
    oldNumeric: number,
    newNumeric: number,
  ): boolean => {
    if (directLookup.kind === 'approximate-uniform-numeric' && directLookup.repeatedRunLength !== undefined) {
      return false
    }
    if (directLookup.tailPatch !== undefined || row !== directLookup.rowEnd || directLookup.length < 2) {
      return false
    }
    const expectedOldTail = directLookup.start + directLookup.step * (directLookup.length - 1)
    if (!sameExactNumericValue(oldNumeric, expectedOldTail)) {
      return false
    }
    return directLookup.step > 0 ? newNumeric > oldNumeric : newNumeric < oldNumeric
  }

  const patchUniformLookupTailWrites = (request: {
    sheetId: number
    col: number
    row: number
    oldNumeric: number
    newNumeric: number
    exact: boolean
    sorted: boolean
  }): { exact: boolean; sorted: boolean } => {
    const patchDependents = (entityId: number, kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric'): boolean => {
      const singleDependent = args.getSingleEntityDependent(entityId)
      if (singleDependent === -1) {
        return true
      }
      const currentSheet = args.state.workbook.getSheetById(request.sheetId)
      const currentColumnVersion = currentSheet?.columnVersions[request.col] ?? 0
      if (singleDependent >= 0) {
        const directLookup = args.state.formulas.get(singleDependent)?.directLookup
        if (
          directLookup?.kind !== kind ||
          !canPatchUniformLookupTailWrite(directLookup, request.row, request.oldNumeric, request.newNumeric)
        ) {
          return false
        }
        directLookup.tailPatch = {
          row: request.row,
          oldNumeric: request.oldNumeric,
          newNumeric: request.newNumeric,
          columnVersion: currentColumnVersion,
        }
        return true
      }
      const dependents = args.getEntityDependents(entityId)
      if (dependents.length === 0) {
        return true
      }
      for (let index = 0; index < dependents.length; index += 1) {
        const directLookup = args.state.formulas.get(dependents[index]!)?.directLookup
        if (
          directLookup?.kind !== kind ||
          !canPatchUniformLookupTailWrite(directLookup, request.row, request.oldNumeric, request.newNumeric)
        ) {
          return false
        }
      }
      for (let index = 0; index < dependents.length; index += 1) {
        const directLookup = args.state.formulas.get(dependents[index]!)?.directLookup
        if (directLookup?.kind === kind) {
          directLookup.tailPatch = {
            row: request.row,
            oldNumeric: request.oldNumeric,
            newNumeric: request.newNumeric,
            columnVersion: currentColumnVersion,
          }
        }
      }
      return true
    }

    return {
      exact: request.exact && patchDependents(makeExactLookupColumnEntity(request.sheetId, request.col), 'exact-uniform-numeric'),
      sorted: request.sorted && patchDependents(makeSortedLookupColumnEntity(request.sheetId, request.col), 'approximate-uniform-numeric'),
    }
  }

  const canSkipApproximateLookupDirtyMark = (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' | 'approximate-uniform-numeric' }>,
    request: {
      sheetName: string
      row: number
      col: number
      oldValue: CellValue
      newValue: CellValue
      oldStringId?: number
      newStringId?: number
    },
  ): boolean => {
    const rowStart = directLookup.kind === 'approximate' ? directLookup.prepared.rowStart : directLookup.rowStart
    const rowEnd = directLookup.kind === 'approximate' ? directLookup.prepared.rowEnd : directLookup.rowEnd
    const matchMode = directLookup.kind === 'approximate' ? directLookup.matchMode : directLookup.matchMode
    const operandNumeric = args.access.readApproximateNumericValueForLookup(directLookup.operandCellIndex)
    if (operandNumeric !== undefined) {
      const oldNumeric = normalizeApproximateNumericValue(request.oldValue)
      const newNumeric = normalizeApproximateNumericValue(request.newValue)
      if (oldNumeric === undefined || newNumeric === undefined) {
        return false
      }
      if (matchMode === 1) {
        return (
          oldNumeric > operandNumeric &&
          newNumeric > operandNumeric &&
          args.access.isLocallySortedNumericWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode, newNumeric)
        )
      }
      return (
        oldNumeric < operandNumeric &&
        newNumeric < operandNumeric &&
        args.access.isLocallySortedNumericWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode, newNumeric)
      )
    }
    const operand = args.access.readCellValueForLookup(directLookup.operandCellIndex)
    const operandText = normalizeApproximateTextValue(operand.value, (id) => args.state.strings.get(id), operand.stringId)
    if (operandText === undefined) {
      return false
    }
    const oldText = normalizeApproximateTextValue(request.oldValue, (id) => args.state.strings.get(id), request.oldStringId)
    const newText = normalizeApproximateTextValue(request.newValue, (id) => args.state.strings.get(id), request.newStringId)
    if (oldText === undefined || newText === undefined) {
      return false
    }
    if (matchMode === 1) {
      return (
        oldText > operandText &&
        newText > operandText &&
        args.access.isLocallySortedTextWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode, newText)
      )
    }
    return (
      oldText < operandText &&
      newText < operandText &&
      args.access.isLocallySortedTextWrite(request.sheetName, request.row, request.col, rowStart, rowEnd, matchMode, newText)
    )
  }

  return {
    planSingleExactLookupNumericColumnWrite,
    canSkipSingleExactLookupNumericColumnWrite,
    planExactLookupNumericColumnWrite,
    canSkipExactLookupNumericColumnWrite,
    planSingleApproximateLookupNumericColumnWrite,
    canSkipSingleApproximateLookupNumericColumnWrite,
    planApproximateLookupNumericColumnWrite,
    canSkipApproximateLookupNumericColumnWrite,
    canSkipApproximateLookupNewNumericColumnWrite,
    canPatchUniformLookupTailWrite,
    patchUniformLookupTailWrites,
    canSkipApproximateLookupDirtyMark,
  }
}
