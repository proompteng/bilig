import { compileCriteriaMatcher, matchesCompiledCriteria } from '@bilig/formula'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { emptyValue } from '../../engine-value-utils.js'
import type { EngineRuntimeState, RuntimeDirectCriteriaDescriptor } from '../runtime-state.js'
import { normalizeApproximateTextValue } from './direct-lookup-helpers.js'
import { directAggregateNumericContribution, directCriteriaValueString } from './direct-formula-recalc-helpers.js'

export interface OperationLookupAccess {
  readonly readCellValueForLookup: (cellIndex: number | undefined) => { value: CellValue; stringId: number | undefined }
  readonly readApproximateNumericValueForLookup: (cellIndex: number | undefined) => number | undefined
  readonly readExactNumericValueForLookup: (cellIndex: number | undefined) => number | undefined
  readonly readCellValueAtForLookup: (sheetName: string, row: number, col: number) => { value: CellValue; stringId: number | undefined }
  readonly readDirectCriteriaOperandValue: (operand: RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion']) => CellValue
  readonly directCriteriaMatchesChangedAggregateRow: (
    directCriteria: RuntimeDirectCriteriaDescriptor,
    aggregateRange: NonNullable<RuntimeDirectCriteriaDescriptor['aggregateRange']>,
    requestRow: number,
  ) => boolean | undefined
  readonly tryDirectCriteriaSumDelta: (
    directCriteria: RuntimeDirectCriteriaDescriptor,
    request: {
      sheetName: string
      row: number
      col: number
      oldValue?: CellValue
      newValue?: CellValue
    },
  ) => number | undefined
  readonly readApproximateNumericValueAtForLookup: (sheetName: string, row: number, col: number) => number | undefined
  readonly isLocallySortedNumericWrite: (
    sheetName: string,
    row: number,
    col: number,
    rowStart: number,
    rowEnd: number,
    matchMode: 1 | -1,
    current: number,
  ) => boolean
  readonly isLocallySortedTextWrite: (
    sheetName: string,
    row: number,
    col: number,
    rowStart: number,
    rowEnd: number,
    matchMode: 1 | -1,
    current: string,
  ) => boolean
}

export function createOperationLookupAccess(args: {
  readonly workbook: EngineRuntimeState['workbook']
  readonly strings: EngineRuntimeState['strings']
}): OperationLookupAccess {
  const readCellValueForLookup = (cellIndex: number | undefined): { value: CellValue; stringId: number | undefined } => {
    if (cellIndex === undefined) {
      return { value: emptyValue(), stringId: undefined }
    }
    const stringId = args.workbook.cellStore.stringIds[cellIndex]
    return {
      value: args.workbook.cellStore.getValue(cellIndex, (id) => args.strings.get(id)),
      stringId,
    }
  }

  const readApproximateNumericValueForLookup = (cellIndex: number | undefined): number | undefined => {
    if (cellIndex === undefined) {
      return 0
    }
    const cellStore = args.workbook.cellStore
    switch ((cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
      case ValueTag.Empty:
        return 0
      case ValueTag.Number: {
        const value = cellStore.numbers[cellIndex] ?? 0
        return Object.is(value, -0) ? 0 : value
      }
      case ValueTag.Boolean:
        return (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0
      case ValueTag.String:
      case ValueTag.Error:
        return undefined
    }
  }

  const readExactNumericValueForLookup = (cellIndex: number | undefined): number | undefined => {
    if (cellIndex === undefined) {
      return undefined
    }
    const cellStore = args.workbook.cellStore
    if (((cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) !== ValueTag.Number) {
      return undefined
    }
    const value = cellStore.numbers[cellIndex] ?? 0
    return Object.is(value, -0) ? 0 : value
  }

  const readCellValueAtForLookup = (sheetName: string, row: number, col: number): { value: CellValue; stringId: number | undefined } => {
    const sheet = args.workbook.getSheet(sheetName)
    if (!sheet) {
      return { value: emptyValue(), stringId: undefined }
    }
    if (sheet.structureVersion === 1) {
      const cellIndex = sheet.grid.getPhysical(row, col)
      return readCellValueForLookup(cellIndex === -1 ? undefined : cellIndex)
    }
    return readCellValueForLookup(sheet.logical.getVisibleCell(row, col))
  }

  const readDirectCriteriaOperandValue = (operand: RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion']): CellValue => {
    if (operand.kind === 'literal') {
      return operand.value
    }
    const value = readCellValueForLookup(operand.cellIndex).value
    if (operand.kind === 'cell') {
      return value
    }
    if (value.tag === ValueTag.Error) {
      return value
    }
    return { tag: ValueTag.String, value: `${operand.prefix}${directCriteriaValueString(value)}${operand.suffix}`, stringId: 0 }
  }

  const directCriteriaMatchesChangedAggregateRow = (
    directCriteria: RuntimeDirectCriteriaDescriptor,
    aggregateRange: NonNullable<RuntimeDirectCriteriaDescriptor['aggregateRange']>,
    requestRow: number,
  ): boolean | undefined => {
    const rowOffset = requestRow - aggregateRange.rowStart
    for (let index = 0; index < directCriteria.criteriaPairs.length; index += 1) {
      const pair = directCriteria.criteriaPairs[index]!
      const criteria = readDirectCriteriaOperandValue(pair.criterion)
      if (criteria.tag === ValueTag.Error) {
        return undefined
      }
      const candidate = readCellValueAtForLookup(pair.range.sheetName, pair.range.rowStart + rowOffset, pair.range.col).value
      if (!matchesCompiledCriteria(candidate, compileCriteriaMatcher(criteria))) {
        return false
      }
    }
    return true
  }

  const tryDirectCriteriaSumDelta = (
    directCriteria: RuntimeDirectCriteriaDescriptor,
    request: {
      sheetName: string
      row: number
      col: number
      oldValue?: CellValue
      newValue?: CellValue
    },
  ): number | undefined => {
    const aggregateRange = directCriteria.aggregateRange
    if (
      directCriteria.aggregateKind !== 'sum' ||
      (directCriteria.resultTransforms?.length ?? 0) > 0 ||
      aggregateRange === undefined ||
      aggregateRange.sheetName !== request.sheetName ||
      aggregateRange.col !== request.col ||
      request.row < aggregateRange.rowStart ||
      request.row > aggregateRange.rowEnd ||
      request.oldValue === undefined ||
      request.newValue === undefined
    ) {
      return undefined
    }
    const oldContribution = directAggregateNumericContribution(request.oldValue)
    const newContribution = directAggregateNumericContribution(request.newValue)
    if (oldContribution === undefined || newContribution === undefined) {
      return undefined
    }
    const matchesRow = directCriteriaMatchesChangedAggregateRow(directCriteria, aggregateRange, request.row)
    if (matchesRow === undefined) {
      return undefined
    }
    return matchesRow ? newContribution - oldContribution : 0
  }

  const readApproximateNumericValueAtForLookup = (sheetName: string, row: number, col: number): number | undefined => {
    const sheet = args.workbook.getSheet(sheetName)
    if (!sheet) {
      return 0
    }
    if (sheet.structureVersion === 1) {
      const cellIndex = sheet.grid.getPhysical(row, col)
      return readApproximateNumericValueForLookup(cellIndex === -1 ? undefined : cellIndex)
    }
    return readApproximateNumericValueForLookup(sheet.logical.getVisibleCell(row, col))
  }

  const isLocallySortedNumericWrite = (
    sheetName: string,
    row: number,
    col: number,
    rowStart: number,
    rowEnd: number,
    matchMode: 1 | -1,
    current: number,
  ): boolean => {
    if (row > rowStart) {
      const previous = readApproximateNumericValueAtForLookup(sheetName, row - 1, col)
      if (previous === undefined || (matchMode === 1 ? previous > current : previous < current)) {
        return false
      }
    }
    if (row < rowEnd) {
      const next = readApproximateNumericValueAtForLookup(sheetName, row + 1, col)
      if (next === undefined || (matchMode === 1 ? current > next : current < next)) {
        return false
      }
    }
    return true
  }

  const isLocallySortedTextWrite = (
    sheetName: string,
    row: number,
    col: number,
    rowStart: number,
    rowEnd: number,
    matchMode: 1 | -1,
    current: string,
  ): boolean => {
    if (row > rowStart) {
      const previousCell = readCellValueAtForLookup(sheetName, row - 1, col)
      const previous = normalizeApproximateTextValue(previousCell.value, (id) => args.strings.get(id), previousCell.stringId)
      if (previous === undefined || (matchMode === 1 ? previous > current : previous < current)) {
        return false
      }
    }
    if (row < rowEnd) {
      const nextCell = readCellValueAtForLookup(sheetName, row + 1, col)
      const next = normalizeApproximateTextValue(nextCell.value, (id) => args.strings.get(id), nextCell.stringId)
      if (next === undefined || (matchMode === 1 ? current > next : current < next)) {
        return false
      }
    }
    return true
  }

  return {
    readCellValueForLookup,
    readApproximateNumericValueForLookup,
    readExactNumericValueForLookup,
    readCellValueAtForLookup,
    readDirectCriteriaOperandValue,
    directCriteriaMatchesChangedAggregateRow,
    tryDirectCriteriaSumDelta,
    readApproximateNumericValueAtForLookup,
    isLocallySortedNumericWrite,
    isLocallySortedTextWrite,
  }
}
