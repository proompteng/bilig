import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { RuntimeDirectCriteriaDescriptor, RuntimeDirectScalarOperand } from '../runtime-state.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { EngineRuntimeColumnStoreService } from './runtime-column-store-service.js'
import { directErrorResult } from './formula-evaluation-helpers.js'

function readIndexOffset(args: {
  readonly operand: RuntimeDirectScalarOperand
  readonly readCellValueByIndex: (cellIndex: number | undefined) => CellValue
}): { kind: 'number'; value: number } | CellValue {
  const operand = args.operand
  if (operand.kind === 'literal-number') {
    return { kind: 'number', value: operand.value }
  }
  if (operand.kind === 'error') {
    return directErrorResult(operand.code)
  }
  const value = args.readCellValueByIndex(operand.cellIndex)
  switch (value.tag) {
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.Boolean:
      return { kind: 'number', value: value.value ? 1 : 0 }
    case ValueTag.Empty:
      return { kind: 'number', value: 0 }
    case ValueTag.Error:
      return value
    case ValueTag.String:
      return directErrorResult(ErrorCode.Value)
  }
}

export function tryEvaluateDirectIndexOffset(args: {
  readonly directCriteria: RuntimeDirectCriteriaDescriptor
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly readCellValueByIndex: (cellIndex: number | undefined) => CellValue
}): CellValue | undefined {
  if (args.directCriteria.offsetOperand === undefined || args.directCriteria.criteriaPairs.length !== 0) {
    return undefined
  }
  const aggregateRange = args.directCriteria.aggregateRange
  if (!aggregateRange) {
    return undefined
  }
  const offset = readIndexOffset({
    operand: args.directCriteria.offsetOperand,
    readCellValueByIndex: args.readCellValueByIndex,
  })
  if ('tag' in offset) {
    return offset
  }
  if (!Number.isFinite(offset.value)) {
    return directErrorResult(ErrorCode.Value)
  }
  const rowOffset = Math.trunc(offset.value)
  if (rowOffset < 1 || rowOffset > aggregateRange.length) {
    return directErrorResult(ErrorCode.Ref)
  }
  return args.runtimeColumnStore
    .getColumnView({
      sheetName: aggregateRange.sheetName,
      rowStart: aggregateRange.rowStart,
      rowEnd: aggregateRange.rowEnd,
      col: aggregateRange.col,
    })
    .readCellValueAt(rowOffset - 1)
}

export function tryEvaluateDirectIndexExactMatch(args: {
  readonly directCriteria: RuntimeDirectCriteriaDescriptor
  readonly exactLookup: Pick<ExactColumnIndexService, 'prepareVectorLookup' | 'findPreparedVectorMatch'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly lookupValue: CellValue
}): CellValue | undefined {
  if (
    args.directCriteria.aggregateKind !== 'first' ||
    args.directCriteria.firstMatchMode !== 'exact-lookup' ||
    args.directCriteria.aggregateRange === undefined ||
    args.directCriteria.criteriaPairs.length !== 1
  ) {
    return undefined
  }
  const exactPair = args.directCriteria.criteriaPairs[0]!
  const exactMatch = args.exactLookup.findPreparedVectorMatch({
    lookupValue: args.lookupValue,
    prepared: args.exactLookup.prepareVectorLookup({
      sheetName: exactPair.range.sheetName,
      rowStart: exactPair.range.rowStart,
      rowEnd: exactPair.range.rowEnd,
      col: exactPair.range.col,
    }),
    searchMode: 1,
  })
  if (!exactMatch.handled) {
    return undefined
  }
  if (exactMatch.position === undefined) {
    return directErrorResult(ErrorCode.NA)
  }
  return args.runtimeColumnStore
    .getColumnView({
      sheetName: args.directCriteria.aggregateRange.sheetName,
      rowStart: args.directCriteria.aggregateRange.rowStart,
      rowEnd: args.directCriteria.aggregateRange.rowEnd,
      col: args.directCriteria.aggregateRange.col,
    })
    .readCellValueAt(exactMatch.position - 1)
}
