import { parseCellAddress } from '@bilig/formula'
import type { ErrorCode } from '@bilig/protocol'

type DirectLookupOperandInstruction =
  | { opcode: 'push-cell'; address: string; sheetName?: string }
  | { opcode: 'push-number'; value: number }
  | { opcode: 'push-boolean'; value: boolean }
  | { opcode: 'push-string'; value: string }
  | { opcode: 'push-error'; code: ErrorCode }
  | { opcode: 'push-name'; name: string }

type DirectExactLookupInstruction = {
  opcode: 'lookup-exact-match'
  sheetName?: string
  start: string
  end: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
  searchMode: 1 | -1
}

type DirectApproximateLookupInstruction = {
  opcode: 'lookup-approximate-match'
  sheetName?: string
  start: string
  end: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
  matchMode: 1 | -1
}

type DirectVectorLookupInstruction = DirectExactLookupInstruction | DirectApproximateLookupInstruction

export type RuntimeDirectLookupBinding =
  | {
      kind: 'exact'
      operandSheetName: string
      operandAddress: string
      lookupSheetName: string
      rowStart: number
      rowEnd: number
      col: number
      searchMode: 1 | -1
    }
  | {
      kind: 'approximate'
      operandSheetName: string
      operandAddress: string
      lookupSheetName: string
      rowStart: number
      rowEnd: number
      col: number
      matchMode: 1 | -1
    }

function isDirectLookupOperandInstruction(value: unknown): value is DirectLookupOperandInstruction {
  if (!value || typeof value !== 'object') {
    return false
  }
  const opcode = Reflect.get(value, 'opcode')
  switch (opcode) {
    case 'push-cell':
      return typeof Reflect.get(value, 'address') === 'string'
    case 'push-number':
      return typeof Reflect.get(value, 'value') === 'number'
    case 'push-boolean':
      return typeof Reflect.get(value, 'value') === 'boolean'
    case 'push-string':
      return typeof Reflect.get(value, 'value') === 'string'
    case 'push-error':
      return typeof Reflect.get(value, 'code') === 'number'
    case 'push-name':
      return typeof Reflect.get(value, 'name') === 'string'
    default:
      return false
  }
}

function isDirectVectorLookupInstruction(value: unknown): value is DirectVectorLookupInstruction {
  if (!value || typeof value !== 'object') {
    return false
  }
  const opcode = Reflect.get(value, 'opcode')
  if (
    (opcode !== 'lookup-exact-match' && opcode !== 'lookup-approximate-match') ||
    typeof Reflect.get(value, 'start') !== 'string' ||
    typeof Reflect.get(value, 'end') !== 'string' ||
    typeof Reflect.get(value, 'startRow') !== 'number' ||
    typeof Reflect.get(value, 'endRow') !== 'number' ||
    typeof Reflect.get(value, 'startCol') !== 'number' ||
    typeof Reflect.get(value, 'endCol') !== 'number'
  ) {
    return false
  }
  if (opcode === 'lookup-exact-match') {
    const searchMode = Reflect.get(value, 'searchMode')
    return searchMode === 1 || searchMode === -1
  }
  const matchMode = Reflect.get(value, 'matchMode')
  return matchMode === 1 || matchMode === -1
}

function isReturnInstruction(value: unknown): value is { opcode: 'return' } {
  return !!value && typeof value === 'object' && Reflect.get(value, 'opcode') === 'return'
}

export function resolveRuntimeDirectLookupBinding(
  jsPlan: readonly unknown[],
  ownerSheetName: string,
): RuntimeDirectLookupBinding | undefined {
  const [operand, lookup, terminal] = jsPlan
  if (!operand || !lookup || !terminal || !isReturnInstruction(terminal)) {
    return undefined
  }
  if (!isDirectLookupOperandInstruction(operand) || !isDirectVectorLookupInstruction(lookup)) {
    return undefined
  }
  if (operand.opcode !== 'push-cell' || lookup.startCol !== lookup.endCol) {
    return undefined
  }

  const operandSheetName = operand.sheetName ?? ownerSheetName
  const operandAddress = parseCellAddress(operand.address, operandSheetName).text
  const lookupSheetName = lookup.sheetName ?? ownerSheetName
  if (lookup.opcode === 'lookup-exact-match') {
    return {
      kind: 'exact',
      operandSheetName,
      operandAddress,
      lookupSheetName,
      rowStart: lookup.startRow,
      rowEnd: lookup.endRow,
      col: lookup.startCol,
      searchMode: lookup.searchMode,
    }
  }

  return {
    kind: 'approximate',
    operandSheetName,
    operandAddress,
    lookupSheetName,
    rowStart: lookup.startRow,
    rowEnd: lookup.endRow,
    col: lookup.startCol,
    matchMode: lookup.matchMode,
  }
}
