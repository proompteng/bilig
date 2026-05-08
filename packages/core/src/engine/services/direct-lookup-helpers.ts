import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { exactLookupNumberKey, normalizeExactLookupNumber, sameExactLookupNumber as formulaSameExactLookupNumber } from '@bilig/formula'
import type { RuntimeDirectLookupDescriptor } from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'

export type UniformNumericDirectLookup = Extract<
  RuntimeDirectLookupDescriptor,
  { kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric' }
>

export interface DirectLookupSheetVersionState {
  readonly structureVersion?: number
  readonly columnVersions?: ArrayLike<number | undefined>
}

export function directLookupVersionMatches(
  lookupSheet: DirectLookupSheetVersionState | undefined,
  lookup: UniformNumericDirectLookup,
): boolean {
  if ((lookupSheet?.structureVersion ?? 0) !== lookup.structureVersion) {
    return false
  }
  const currentColumnVersion = lookupSheet?.columnVersions?.[lookup.col] ?? 0
  return currentColumnVersion === lookup.columnVersion || currentColumnVersion === lookup.tailPatch?.columnVersion
}

export function withOptionalLookupStringIds(request: {
  sheetName: string
  row: number
  col: number
  oldValue: CellValue
  newValue: CellValue
  oldStringId: number | undefined
  newStringId: number | undefined
  inputCellIndex?: number
}): {
  sheetName: string
  row: number
  col: number
  oldValue: CellValue
  newValue: CellValue
  oldStringId?: number
  newStringId?: number
  inputCellIndex?: number
} {
  return {
    sheetName: request.sheetName,
    row: request.row,
    col: request.col,
    oldValue: request.oldValue,
    newValue: request.newValue,
    ...(request.oldStringId === undefined ? {} : { oldStringId: request.oldStringId }),
    ...(request.newStringId === undefined ? {} : { newStringId: request.newStringId }),
    ...(request.inputCellIndex === undefined ? {} : { inputCellIndex: request.inputCellIndex }),
  }
}

export function normalizeExactLookupKey(value: CellValue, lookupString: (id: number) => string, stringId = 0): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'e:'
    case ValueTag.Number:
      return exactLookupNumberKey(value.value)
    case ValueTag.Boolean:
      return value.value ? 'b:1' : 'b:0'
    case ValueTag.String:
      return `s:${(stringId !== 0 ? lookupString(stringId) : value.value).toUpperCase()}`
    case ValueTag.Error:
      return undefined
  }
}

export function normalizeExactNumericValue(value: CellValue): number | undefined {
  return value.tag === ValueTag.Number ? normalizeExactLookupNumber(value.value) : undefined
}

export function sameExactNumericValue(left: number, right: number): boolean {
  return formulaSameExactLookupNumber(left, right)
}

export function normalizeApproximateNumericValue(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return 0
    case ValueTag.Number:
      return Object.is(value.value, -0) ? 0 : value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.String:
    case ValueTag.Error:
      return undefined
  }
}

export function normalizeApproximateTextValue(value: CellValue, lookupString: (id: number) => string, stringId = 0): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.String:
      return (stringId !== 0 ? lookupString(stringId) : value.value).toUpperCase()
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.Error:
      return undefined
  }
}

export function exactLookupLiteralNumericValue(value: unknown): number | undefined {
  return typeof value === 'number' ? normalizeExactLookupNumber(value) : undefined
}

export function canSkipUniformApproximateNumericTailWrite(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>,
  row: number,
  operandNumeric: number,
  oldNumeric: number,
  newNumeric: number,
): boolean {
  if (directLookup.repeatedRunLength !== undefined) {
    return false
  }
  if (directLookup.matchMode === 1 && directLookup.step > 0) {
    return row === directLookup.rowEnd && oldNumeric > operandNumeric && newNumeric > operandNumeric && newNumeric >= oldNumeric
  }
  if (directLookup.matchMode === -1 && directLookup.step < 0) {
    return row === directLookup.rowEnd && oldNumeric < operandNumeric && newNumeric < operandNumeric && newNumeric <= oldNumeric
  }
  return false
}

export function canSkipUniformApproximateNumericTailWriteFromCurrentResult(
  cellStore: {
    readonly tags: ArrayLike<ValueTag | undefined>
    readonly numbers: ArrayLike<number | undefined>
  },
  formulaCellIndex: number,
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>,
  row: number,
  oldNumeric: number,
  newNumeric: number,
): boolean {
  if (directLookup.repeatedRunLength !== undefined) {
    return false
  }
  if (
    !(
      (directLookup.matchMode === 1 && directLookup.step > 0 && row === directLookup.rowEnd && newNumeric >= oldNumeric) ||
      (directLookup.matchMode === -1 && directLookup.step < 0 && row === directLookup.rowEnd && newNumeric <= oldNumeric)
    )
  ) {
    return false
  }
  return (
    (cellStore.tags[formulaCellIndex] ?? ValueTag.Empty) === ValueTag.Number &&
    (cellStore.numbers[formulaCellIndex] ?? 0) < directLookup.length
  )
}

export function canSkipUniformExactNumericTailWriteFromCurrentResult(
  cellStore: {
    readonly tags: ArrayLike<ValueTag | undefined>
    readonly numbers: ArrayLike<number | undefined>
  },
  formulaCellIndex: number,
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>,
  row: number,
  oldNumeric: number,
  newNumeric: number,
): boolean {
  if (directLookup.tailPatch !== undefined || row !== directLookup.rowEnd || directLookup.length < 2) {
    return false
  }
  const expectedOldTail = directLookup.start + directLookup.step * (directLookup.length - 1)
  if (!sameExactNumericValue(oldNumeric, expectedOldTail)) {
    return false
  }
  if (!(directLookup.step > 0 ? newNumeric > oldNumeric : newNumeric < oldNumeric)) {
    return false
  }
  return (
    (cellStore.tags[formulaCellIndex] ?? ValueTag.Empty) === ValueTag.Number &&
    (cellStore.numbers[formulaCellIndex] ?? 0) < directLookup.length
  )
}

export function directLookupRowBounds(
  directLookup: Extract<
    RuntimeDirectLookupDescriptor,
    { kind: 'exact' | 'exact-uniform-numeric' | 'approximate' | 'approximate-uniform-numeric' }
  >,
): { rowStart: number; rowEnd: number } {
  switch (directLookup.kind) {
    case 'exact':
    case 'approximate':
      return {
        rowStart: directLookup.prepared.rowStart,
        rowEnd: directLookup.prepared.rowEnd,
      }
    case 'exact-uniform-numeric':
    case 'approximate-uniform-numeric':
      return {
        rowStart: directLookup.rowStart,
        rowEnd: directLookup.rowEnd,
      }
  }
}

export function exactUniformLookupCurrentResult(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>,
  lookupValue: number,
): DirectScalarCurrentOperand {
  const numericResult = exactUniformLookupNumericResult(directLookup, lookupValue)
  return numericResult === undefined ? { kind: 'error', code: ErrorCode.NA } : { kind: 'number', value: numericResult }
}

export function exactUniformLookupNumericResult(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>,
  lookupValue: number,
): number | undefined {
  const normalizedLookupValue = normalizeExactLookupNumber(lookupValue)
  const tailPatch = directLookup.tailPatch
  if (tailPatch === undefined && directLookup.step === 1) {
    if (!Number.isInteger(normalizedLookupValue)) {
      return undefined
    }
    const position = normalizedLookupValue - normalizeExactLookupNumber(directLookup.start) + 1
    return position >= 1 && position <= directLookup.length ? position : undefined
  }
  if (tailPatch === undefined && directLookup.step === -1) {
    if (!Number.isInteger(normalizedLookupValue)) {
      return undefined
    }
    const position = normalizeExactLookupNumber(directLookup.start) - normalizedLookupValue + 1
    return position >= 1 && position <= directLookup.length ? position : undefined
  }
  if (tailPatch !== undefined) {
    if (sameExactNumericValue(normalizedLookupValue, tailPatch.newNumeric)) {
      return tailPatch.row - directLookup.rowStart + 1
    }
    if (sameExactNumericValue(normalizedLookupValue, tailPatch.oldNumeric)) {
      return undefined
    }
  }
  if (directLookup.step === 1) {
    if (!Number.isInteger(normalizedLookupValue)) {
      return undefined
    }
    const position = normalizedLookupValue - normalizeExactLookupNumber(directLookup.start) + 1
    return position >= 1 && position <= directLookup.length ? position : undefined
  }
  if (directLookup.step === -1) {
    if (!Number.isInteger(normalizedLookupValue)) {
      return undefined
    }
    const position = normalizeExactLookupNumber(directLookup.start) - normalizedLookupValue + 1
    return position >= 1 && position <= directLookup.length ? position : undefined
  }
  const relative = (normalizedLookupValue - normalizeExactLookupNumber(directLookup.start)) / normalizeExactLookupNumber(directLookup.step)
  const nearestOffset = Math.round(relative)
  if (nearestOffset < 0 || nearestOffset >= directLookup.length) {
    return undefined
  }
  const candidate = directLookup.start + directLookup.step * nearestOffset
  return sameExactNumericValue(candidate, normalizedLookupValue) ? nearestOffset + 1 : undefined
}

export function approximateUniformLookupCurrentResult(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>,
  lookupValue: number,
): DirectScalarCurrentOperand | undefined {
  const numericResult = approximateUniformLookupNumericResult(directLookup, lookupValue)
  if (numericResult !== undefined) {
    return { kind: 'number', value: numericResult }
  }
  if (
    (directLookup.matchMode === 1 && directLookup.step > 0 && lookupValue < directLookup.start) ||
    (directLookup.matchMode === -1 && directLookup.step < 0 && lookupValue > directLookup.start)
  ) {
    return { kind: 'error', code: ErrorCode.NA }
  }
  return undefined
}

export function approximateUniformLookupNumericResult(
  directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>,
  lookupValue: number,
): number | undefined {
  if (directLookup.repeatedRunLength !== undefined) {
    const repeatedUniformResult = approximateRepeatedUniformLookupCurrentResult(
      {
        length: directLookup.length,
        repeatedUniformStart: directLookup.start,
        repeatedUniformStep: directLookup.step,
        repeatedUniformRunLength: directLookup.repeatedRunLength,
      },
      directLookup.matchMode,
      lookupValue,
    )
    return repeatedUniformResult?.kind === 'number' ? repeatedUniformResult.value : undefined
  }
  const tailPatch = directLookup.tailPatch
  if (tailPatch === undefined && directLookup.matchMode === 1 && directLookup.step === 1) {
    if (lookupValue < directLookup.start) {
      return undefined
    }
    const position = Math.floor(lookupValue - directLookup.start) + 1
    return position >= directLookup.length ? directLookup.length : position
  }
  if (tailPatch === undefined && directLookup.matchMode === -1 && directLookup.step === -1) {
    if (lookupValue > directLookup.start) {
      return undefined
    }
    const position = Math.floor(directLookup.start - lookupValue) + 1
    return position >= directLookup.length ? directLookup.length : position
  }
  const lastValue = directLookup.start + directLookup.step * (directLookup.length - 1)
  if (directLookup.matchMode === 1 && directLookup.step > 0) {
    if (lookupValue < directLookup.start) {
      return undefined
    }
    if (tailPatch !== undefined && tailPatch.row === directLookup.rowEnd && tailPatch.newNumeric > tailPatch.oldNumeric) {
      if (lookupValue >= tailPatch.newNumeric) {
        return directLookup.length
      }
      if (lookupValue >= tailPatch.oldNumeric) {
        return directLookup.length - 1
      }
    }
    if (lookupValue >= lastValue) {
      return directLookup.length
    }
    const position =
      directLookup.step === 1
        ? Math.floor(lookupValue - directLookup.start) + 1
        : Math.floor((lookupValue - directLookup.start) / directLookup.step) + 1
    return Math.min(directLookup.length, Math.max(1, position))
  }
  if (directLookup.matchMode === -1 && directLookup.step < 0) {
    if (lookupValue > directLookup.start) {
      return undefined
    }
    if (tailPatch !== undefined && tailPatch.row === directLookup.rowEnd && tailPatch.newNumeric < tailPatch.oldNumeric) {
      if (lookupValue <= tailPatch.newNumeric) {
        return directLookup.length
      }
      if (lookupValue <= tailPatch.oldNumeric) {
        return directLookup.length - 1
      }
    }
    if (lookupValue <= lastValue) {
      return directLookup.length
    }
    const position =
      directLookup.step === -1
        ? Math.floor(directLookup.start - lookupValue) + 1
        : Math.floor((directLookup.start - lookupValue) / -directLookup.step) + 1
    return Math.min(directLookup.length, Math.max(1, position))
  }
  return undefined
}

export function approximateRepeatedUniformLookupCurrentResult(
  prepared: {
    readonly length: number
    readonly repeatedUniformStart: number | undefined
    readonly repeatedUniformStep: number | undefined
    readonly repeatedUniformRunLength: number | undefined
  },
  matchMode: 1 | -1,
  lookupValue: number,
): DirectScalarCurrentOperand | undefined {
  const { repeatedUniformStart, repeatedUniformStep, repeatedUniformRunLength, length } = prepared
  if (repeatedUniformStart === undefined || repeatedUniformStep === undefined || repeatedUniformRunLength === undefined || length <= 0) {
    return undefined
  }
  const groupCount = Math.ceil(length / repeatedUniformRunLength)
  const lastValue = repeatedUniformStart + repeatedUniformStep * (groupCount - 1)
  if (matchMode === 1 && repeatedUniformStep > 0) {
    if (lookupValue < repeatedUniformStart) {
      return { kind: 'error', code: ErrorCode.NA }
    }
    if (lookupValue >= lastValue) {
      return { kind: 'number', value: length }
    }
    const group = Math.floor((lookupValue - repeatedUniformStart) / repeatedUniformStep)
    return { kind: 'number', value: Math.min(length, (group + 1) * repeatedUniformRunLength) }
  }
  if (matchMode === -1 && repeatedUniformStep < 0) {
    if (lookupValue > repeatedUniformStart) {
      return { kind: 'error', code: ErrorCode.NA }
    }
    if (lookupValue <= lastValue) {
      return { kind: 'number', value: length }
    }
    const group = Math.floor((repeatedUniformStart - lookupValue) / -repeatedUniformStep)
    return { kind: 'number', value: Math.min(length, (group + 1) * repeatedUniformRunLength) }
  }
  return undefined
}
