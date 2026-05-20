import { formatAddress } from '@bilig/formula'
import type { Rectangle } from './gridTypes.js'
import { serializeClipboardMatrix, serializeClipboardPlainText } from './gridClipboard.js'

export interface InternalClipboardRange {
  operation: 'copy' | 'cut'
  sourceStartAddress: string
  sourceEndAddress: string
  signature: string
  plainText: string
  valuesOnlyPlainText: string
  rowCount: number
  colCount: number
}

export function buildInternalClipboardRange(
  range: Rectangle,
  values: readonly (readonly string[])[],
  operation: InternalClipboardRange['operation'] = 'copy',
  valuesOnlyValues: readonly (readonly string[])[] = values,
): InternalClipboardRange {
  return {
    operation,
    sourceStartAddress: formatAddress(range.y, range.x),
    sourceEndAddress: formatAddress(range.y + range.height - 1, range.x + range.width - 1),
    signature: serializeClipboardMatrix(values),
    plainText: serializeClipboardPlainText(values),
    valuesOnlyPlainText: serializeClipboardPlainText(valuesOnlyValues),
    rowCount: range.height,
    colCount: range.width,
  }
}

export function matchesInternalClipboardPaste(
  internalClipboard: InternalClipboardRange | null,
  values: readonly (readonly string[])[],
): boolean {
  if (!internalClipboard || values.length === 0 || values[0]?.length === 0) {
    return false
  }
  if (
    internalClipboard.signature === serializeClipboardMatrix(values) &&
    internalClipboard.rowCount === values.length &&
    internalClipboard.colCount === (values[0]?.length ?? 0)
  ) {
    return true
  }
  if (values.length > internalClipboard.rowCount || values.some((row) => row.length !== internalClipboard.colCount)) {
    return false
  }
  const internalValues = deserializeClipboardSignature(internalClipboard.signature)
  if (internalValues.length !== internalClipboard.rowCount || internalValues.some((row) => row.length !== internalClipboard.colCount)) {
    return false
  }
  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex] ?? []
    const internalRow = internalValues[rowIndex] ?? []
    for (let colIndex = 0; colIndex < internalClipboard.colCount; colIndex += 1) {
      if ((row[colIndex] ?? '') !== (internalRow[colIndex] ?? '')) {
        return false
      }
    }
  }
  for (let rowIndex = values.length; rowIndex < internalValues.length; rowIndex += 1) {
    if (!internalValues[rowIndex]?.every((value) => value.length === 0)) {
      return false
    }
  }
  return true
}

function deserializeClipboardSignature(signature: string): readonly (readonly string[])[] {
  return signature.split('\u001e').map((row) => row.split('\u001f'))
}
