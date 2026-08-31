import type { RawKernelExports } from './raw-kernel-exports.js'
import { float64ArraySpec, RawKernelArrayBridge, uint16ArraySpec, uint32ArraySpec, uint8ArraySpec } from './raw-kernel-array-bridge.js'

export function evalDenseNumericRowAggregateBatchRaw(
  raw: RawKernelExports,
  aggregateKind: number,
  values: Float64Array,
  rowCount: number,
  prefixColCount: number,
  startColOffset: number,
  aggregateColCount: number,
  resultOffset: number,
  outNumbers: Float64Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [valuesPtr, outNumbersPtr] = arrayBridge.lowerTypedArrays([
    [values, float64ArraySpec],
    [outNumbers, float64ArraySpec],
  ])
  try {
    raw.evalDenseNumericRowAggregateBatch(
      aggregateKind,
      valuesPtr,
      rowCount,
      prefixColCount,
      startColOffset,
      aggregateColCount,
      resultOffset,
      outNumbersPtr,
    )
    arrayBridge.copyTypedArray(outNumbersPtr, outNumbers, float64ArraySpec)
  } finally {
    raw.__unpin(valuesPtr)
    raw.__unpin(outNumbersPtr)
  }
}

export function evalAnchoredPrefixAggregateBatchRaw(
  raw: RawKernelExports,
  aggregateKind: number,
  tags: Uint8Array,
  numbers: Float64Array,
  errors: Uint16Array,
  rowCount: number,
  colCount: number,
  formulaRowEnds: Uint32Array,
  resultOffsets: Float64Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [tagsPtr, numbersPtr, errorsPtr, formulaRowEndsPtr, resultOffsetsPtr, outTagsPtr, outNumbersPtr, outErrorsPtr] =
    arrayBridge.lowerTypedArrays([
      [tags, uint8ArraySpec],
      [numbers, float64ArraySpec],
      [errors, uint16ArraySpec],
      [formulaRowEnds, uint32ArraySpec],
      [resultOffsets, float64ArraySpec],
      [outTags, uint8ArraySpec],
      [outNumbers, float64ArraySpec],
      [outErrors, uint16ArraySpec],
    ])
  try {
    raw.evalAnchoredPrefixAggregateBatch(
      aggregateKind,
      tagsPtr,
      numbersPtr,
      errorsPtr,
      rowCount,
      colCount,
      formulaRowEndsPtr,
      resultOffsetsPtr,
      outTagsPtr,
      outNumbersPtr,
      outErrorsPtr,
    )
    arrayBridge.copyTypedArray(outTagsPtr, outTags, uint8ArraySpec)
    arrayBridge.copyTypedArray(outNumbersPtr, outNumbers, float64ArraySpec)
    arrayBridge.copyTypedArray(outErrorsPtr, outErrors, uint16ArraySpec)
  } finally {
    raw.__unpin(tagsPtr)
    raw.__unpin(numbersPtr)
    raw.__unpin(errorsPtr)
    raw.__unpin(formulaRowEndsPtr)
    raw.__unpin(resultOffsetsPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outErrorsPtr)
  }
}

export function evalDenseCellRangeAggregateBatchRaw(
  raw: RawKernelExports,
  aggregateKind: number,
  tags: Uint8Array,
  numbers: Float64Array,
  errors: Uint16Array,
  rowCount: number,
  cellCount: number,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [tagsPtr, numbersPtr, errorsPtr, outTagsPtr, outNumbersPtr, outErrorsPtr] = arrayBridge.lowerTypedArrays([
    [tags, uint8ArraySpec],
    [numbers, float64ArraySpec],
    [errors, uint16ArraySpec],
    [outTags, uint8ArraySpec],
    [outNumbers, float64ArraySpec],
    [outErrors, uint16ArraySpec],
  ])
  try {
    raw.evalDenseCellRangeAggregateBatch(
      aggregateKind,
      tagsPtr,
      numbersPtr,
      errorsPtr,
      rowCount,
      cellCount,
      outTagsPtr,
      outNumbersPtr,
      outErrorsPtr,
    )
    arrayBridge.copyTypedArray(outTagsPtr, outTags, uint8ArraySpec)
    arrayBridge.copyTypedArray(outNumbersPtr, outNumbers, float64ArraySpec)
    arrayBridge.copyTypedArray(outErrorsPtr, outErrors, uint16ArraySpec)
  } finally {
    raw.__unpin(tagsPtr)
    raw.__unpin(numbersPtr)
    raw.__unpin(errorsPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outErrorsPtr)
  }
}
