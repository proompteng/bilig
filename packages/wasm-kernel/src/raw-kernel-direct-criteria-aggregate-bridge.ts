import type { RawKernelExports } from './raw-kernel-exports.js'
import { float64ArraySpec, RawKernelArrayBridge, uint16ArraySpec, uint32ArraySpec, uint8ArraySpec } from './raw-kernel-array-bridge.js'

export function evalDirectCriteriaMatchedAggregateBatchRaw(
  raw: RawKernelExports,
  aggregateKinds: Uint8Array,
  matchStarts: Uint32Array,
  matchLengths: Uint32Array,
  matchedRows: Uint32Array,
  aggregateTags: Uint8Array,
  aggregateNumbers: Float64Array,
  aggregateErrors: Uint16Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [
    aggregateKindsPtr,
    matchStartsPtr,
    matchLengthsPtr,
    matchedRowsPtr,
    aggregateTagsPtr,
    aggregateNumbersPtr,
    aggregateErrorsPtr,
    outTagsPtr,
    outNumbersPtr,
    outErrorsPtr,
  ] = arrayBridge.lowerTypedArrays([
    [aggregateKinds, uint8ArraySpec],
    [matchStarts, uint32ArraySpec],
    [matchLengths, uint32ArraySpec],
    [matchedRows, uint32ArraySpec],
    [aggregateTags, uint8ArraySpec],
    [aggregateNumbers, float64ArraySpec],
    [aggregateErrors, uint16ArraySpec],
    [outTags, uint8ArraySpec],
    [outNumbers, float64ArraySpec],
    [outErrors, uint16ArraySpec],
  ])
  try {
    raw.evalDirectCriteriaMatchedAggregateBatch(
      aggregateKindsPtr,
      matchStartsPtr,
      matchLengthsPtr,
      matchedRowsPtr,
      aggregateTagsPtr,
      aggregateNumbersPtr,
      aggregateErrorsPtr,
      outTagsPtr,
      outNumbersPtr,
      outErrorsPtr,
    )
    arrayBridge.copyTypedArray(outTagsPtr, outTags, uint8ArraySpec)
    arrayBridge.copyTypedArray(outNumbersPtr, outNumbers, float64ArraySpec)
    arrayBridge.copyTypedArray(outErrorsPtr, outErrors, uint16ArraySpec)
  } finally {
    raw.__unpin(aggregateKindsPtr)
    raw.__unpin(matchStartsPtr)
    raw.__unpin(matchLengthsPtr)
    raw.__unpin(matchedRowsPtr)
    raw.__unpin(aggregateTagsPtr)
    raw.__unpin(aggregateNumbersPtr)
    raw.__unpin(aggregateErrorsPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outErrorsPtr)
  }
}

export function evalDirectCriteriaPredicateAggregateBatchRaw(
  raw: RawKernelExports,
  aggregateKind: number,
  rowCount: number,
  criteriaOps: Uint8Array,
  criteriaKinds: Uint8Array,
  criteriaValues: Float64Array,
  criteriaStringIds: Uint32Array,
  criteriaTags: Uint8Array,
  criteriaNumbers: Float64Array,
  criteriaStringIdsByRow: Uint32Array,
  aggregateTags: Uint8Array,
  aggregateNumbers: Float64Array,
  aggregateErrors: Uint16Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [
    criteriaOpsPtr,
    criteriaKindsPtr,
    criteriaValuesPtr,
    criteriaStringIdsPtr,
    criteriaTagsPtr,
    criteriaNumbersPtr,
    criteriaStringIdsByRowPtr,
    aggregateTagsPtr,
    aggregateNumbersPtr,
    aggregateErrorsPtr,
    outTagsPtr,
    outNumbersPtr,
    outErrorsPtr,
  ] = arrayBridge.lowerTypedArrays([
    [criteriaOps, uint8ArraySpec],
    [criteriaKinds, uint8ArraySpec],
    [criteriaValues, float64ArraySpec],
    [criteriaStringIds, uint32ArraySpec],
    [criteriaTags, uint8ArraySpec],
    [criteriaNumbers, float64ArraySpec],
    [criteriaStringIdsByRow, uint32ArraySpec],
    [aggregateTags, uint8ArraySpec],
    [aggregateNumbers, float64ArraySpec],
    [aggregateErrors, uint16ArraySpec],
    [outTags, uint8ArraySpec],
    [outNumbers, float64ArraySpec],
    [outErrors, uint16ArraySpec],
  ])
  try {
    raw.evalDirectCriteriaPredicateAggregateBatch(
      aggregateKind,
      rowCount,
      criteriaOpsPtr,
      criteriaKindsPtr,
      criteriaValuesPtr,
      criteriaStringIdsPtr,
      criteriaTagsPtr,
      criteriaNumbersPtr,
      criteriaStringIdsByRowPtr,
      aggregateTagsPtr,
      aggregateNumbersPtr,
      aggregateErrorsPtr,
      outTagsPtr,
      outNumbersPtr,
      outErrorsPtr,
    )
    arrayBridge.copyTypedArray(outTagsPtr, outTags, uint8ArraySpec)
    arrayBridge.copyTypedArray(outNumbersPtr, outNumbers, float64ArraySpec)
    arrayBridge.copyTypedArray(outErrorsPtr, outErrors, uint16ArraySpec)
  } finally {
    raw.__unpin(criteriaOpsPtr)
    raw.__unpin(criteriaKindsPtr)
    raw.__unpin(criteriaValuesPtr)
    raw.__unpin(criteriaStringIdsPtr)
    raw.__unpin(criteriaTagsPtr)
    raw.__unpin(criteriaNumbersPtr)
    raw.__unpin(criteriaStringIdsByRowPtr)
    raw.__unpin(aggregateTagsPtr)
    raw.__unpin(aggregateNumbersPtr)
    raw.__unpin(aggregateErrorsPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outErrorsPtr)
  }
}
