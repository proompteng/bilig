import type { RawKernelExports } from './raw-kernel-exports.js'

type TypedArrayValue = Uint8Array | Uint16Array | Uint32Array | Float64Array

const ARRAY_BUFFER_CLASS_ID = 1
const UINT8_ARRAY_CLASS_ID = 4
const FLOAT64_ARRAY_CLASS_ID = 5
const UINT16_ARRAY_CLASS_ID = 6
const UINT32_ARRAY_CLASS_ID = 7

interface LoweredArraySpec<T extends TypedArrayValue> {
  align: number
  classId: number
  ctor: {
    new (buffer: ArrayBufferLike, byteOffset: number, length: number): T
  }
}

const uint8Spec: LoweredArraySpec<Uint8Array> = {
  align: 0,
  classId: UINT8_ARRAY_CLASS_ID,
  ctor: Uint8Array,
}

const uint16Spec: LoweredArraySpec<Uint16Array> = {
  align: 1,
  classId: UINT16_ARRAY_CLASS_ID,
  ctor: Uint16Array,
}

const uint32Spec: LoweredArraySpec<Uint32Array> = {
  align: 2,
  classId: UINT32_ARRAY_CLASS_ID,
  ctor: Uint32Array,
}

const float64Spec: LoweredArraySpec<Float64Array> = {
  align: 3,
  classId: FLOAT64_ARRAY_CLASS_ID,
  ctor: Float64Array,
}

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
  let dataView = new DataView(raw.memory.buffer)
  const setUint32 = (pointer: number, value: number): void => {
    try {
      dataView.setUint32(pointer, value, true)
    } catch {
      dataView = new DataView(raw.memory.buffer)
      dataView.setUint32(pointer, value, true)
    }
  }
  const getUint32 = (pointer: number): number => {
    try {
      return dataView.getUint32(pointer, true)
    } catch {
      dataView = new DataView(raw.memory.buffer)
      return dataView.getUint32(pointer, true)
    }
  }
  const lowerTypedArray = <T extends TypedArrayValue>(values: T, spec: LoweredArraySpec<T>): number => {
    const byteLength = values.length << spec.align
    const bufferPtr = raw.__pin(raw.__new(byteLength, ARRAY_BUFFER_CLASS_ID))
    const headerPtr = raw.__pin(raw.__new(12, spec.classId))
    try {
      setUint32(headerPtr, bufferPtr)
      setUint32(headerPtr + 4, bufferPtr)
      setUint32(headerPtr + 8, byteLength)
      new spec.ctor(raw.memory.buffer, bufferPtr, values.length).set(values)
      return headerPtr
    } finally {
      raw.__unpin(bufferPtr)
    }
  }
  const copyLoweredTypedArray = <T extends TypedArrayValue>(pointer: number, target: T, spec: LoweredArraySpec<T>): void => {
    target.set(new spec.ctor(raw.memory.buffer, getUint32(pointer + 4), target.length))
  }

  const aggregateKindsPtr = lowerTypedArray(aggregateKinds, uint8Spec)
  const matchStartsPtr = lowerTypedArray(matchStarts, uint32Spec)
  const matchLengthsPtr = lowerTypedArray(matchLengths, uint32Spec)
  const matchedRowsPtr = lowerTypedArray(matchedRows, uint32Spec)
  const aggregateTagsPtr = lowerTypedArray(aggregateTags, uint8Spec)
  const aggregateNumbersPtr = lowerTypedArray(aggregateNumbers, float64Spec)
  const aggregateErrorsPtr = lowerTypedArray(aggregateErrors, uint16Spec)
  const outTagsPtr = lowerTypedArray(outTags, uint8Spec)
  const outNumbersPtr = lowerTypedArray(outNumbers, float64Spec)
  const outErrorsPtr = lowerTypedArray(outErrors, uint16Spec)
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
    copyLoweredTypedArray(outTagsPtr, outTags, uint8Spec)
    copyLoweredTypedArray(outNumbersPtr, outNumbers, float64Spec)
    copyLoweredTypedArray(outErrorsPtr, outErrors, uint16Spec)
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
  let dataView = new DataView(raw.memory.buffer)
  const setUint32 = (pointer: number, value: number): void => {
    try {
      dataView.setUint32(pointer, value, true)
    } catch {
      dataView = new DataView(raw.memory.buffer)
      dataView.setUint32(pointer, value, true)
    }
  }
  const getUint32 = (pointer: number): number => {
    try {
      return dataView.getUint32(pointer, true)
    } catch {
      dataView = new DataView(raw.memory.buffer)
      return dataView.getUint32(pointer, true)
    }
  }
  const lowerTypedArray = <T extends TypedArrayValue>(values: T, spec: LoweredArraySpec<T>): number => {
    const byteLength = values.length << spec.align
    const bufferPtr = raw.__pin(raw.__new(byteLength, ARRAY_BUFFER_CLASS_ID))
    const headerPtr = raw.__pin(raw.__new(12, spec.classId))
    try {
      setUint32(headerPtr, bufferPtr)
      setUint32(headerPtr + 4, bufferPtr)
      setUint32(headerPtr + 8, byteLength)
      new spec.ctor(raw.memory.buffer, bufferPtr, values.length).set(values)
      return headerPtr
    } finally {
      raw.__unpin(bufferPtr)
    }
  }
  const copyLoweredTypedArray = <T extends TypedArrayValue>(pointer: number, target: T, spec: LoweredArraySpec<T>): void => {
    target.set(new spec.ctor(raw.memory.buffer, getUint32(pointer + 4), target.length))
  }

  const criteriaOpsPtr = lowerTypedArray(criteriaOps, uint8Spec)
  const criteriaKindsPtr = lowerTypedArray(criteriaKinds, uint8Spec)
  const criteriaValuesPtr = lowerTypedArray(criteriaValues, float64Spec)
  const criteriaStringIdsPtr = lowerTypedArray(criteriaStringIds, uint32Spec)
  const criteriaTagsPtr = lowerTypedArray(criteriaTags, uint8Spec)
  const criteriaNumbersPtr = lowerTypedArray(criteriaNumbers, float64Spec)
  const criteriaStringIdsByRowPtr = lowerTypedArray(criteriaStringIdsByRow, uint32Spec)
  const aggregateTagsPtr = lowerTypedArray(aggregateTags, uint8Spec)
  const aggregateNumbersPtr = lowerTypedArray(aggregateNumbers, float64Spec)
  const aggregateErrorsPtr = lowerTypedArray(aggregateErrors, uint16Spec)
  const outTagsPtr = lowerTypedArray(outTags, uint8Spec)
  const outNumbersPtr = lowerTypedArray(outNumbers, float64Spec)
  const outErrorsPtr = lowerTypedArray(outErrors, uint16Spec)
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
    copyLoweredTypedArray(outTagsPtr, outTags, uint8Spec)
    copyLoweredTypedArray(outNumbersPtr, outNumbers, float64Spec)
    copyLoweredTypedArray(outErrorsPtr, outErrors, uint16Spec)
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
