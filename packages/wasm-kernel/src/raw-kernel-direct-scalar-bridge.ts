import type { RawKernelExports } from './raw-kernel-exports.js'
import { float64ArraySpec, RawKernelArrayBridge, uint16ArraySpec, uint32ArraySpec, uint8ArraySpec } from './raw-kernel-array-bridge.js'

export function evalDirectScalarValueBatchRaw(
  raw: RawKernelExports,
  operators: Uint8Array,
  leftBatchRefs: Uint32Array,
  leftTags: Uint8Array,
  leftValues: Float64Array,
  leftErrors: Uint16Array,
  rightBatchRefs: Uint32Array,
  rightTags: Uint8Array,
  rightValues: Float64Array,
  rightErrors: Uint16Array,
  resultOffsets: Float64Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [
    operatorsPtr,
    leftBatchRefsPtr,
    leftTagsPtr,
    leftValuesPtr,
    leftErrorsPtr,
    rightBatchRefsPtr,
    rightTagsPtr,
    rightValuesPtr,
    rightErrorsPtr,
    resultOffsetsPtr,
    outTagsPtr,
    outNumbersPtr,
    outErrorsPtr,
  ] = arrayBridge.lowerTypedArrays([
    [operators, uint8ArraySpec],
    [leftBatchRefs, uint32ArraySpec],
    [leftTags, uint8ArraySpec],
    [leftValues, float64ArraySpec],
    [leftErrors, uint16ArraySpec],
    [rightBatchRefs, uint32ArraySpec],
    [rightTags, uint8ArraySpec],
    [rightValues, float64ArraySpec],
    [rightErrors, uint16ArraySpec],
    [resultOffsets, float64ArraySpec],
    [outTags, uint8ArraySpec],
    [outNumbers, float64ArraySpec],
    [outErrors, uint16ArraySpec],
  ])
  try {
    raw.evalDirectScalarValueBatch(
      operatorsPtr,
      leftBatchRefsPtr,
      leftTagsPtr,
      leftValuesPtr,
      leftErrorsPtr,
      rightBatchRefsPtr,
      rightTagsPtr,
      rightValuesPtr,
      rightErrorsPtr,
      resultOffsetsPtr,
      outTagsPtr,
      outNumbersPtr,
      outErrorsPtr,
    )
    arrayBridge.copyTypedArray(outTagsPtr, outTags, uint8ArraySpec)
    arrayBridge.copyTypedArray(outNumbersPtr, outNumbers, float64ArraySpec)
    arrayBridge.copyTypedArray(outErrorsPtr, outErrors, uint16ArraySpec)
  } finally {
    raw.__unpin(operatorsPtr)
    raw.__unpin(leftBatchRefsPtr)
    raw.__unpin(leftTagsPtr)
    raw.__unpin(leftValuesPtr)
    raw.__unpin(leftErrorsPtr)
    raw.__unpin(rightBatchRefsPtr)
    raw.__unpin(rightTagsPtr)
    raw.__unpin(rightValuesPtr)
    raw.__unpin(rightErrorsPtr)
    raw.__unpin(resultOffsetsPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outErrorsPtr)
  }
}

export function evalDirectScalarStoreTargetBatchRaw(
  raw: RawKernelExports,
  targets: Uint32Array,
  operators: Uint8Array,
  leftBatchRefs: Uint32Array,
  leftTags: Uint8Array,
  leftValues: Float64Array,
  leftErrors: Uint16Array,
  rightBatchRefs: Uint32Array,
  rightTags: Uint8Array,
  rightValues: Float64Array,
  rightErrors: Uint16Array,
  resultOffsets: Float64Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [
    targetsPtr,
    operatorsPtr,
    leftBatchRefsPtr,
    leftTagsPtr,
    leftValuesPtr,
    leftErrorsPtr,
    rightBatchRefsPtr,
    rightTagsPtr,
    rightValuesPtr,
    rightErrorsPtr,
    resultOffsetsPtr,
  ] = arrayBridge.lowerTypedArrays([
    [targets, uint32ArraySpec],
    [operators, uint8ArraySpec],
    [leftBatchRefs, uint32ArraySpec],
    [leftTags, uint8ArraySpec],
    [leftValues, float64ArraySpec],
    [leftErrors, uint16ArraySpec],
    [rightBatchRefs, uint32ArraySpec],
    [rightTags, uint8ArraySpec],
    [rightValues, float64ArraySpec],
    [rightErrors, uint16ArraySpec],
    [resultOffsets, float64ArraySpec],
  ])
  try {
    raw.evalDirectScalarStoreTargetBatch(
      targetsPtr,
      operatorsPtr,
      leftBatchRefsPtr,
      leftTagsPtr,
      leftValuesPtr,
      leftErrorsPtr,
      rightBatchRefsPtr,
      rightTagsPtr,
      rightValuesPtr,
      rightErrorsPtr,
      resultOffsetsPtr,
    )
  } finally {
    raw.__unpin(targetsPtr)
    raw.__unpin(operatorsPtr)
    raw.__unpin(leftBatchRefsPtr)
    raw.__unpin(leftTagsPtr)
    raw.__unpin(leftValuesPtr)
    raw.__unpin(leftErrorsPtr)
    raw.__unpin(rightBatchRefsPtr)
    raw.__unpin(rightTagsPtr)
    raw.__unpin(rightValuesPtr)
    raw.__unpin(rightErrorsPtr)
    raw.__unpin(resultOffsetsPtr)
  }
}

export function evalDirectConditionalPickBatchRaw(
  raw: RawKernelExports,
  conditionStarts: Uint32Array,
  conditionLengths: Uint32Array,
  conditionOps: Uint8Array,
  leftTags: Uint8Array,
  leftNumbers: Float64Array,
  leftStringIds: Uint32Array,
  leftErrors: Uint16Array,
  rightTags: Uint8Array,
  rightNumbers: Float64Array,
  rightStringIds: Uint32Array,
  rightErrors: Uint16Array,
  branchTags: Uint8Array,
  branchNumbers: Float64Array,
  branchStringIds: Uint32Array,
  branchErrors: Uint16Array,
  defaultTags: Uint8Array,
  defaultNumbers: Float64Array,
  defaultStringIds: Uint32Array,
  defaultErrors: Uint16Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outStringIds: Uint32Array,
  outErrors: Uint16Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [
    conditionStartsPtr,
    conditionLengthsPtr,
    conditionOpsPtr,
    leftTagsPtr,
    leftNumbersPtr,
    leftStringIdsPtr,
    leftErrorsPtr,
    rightTagsPtr,
    rightNumbersPtr,
    rightStringIdsPtr,
    rightErrorsPtr,
    branchTagsPtr,
    branchNumbersPtr,
    branchStringIdsPtr,
    branchErrorsPtr,
    defaultTagsPtr,
    defaultNumbersPtr,
    defaultStringIdsPtr,
    defaultErrorsPtr,
    outTagsPtr,
    outNumbersPtr,
    outStringIdsPtr,
    outErrorsPtr,
  ] = arrayBridge.lowerTypedArrays([
    [conditionStarts, uint32ArraySpec],
    [conditionLengths, uint32ArraySpec],
    [conditionOps, uint8ArraySpec],
    [leftTags, uint8ArraySpec],
    [leftNumbers, float64ArraySpec],
    [leftStringIds, uint32ArraySpec],
    [leftErrors, uint16ArraySpec],
    [rightTags, uint8ArraySpec],
    [rightNumbers, float64ArraySpec],
    [rightStringIds, uint32ArraySpec],
    [rightErrors, uint16ArraySpec],
    [branchTags, uint8ArraySpec],
    [branchNumbers, float64ArraySpec],
    [branchStringIds, uint32ArraySpec],
    [branchErrors, uint16ArraySpec],
    [defaultTags, uint8ArraySpec],
    [defaultNumbers, float64ArraySpec],
    [defaultStringIds, uint32ArraySpec],
    [defaultErrors, uint16ArraySpec],
    [outTags, uint8ArraySpec],
    [outNumbers, float64ArraySpec],
    [outStringIds, uint32ArraySpec],
    [outErrors, uint16ArraySpec],
  ])
  try {
    raw.evalDirectConditionalPickBatch(
      conditionStartsPtr,
      conditionLengthsPtr,
      conditionOpsPtr,
      leftTagsPtr,
      leftNumbersPtr,
      leftStringIdsPtr,
      leftErrorsPtr,
      rightTagsPtr,
      rightNumbersPtr,
      rightStringIdsPtr,
      rightErrorsPtr,
      branchTagsPtr,
      branchNumbersPtr,
      branchStringIdsPtr,
      branchErrorsPtr,
      defaultTagsPtr,
      defaultNumbersPtr,
      defaultStringIdsPtr,
      defaultErrorsPtr,
      outTagsPtr,
      outNumbersPtr,
      outStringIdsPtr,
      outErrorsPtr,
    )
    arrayBridge.copyTypedArray(outTagsPtr, outTags, uint8ArraySpec)
    arrayBridge.copyTypedArray(outNumbersPtr, outNumbers, float64ArraySpec)
    arrayBridge.copyTypedArray(outStringIdsPtr, outStringIds, uint32ArraySpec)
    arrayBridge.copyTypedArray(outErrorsPtr, outErrors, uint16ArraySpec)
  } finally {
    raw.__unpin(conditionStartsPtr)
    raw.__unpin(conditionLengthsPtr)
    raw.__unpin(conditionOpsPtr)
    raw.__unpin(leftTagsPtr)
    raw.__unpin(leftNumbersPtr)
    raw.__unpin(leftStringIdsPtr)
    raw.__unpin(leftErrorsPtr)
    raw.__unpin(rightTagsPtr)
    raw.__unpin(rightNumbersPtr)
    raw.__unpin(rightStringIdsPtr)
    raw.__unpin(rightErrorsPtr)
    raw.__unpin(branchTagsPtr)
    raw.__unpin(branchNumbersPtr)
    raw.__unpin(branchStringIdsPtr)
    raw.__unpin(branchErrorsPtr)
    raw.__unpin(defaultTagsPtr)
    raw.__unpin(defaultNumbersPtr)
    raw.__unpin(defaultStringIdsPtr)
    raw.__unpin(defaultErrorsPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outStringIdsPtr)
    raw.__unpin(outErrorsPtr)
  }
}

export function evalDenseDirectScalarRowChainStoreTargetBatchRaw(
  raw: RawKernelExports,
  leftValues: Float64Array,
  rightValues: Float64Array,
  firstTargets: Uint32Array,
  secondTargets: Uint32Array,
  rowCount: number,
  firstFormulaCode: number,
  secondFormulaScale: number,
  secondFormulaOffset: number,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [leftValuesPtr, rightValuesPtr, firstTargetsPtr, secondTargetsPtr] = arrayBridge.lowerTypedArrays([
    [leftValues, float64ArraySpec],
    [rightValues, float64ArraySpec],
    [firstTargets, uint32ArraySpec],
    [secondTargets, uint32ArraySpec],
  ])
  try {
    raw.evalDenseDirectScalarRowChainStoreTargetBatch(
      leftValuesPtr,
      rightValuesPtr,
      firstTargetsPtr,
      secondTargetsPtr,
      rowCount,
      firstFormulaCode,
      secondFormulaScale,
      secondFormulaOffset,
    )
  } finally {
    raw.__unpin(leftValuesPtr)
    raw.__unpin(rightValuesPtr)
    raw.__unpin(firstTargetsPtr)
    raw.__unpin(secondTargetsPtr)
  }
}

export function evalDenseDirectScalarRowChainDivideStoreTargetBatchRaw(
  raw: RawKernelExports,
  leftValues: Float64Array,
  rightValues: Float64Array,
  denominatorValues: Float64Array,
  firstTargets: Uint32Array,
  secondTargets: Uint32Array,
  rowCount: number,
  firstFormulaCode: number,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [leftValuesPtr, rightValuesPtr, denominatorValuesPtr, firstTargetsPtr, secondTargetsPtr] = arrayBridge.lowerTypedArrays([
    [leftValues, float64ArraySpec],
    [rightValues, float64ArraySpec],
    [denominatorValues, float64ArraySpec],
    [firstTargets, uint32ArraySpec],
    [secondTargets, uint32ArraySpec],
  ])
  try {
    raw.evalDenseDirectScalarRowChainDivideStoreTargetBatch(
      leftValuesPtr,
      rightValuesPtr,
      denominatorValuesPtr,
      firstTargetsPtr,
      secondTargetsPtr,
      rowCount,
      firstFormulaCode,
    )
  } finally {
    raw.__unpin(leftValuesPtr)
    raw.__unpin(rightValuesPtr)
    raw.__unpin(denominatorValuesPtr)
    raw.__unpin(firstTargetsPtr)
    raw.__unpin(secondTargetsPtr)
  }
}
