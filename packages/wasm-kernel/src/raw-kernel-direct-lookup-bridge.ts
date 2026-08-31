import type { RawKernelExports } from './raw-kernel-exports.js'
import { float64ArraySpec, RawKernelArrayBridge, uint16ArraySpec, uint32ArraySpec, uint8ArraySpec } from './raw-kernel-array-bridge.js'

export function evalUniformNumericLookupBatchRaw(
  raw: RawKernelExports,
  kinds: Uint8Array,
  matchModes: Uint8Array,
  starts: Float64Array,
  steps: Float64Array,
  lengths: Uint32Array,
  repeatedRunLengths: Uint32Array,
  lookupTags: Uint8Array,
  lookupNumbers: Float64Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [
    kindsPtr,
    matchModesPtr,
    startsPtr,
    stepsPtr,
    lengthsPtr,
    repeatedRunLengthsPtr,
    lookupTagsPtr,
    lookupNumbersPtr,
    outTagsPtr,
    outNumbersPtr,
    outErrorsPtr,
  ] = arrayBridge.lowerTypedArrays([
    [kinds, uint8ArraySpec],
    [matchModes, uint8ArraySpec],
    [starts, float64ArraySpec],
    [steps, float64ArraySpec],
    [lengths, uint32ArraySpec],
    [repeatedRunLengths, uint32ArraySpec],
    [lookupTags, uint8ArraySpec],
    [lookupNumbers, float64ArraySpec],
    [outTags, uint8ArraySpec],
    [outNumbers, float64ArraySpec],
    [outErrors, uint16ArraySpec],
  ])
  try {
    raw.evalUniformNumericLookupBatch(
      kindsPtr,
      matchModesPtr,
      startsPtr,
      stepsPtr,
      lengthsPtr,
      repeatedRunLengthsPtr,
      lookupTagsPtr,
      lookupNumbersPtr,
      outTagsPtr,
      outNumbersPtr,
      outErrorsPtr,
    )
    arrayBridge.copyTypedArray(outTagsPtr, outTags, uint8ArraySpec)
    arrayBridge.copyTypedArray(outNumbersPtr, outNumbers, float64ArraySpec)
    arrayBridge.copyTypedArray(outErrorsPtr, outErrors, uint16ArraySpec)
  } finally {
    raw.__unpin(kindsPtr)
    raw.__unpin(matchModesPtr)
    raw.__unpin(startsPtr)
    raw.__unpin(stepsPtr)
    raw.__unpin(lengthsPtr)
    raw.__unpin(repeatedRunLengthsPtr)
    raw.__unpin(lookupTagsPtr)
    raw.__unpin(lookupNumbersPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outErrorsPtr)
  }
}
