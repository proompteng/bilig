import type { RawKernelExports } from './raw-kernel-exports.js'
import { RawKernelArrayBridge, uint32ArraySpec } from './raw-kernel-array-bridge.js'

export function evalBatchRaw(raw: RawKernelExports, cellIndices: Uint32Array): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [cellIndicesPtr] = arrayBridge.lowerTypedArrays([[cellIndices, uint32ArraySpec]])
  try {
    raw.evalBatch(cellIndicesPtr)
  } finally {
    raw.__unpin(cellIndicesPtr)
  }
}
