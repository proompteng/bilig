import type { RawKernelExports } from './raw-kernel-exports.js'
import { RawKernelArrayBridge, uint32ArraySpec, uint8ArraySpec } from './raw-kernel-array-bridge.js'

export function materializePivotTableRaw(
  raw: RawKernelExports,
  sourceRangeIndex: number,
  sourceWidth: number,
  groupByColumnIndices: Uint32Array,
  valueColumnIndices: Uint32Array,
  valueAggregations: Uint8Array,
): void {
  const arrayBridge = new RawKernelArrayBridge(raw)
  const [groupByPtr, valueColsPtr, valueAggsPtr] = arrayBridge.lowerTypedArrays([
    [groupByColumnIndices, uint32ArraySpec],
    [valueColumnIndices, uint32ArraySpec],
    [valueAggregations, uint8ArraySpec],
  ])
  try {
    raw.materializePivotTable(
      sourceRangeIndex,
      sourceWidth,
      groupByColumnIndices.length,
      groupByPtr,
      valueColumnIndices.length,
      valueColsPtr,
      valueAggsPtr,
    )
  } finally {
    raw.__unpin(groupByPtr)
    raw.__unpin(valueColsPtr)
    raw.__unpin(valueAggsPtr)
  }
}
