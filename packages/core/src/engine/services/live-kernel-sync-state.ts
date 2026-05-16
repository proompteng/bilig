import type { EngineRuntimeScratchService } from './runtime-scratch-service.js'
import type { U32 } from '../runtime-state.js'

export function deferKernelSyncNow(args: {
  readonly scratch: EngineRuntimeScratchService
  readonly cellStoreSize: number
  readonly cellIndices: readonly number[] | U32
}): void {
  if (args.cellIndices.length === 0) {
    return
  }
  args.scratch.ensureRecalcCapacityNow(args.cellStoreSize + 1)
  const pendingKernelSync = args.scratch.getPendingKernelSyncNow()
  let deferredCount = args.scratch.getDeferredKernelSyncCountNow()
  let deferredEpoch = args.scratch.getDeferredKernelSyncEpochNow() + 1
  const deferredSeen = args.scratch.getDeferredKernelSyncSeenNow()
  if (deferredEpoch === 0xffff_ffff) {
    deferredEpoch = 1
    deferredSeen.fill(0)
  }
  args.scratch.setDeferredKernelSyncEpochNow(deferredEpoch)
  for (let index = 0; index < deferredCount; index += 1) {
    const cellIndex = pendingKernelSync[index]
    if (cellIndex !== undefined) {
      deferredSeen[cellIndex] = deferredEpoch
    }
  }
  for (let index = 0; index < args.cellIndices.length; index += 1) {
    const cellIndex = args.cellIndices[index]!
    if (deferredSeen[cellIndex] === deferredEpoch) {
      continue
    }
    deferredSeen[cellIndex] = deferredEpoch
    pendingKernelSync[deferredCount] = cellIndex
    deferredCount += 1
  }
  args.scratch.setDeferredKernelSyncCountNow(deferredCount)
}
