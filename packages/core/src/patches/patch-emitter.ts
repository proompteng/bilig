import type { EngineRuntimeState } from '../engine/runtime-state.js'
import { materializeEnginePatches, type EnginePatchCaptureRequest } from './materialize-changed-cells.js'
import type { EnginePatch } from './patch-types.js'

export interface EnginePatchEmitterService {
  readonly captureChangedPatches: (
    changedCellIndices: readonly number[] | Uint32Array,
    request?: Omit<EnginePatchCaptureRequest, 'changedCellIndices'>,
  ) => readonly EnginePatch[]
}

export function createEnginePatchEmitterService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'> & { counters?: EngineRuntimeState['counters'] }
}): EnginePatchEmitterService {
  return {
    captureChangedPatches(changedCellIndices, request = {}) {
      return materializeEnginePatches(args.state, {
        changedCellIndices,
        ...request,
      })
    },
  }
}
