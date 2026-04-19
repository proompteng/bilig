import type { EngineChangedCell } from '@bilig/protocol'
import type { EngineRuntimeState } from '../runtime-state.js'
import type { EnginePatch } from '../../patches/patch-types.js'
import {
  materializeChangedCells,
  materializeEnginePatches,
  type EnginePatchCaptureRequest,
} from '../../patches/materialize-changed-cells.js'

export interface EngineChangeSetEmitterService {
  readonly captureChangedCells: (changedCellIndices: readonly number[] | Uint32Array) => readonly EngineChangedCell[]
  readonly captureChangedPatches: (
    changedCellIndices: readonly number[] | Uint32Array,
    request?: Omit<EnginePatchCaptureRequest, 'changedCellIndices'>,
  ) => readonly EnginePatch[]
}

export function createEngineChangeSetEmitterService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'> & { counters?: EngineRuntimeState['counters'] }
}): EngineChangeSetEmitterService {
  return {
    captureChangedCells(changedCellIndices) {
      return materializeChangedCells(args.state, changedCellIndices)
    },
    captureChangedPatches(changedCellIndices, request = {}) {
      return materializeEnginePatches(args.state, {
        changedCellIndices,
        ...request,
      })
    },
  }
}
