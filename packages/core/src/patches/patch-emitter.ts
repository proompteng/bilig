import type { EngineRuntimeState } from '../engine/runtime-state.js'
import { materializeChangedCellPatches } from './materialize-changed-cells.js'
import type { EngineCellPatch } from './patch-types.js'

export interface EnginePatchEmitterService {
  readonly captureChangedPatches: (changedCellIndices: readonly number[] | Uint32Array) => readonly EngineCellPatch[]
}

export function createEnginePatchEmitterService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'> & { counters?: EngineRuntimeState['counters'] }
}): EnginePatchEmitterService {
  return {
    captureChangedPatches(changedCellIndices) {
      return materializeChangedCellPatches(args.state, changedCellIndices)
    },
  }
}
