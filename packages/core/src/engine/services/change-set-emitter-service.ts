import type { EngineChangedCell } from '@bilig/protocol'
import type { EngineRuntimeState } from '../runtime-state.js'
import type { EngineCellPatch } from '../../patches/patch-types.js'
import { materializeChangedCells, materializeChangedCellPatches } from '../../patches/materialize-changed-cells.js'

export interface EngineChangeSetEmitterService {
  readonly captureChangedCells: (changedCellIndices: readonly number[] | Uint32Array) => readonly EngineChangedCell[]
  readonly captureChangedPatches: (changedCellIndices: readonly number[] | Uint32Array) => readonly EngineCellPatch[]
}

export function createEngineChangeSetEmitterService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'> & { counters?: EngineRuntimeState['counters'] }
}): EngineChangeSetEmitterService {
  return {
    captureChangedCells(changedCellIndices) {
      return materializeChangedCells(args.state, changedCellIndices)
    },
    captureChangedPatches(changedCellIndices) {
      return materializeChangedCellPatches(args.state, changedCellIndices)
    },
  }
}
