import type { SchedulerResult } from '../../scheduler.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'

export interface EngineDirtyFrontierSchedulerService {
  readonly collectDirty: (changedRoots: readonly number[] | U32) => SchedulerResult
}

export function createEngineDirtyFrontierSchedulerService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'formulas' | 'ranges' | 'scheduler'>
  readonly getEntityDependents: (entityId: number) => Uint32Array
}): EngineDirtyFrontierSchedulerService {
  return {
    collectDirty(changedRoots) {
      return args.state.scheduler.collectDirty(
        changedRoots,
        { getDependents: (entityId) => args.getEntityDependents(entityId) },
        args.state.workbook.cellStore,
        (cellIndex) => args.state.formulas.has(cellIndex),
        args.state.ranges.size,
      )
    },
  }
}
