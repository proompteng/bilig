import type { SpreadsheetEngine } from '@bilig/core'
import type { WorkerEngine } from './worker-runtime-support.js'

export interface WorkerRuntimeLocalHistoryState {
  readonly canUndo: boolean
  readonly canRedo: boolean
}

export interface WorkerRuntimeLocalHistoryContext {
  readonly getProjectionEngine: () => Promise<SpreadsheetEngine & WorkerEngine>
  readonly invalidateProjectionCache: () => void
  readonly updateRuntimeStateFromEngine: (engine: SpreadsheetEngine & WorkerEngine) => void
}

export function buildWorkerRuntimeLocalHistoryState(engine: (SpreadsheetEngine & WorkerEngine) | null): WorkerRuntimeLocalHistoryState {
  if (!engine) {
    return {
      canUndo: false,
      canRedo: false,
    }
  }
  return {
    canUndo: engine.canUndo(),
    canRedo: engine.canRedo(),
  }
}

export async function applyWorkerRuntimeLocalHistoryChange(
  context: WorkerRuntimeLocalHistoryContext,
  direction: 'undo' | 'redo',
): Promise<boolean> {
  const engine = await context.getProjectionEngine()
  const applied = direction === 'undo' ? engine.undo() : engine.redo()
  if (!applied) {
    return false
  }

  context.invalidateProjectionCache()
  context.updateRuntimeStateFromEngine(engine)
  return true
}
