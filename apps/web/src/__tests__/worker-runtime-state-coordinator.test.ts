import { describe, expect, it } from 'vitest'
import { EMPTY_RUNTIME_METRICS } from '../worker-runtime-state.js'
import { WorkerRuntimeStateCoordinator, type WorkerRuntimeStateEngine } from '../worker-runtime-state-coordinator.js'

function createStateEngine(input: {
  readonly workbookName?: string
  readonly syncState?: WorkerRuntimeStateEngine['getSyncState'] extends () => infer T ? T : never
  readonly canUndo?: boolean
  readonly canRedo?: boolean
}): WorkerRuntimeStateEngine {
  return {
    workbook: {
      workbookName: input.workbookName ?? 'Book',
      sheetsByName: new Map([['Sheet1', { id: 1, name: 'Sheet1', order: 0 }]]),
    },
    getDefinedNames: () => [],
    getLastMetrics: () => ({ ...EMPTY_RUNTIME_METRICS, batchId: 7 }),
    getSyncState: () => input.syncState ?? 'live',
    canUndo: () => input.canUndo === true,
    canRedo: () => input.canRedo === true,
  }
}

describe('WorkerRuntimeStateCoordinator', () => {
  it('publishes cached state with live runtime authority fields', () => {
    let currentEngine: WorkerRuntimeStateEngine | null = createStateEngine({ syncState: 'live', canUndo: false })
    let authoritativeRevision = 3
    const coordinator = new WorkerRuntimeStateCoordinator({
      getEngine: () => currentEngine,
      getAuthoritativeRevision: () => authoritativeRevision,
      buildPendingMutationSummary: () => ({
        activeCount: 2,
        failedCount: 1,
        firstFailed: {
          id: 'mutation-1',
          method: 'setCellValue',
          failureMessage: 'network unavailable',
          attemptCount: 2,
        },
      }),
    })

    coordinator.updateRuntimeStateFromEngine(currentEngine)
    coordinator.setExternalSyncState('reconnecting')
    authoritativeRevision = 9
    currentEngine = createStateEngine({ syncState: 'live', canUndo: true, canRedo: false })

    const publicState = coordinator.getRuntimeState(() => currentEngine)

    expect(publicState).toMatchObject({
      workbookName: 'Book',
      sheetNames: ['Sheet1'],
      syncState: 'reconnecting',
      localHistoryState: { canUndo: true, canRedo: false },
      authoritativeRevision: 9,
      pendingMutationSummary: {
        activeCount: 2,
        failedCount: 1,
      },
      localPersistenceMode: 'ephemeral',
    })
    expect(publicState.pendingMutationSummary?.firstFailed).toEqual({
      id: 'mutation-1',
      method: 'setCellValue',
      failureMessage: 'network unavailable',
      attemptCount: 2,
    })
  })

  it('clears cached state and external sync overrides on reset', () => {
    let currentEngine = createStateEngine({ syncState: 'live' })
    const coordinator = new WorkerRuntimeStateCoordinator({
      getEngine: () => currentEngine,
      getAuthoritativeRevision: () => 0,
      buildPendingMutationSummary: () => ({ activeCount: 0, failedCount: 0, firstFailed: null }),
    })

    coordinator.updateRuntimeStateFromEngine(currentEngine)
    coordinator.setExternalSyncState('reconnecting')
    currentEngine = createStateEngine({ workbookName: 'AfterReset', syncState: 'live' })

    coordinator.reset()

    expect(coordinator.getRuntimeState(() => currentEngine)).toMatchObject({
      workbookName: 'AfterReset',
      syncState: 'live',
      localPersistenceMode: 'ephemeral',
    })
  })
})
