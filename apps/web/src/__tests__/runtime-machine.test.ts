import { describe, expect, it, vi } from 'vitest'
import { createActor } from 'xstate'
import { createWorkerRuntimeMachine } from '../runtime-machine.js'
import { ProjectedViewportStore } from '../projected-viewport-store.js'
import type {
  CreateWorkerRuntimeSessionInput,
  WorkerHandle,
  WorkerRuntimeSessionCallbacks,
  WorkerRuntimeSessionController,
} from '../runtime-session.js'

function createWorkerHandle(): WorkerHandle {
  return {
    viewportStore: new ProjectedViewportStore(),
  }
}

function createController(selection = { sheetName: 'Sheet1', address: 'A1' }): WorkerRuntimeSessionController {
  const runtimeState = {
    workbookName: 'bilig-demo',
    sheetNames: ['Sheet1'],
    metrics: {
      batchId: 0,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0,
    },
    syncState: 'local-only',
  } as const
  return {
    handle: createWorkerHandle(),
    runtimeState,
    selection,
    invoke: vi.fn(async (method, ...args) => {
      if (method === 'setExternalSyncState') {
        const nextSyncState = args[0]
        return {
          ...runtimeState,
          syncState: typeof nextSyncState === 'string' ? nextSyncState : runtimeState.syncState,
        }
      }
      return undefined
    }),
    setSelection: vi.fn(async () => undefined),
    subscribeViewport: () => () => {},
    dispose: vi.fn(),
  }
}

describe('worker runtime machine', () => {
  it('boots into localReady and forwards selection changes to the session controller', async () => {
    const controller = createController()
    const createSession = vi.fn(
      async (
        _input: CreateWorkerRuntimeSessionInput,
        callbacks: WorkerRuntimeSessionCallbacks,
      ): Promise<WorkerRuntimeSessionController> => {
        callbacks.onPhase?.('hydratingLocal')
        callbacks.onRuntimeState(controller.runtimeState)
        return controller
      },
    )

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: 'book-1',
        replicaId: 'browser:test',
        persistState: true,
        connectionStateName: 'closed',
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createSession,
      },
    })

    actor.start()
    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'localReady' })).toBe(true)
    })

    const nextSelection = { sheetName: 'Sheet1', address: 'C3' } as const
    actor.send({ type: 'selection.changed', selection: nextSelection })

    await vi.waitFor(() => {
      expect(controller.setSelection).toHaveBeenCalledWith(nextSelection)
    })

    expect(actor.getSnapshot().context.selection).toEqual(nextSelection)
    actor.stop()
  })

  it('ignores stale session selection echoes after a newer local selection', async () => {
    const controller = createController()
    let sessionCallbacks: WorkerRuntimeSessionCallbacks | null = null
    const createSession = vi.fn(
      async (
        _input: CreateWorkerRuntimeSessionInput,
        callbacks: WorkerRuntimeSessionCallbacks,
      ): Promise<WorkerRuntimeSessionController> => {
        sessionCallbacks = callbacks
        callbacks.onPhase?.('hydratingLocal')
        callbacks.onRuntimeState(controller.runtimeState)
        return controller
      },
    )

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: 'book-1',
        replicaId: 'browser:test',
        persistState: true,
        connectionStateName: 'closed',
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createSession,
      },
    })

    actor.start()
    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'localReady' })).toBe(true)
    })

    const nextSelection = { sheetName: 'Sheet1', address: 'C3' } as const
    actor.send({ type: 'selection.changed', selection: nextSelection })
    sessionCallbacks?.onSelection?.({ sheetName: 'Sheet1', address: 'A1' })

    expect(actor.getSnapshot().context.selection).toEqual(nextSelection)

    sessionCallbacks?.onSelection?.(nextSelection)
    expect(actor.getSnapshot().context.selection).toEqual(nextSelection)

    actor.stop()
  })

  it('tracks connected and offline steady states and transient rebase phases', async () => {
    const controller = createController()
    const createSession = vi.fn(
      async (
        _input: CreateWorkerRuntimeSessionInput,
        callbacks: WorkerRuntimeSessionCallbacks,
      ): Promise<WorkerRuntimeSessionController> => {
        callbacks.onPhase?.('hydratingLocal')
        callbacks.onRuntimeState(controller.runtimeState)
        return controller
      },
    )

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: 'book-1',
        replicaId: 'browser:test',
        persistState: true,
        connectionStateName: 'connected',
        zero: { materialize: () => ({ data: null, addListener: () => () => {}, destroy() {} }) },
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createSession,
      },
    })

    actor.start()

    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'live' })).toBe(true)
    })

    actor.send({ type: 'connection.changed', connectionStateName: 'disconnected' })
    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'offline' })).toBe(true)
    })

    actor.send({ type: 'session.phase', phase: 'reconciling' })
    expect(actor.getSnapshot().matches({ active: 'reconciling' })).toBe(true)

    actor.send({ type: 'session.phase', phase: 'recovering' })
    expect(actor.getSnapshot().matches({ active: 'recovering' })).toBe(true)

    actor.send({ type: 'connection.changed', connectionStateName: 'connected' })
    actor.send({ type: 'session.phase', phase: 'steady' })
    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'live' })).toBe(true)
    })

    expect(controller.invoke).toHaveBeenCalledWith('setExternalSyncState', 'live')
    expect(controller.invoke).toHaveBeenCalledWith('setExternalSyncState', 'reconnecting')

    actor.stop()
  })

  it('transitions to failed when session startup rejects and recovers on retry', async () => {
    const controller = createController()
    const createSession = vi
      .fn<(input: CreateWorkerRuntimeSessionInput, callbacks: WorkerRuntimeSessionCallbacks) => Promise<WorkerRuntimeSessionController>>()
      .mockRejectedValueOnce(new Error('bootstrap failed'))
      .mockResolvedValueOnce(controller)

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: 'book-1',
        replicaId: 'browser:test',
        persistState: true,
        connectionStateName: 'closed',
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createSession,
      },
    })

    actor.start()

    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches('failed')).toBe(true)
    })
    expect(actor.getSnapshot().context.error).toBe('bootstrap failed')

    actor.send({ type: 'retry' })

    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'localReady' })).toBe(true)
    })

    actor.stop()
  })

  it('restarts the active session when retry requests a new persistence mode', async () => {
    const initialController = createController()
    const restartedController = createController()
    const createSession = vi
      .fn<(input: CreateWorkerRuntimeSessionInput, callbacks: WorkerRuntimeSessionCallbacks) => Promise<WorkerRuntimeSessionController>>()
      .mockResolvedValueOnce(initialController)
      .mockResolvedValueOnce(restartedController)

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: 'book-1',
        replicaId: 'browser:test',
        persistState: true,
        connectionStateName: 'closed',
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createSession,
      },
    })

    actor.start()

    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'localReady' })).toBe(true)
    })

    actor.send({ type: 'retry', persistState: false })

    await vi.waitFor(() => {
      expect(createSession).toHaveBeenCalledTimes(2)
    })

    expect(initialController.dispose).toHaveBeenCalledTimes(1)
    expect(createSession.mock.calls[1]?.[0].persistState).toBe(false)
    expect(actor.getSnapshot().context.persistState).toBe(false)

    actor.stop()
  })
})
