import { describe, expect, it, vi } from 'vitest'
import { createActor } from 'xstate'
import { createWorkerRuntimeMachine, getWorkerRuntimeHandle } from '../runtime-machine.js'
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

function isUnknownRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function recordProperty(value: unknown, key: string): Record<string, unknown> | undefined {
  if (!isUnknownRecord(value)) {
    return undefined
  }
  const nestedValue = value[key]
  return isUnknownRecord(nestedValue) ? nestedValue : undefined
}

function createController(
  selection = { sheetName: 'Sheet1', address: 'A1' },
  runtimeStateOverride: Partial<WorkerRuntimeSessionController['runtimeState']> = {},
): WorkerRuntimeSessionController {
  const runtimeState = {
    workbookName: 'bilig-demo',
    sheets: [{ id: 1, name: 'Sheet1', order: 0 }],
    sheetNames: ['Sheet1'],
    definedNames: [],
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
    localHistoryState: { canUndo: false, canRedo: false },
    ...runtimeStateOverride,
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

  it('does not flash back to the startup selection when the session becomes ready after local navigation', async () => {
    const controller = createController({ sheetName: 'Sheet1', address: 'A1' })
    let releaseSession: (() => void) | null = null
    const sessionReadyGate = new Promise<void>((resolve) => {
      releaseSession = resolve
    })
    const createSession = vi.fn(
      async (
        _input: CreateWorkerRuntimeSessionInput,
        callbacks: WorkerRuntimeSessionCallbacks,
      ): Promise<WorkerRuntimeSessionController> => {
        callbacks.onPhase?.('hydratingLocal')
        callbacks.onRuntimeState(controller.runtimeState)
        await sessionReadyGate
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
    const nextSelection = { sheetName: 'Sheet1', address: 'B9' } as const
    actor.send({ type: 'selection.changed', selection: nextSelection })
    expect(actor.getSnapshot().context.selection).toEqual(nextSelection)

    releaseSession?.()
    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'localReady' })).toBe(true)
    })

    expect(actor.getSnapshot().context.selection).toEqual(nextSelection)
    await vi.waitFor(() => {
      expect(controller.setSelection).toHaveBeenCalledWith(nextSelection)
    })
    actor.stop()
  })

  it('accepts runtime-reconciled selections when startup sheet names replace the default sheet', async () => {
    const reconciledSelection = { sheetName: 'Dashboard', address: 'A1' } as const
    const controller = createController(reconciledSelection, {
      sheets: [
        { id: 1, name: 'Dashboard', order: 0 },
        { id: 2, name: 'Ledger', order: 1 },
      ],
      sheetNames: ['Dashboard', 'Ledger'],
    })
    const createSession = vi.fn(
      async (
        _input: CreateWorkerRuntimeSessionInput,
        callbacks: WorkerRuntimeSessionCallbacks,
      ): Promise<WorkerRuntimeSessionController> => {
        callbacks.onRuntimeState(controller.runtimeState)
        callbacks.onSelection(reconciledSelection)
        return controller
      },
    )

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: 'xlsx:imported',
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

    expect(actor.getSnapshot().context.selection).toEqual(reconciledSelection)
    actor.stop()
  })

  it('normalizes controller runtime state when the session becomes ready', async () => {
    const controller = {
      ...createController(
        { sheetName: 'Actuals', address: 'A1' },
        {
          sheets: [
            { id: 42, name: 'Actuals', order: 0 },
            { id: 9, name: 'Archive', order: 1 },
          ],
          sheetNames: ['StaleName'],
        },
      ),
      invoke: vi.fn(() => new Promise<never>(() => {})),
    }
    const createSession = vi.fn(async (): Promise<WorkerRuntimeSessionController> => controller)

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: 'book-1',
        replicaId: 'browser:test',
        persistState: true,
        connectionStateName: 'closed',
        initialSelection: { sheetName: 'Actuals', address: 'A1' },
        createSession,
      },
    })

    actor.start()
    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'localReady' })).toBe(true)
    })

    expect(actor.getSnapshot().context.runtimeState?.sheetNames).toEqual(['Actuals', 'Archive'])
    actor.stop()
  })

  it('fails startup when the ready session exposes invalid runtime state', async () => {
    const controller = createController()
    Object.defineProperty(controller, 'runtimeState', {
      configurable: true,
      value: {
        ...controller.runtimeState,
        syncState: 'not-a-runtime-sync-state',
      },
    })
    const createSession = vi.fn(async (): Promise<WorkerRuntimeSessionController> => controller)

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

    expect(actor.getSnapshot().context.error).toBe('Runtime session returned invalid runtime state')
    expect(controller.dispose).toHaveBeenCalledTimes(1)
    actor.stop()
  })

  it('does not store malformed runtime-state events from the session actor', async () => {
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

    const validRuntimeState = actor.getSnapshot().context.runtimeState
    expect(validRuntimeState?.syncState).toBe('local-only')

    actor.send({
      type: 'session.runtime',
      runtimeState: {
        ...controller.runtimeState,
        pendingMutationSummary: {
          activeCount: -1,
          failedCount: 0,
          firstFailed: null,
        },
      },
    })

    expect(actor.getSnapshot().context.runtimeState).toBe(validRuntimeState)
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

  it('keeps cyclic runtime resources out of persisted xstate snapshots', async () => {
    const controller = createController()
    const cyclicHandle = createWorkerHandle() as WorkerHandle & { self?: unknown }
    cyclicHandle.self = cyclicHandle
    const cyclicZero = {
      materialize: () => ({ data: null, addListener: () => () => {}, destroy() {} }),
    } as NonNullable<CreateWorkerRuntimeSessionInput['zero']> & { self?: unknown }
    cyclicZero.self = cyclicZero
    const cyclicController: WorkerRuntimeSessionController = {
      ...controller,
      handle: cyclicHandle,
    }
    const createSession = vi.fn(async (): Promise<WorkerRuntimeSessionController> => {
      return cyclicController
    })

    const actor = createActor(createWorkerRuntimeMachine(), {
      input: {
        documentId: 'book-1',
        replicaId: 'browser:test',
        persistState: true,
        connectionStateName: 'connected',
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        zero: cyclicZero,
        createSession,
      },
    })

    actor.start()

    await vi.waitFor(() => {
      expect(actor.getSnapshot().matches({ active: 'live' })).toBe(true)
    })

    expect(getWorkerRuntimeHandle(actor.getSnapshot().context)).toBe(cyclicHandle)
    expect(() => JSON.stringify(actor.getPersistedSnapshot())).not.toThrow()
    const persistedSnapshot: unknown = actor.getPersistedSnapshot()
    const persistedContext = recordProperty(persistedSnapshot, 'context')
    const persistedSessionInput = recordProperty(persistedContext, 'sessionInput')
    expect(Object.keys(persistedContext ?? {})).not.toContain('controller')
    expect(Object.keys(persistedContext ?? {})).not.toContain('handle')
    expect(Object.keys(persistedSessionInput ?? {})).not.toContain('zero')
    expect(Object.keys(persistedSessionInput ?? {})).not.toContain('createSession')

    actor.stop()
  })
})
