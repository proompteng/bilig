import { afterEach, describe, expect, it, vi } from 'vitest'
import { createWorkerEngineHost } from '@bilig/worker-transport'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { createEmptyWorkbookSnapshot } from '@bilig/zero-sync'
import { createWorkerRuntimeSessionController, type WorkerRuntimeSessionController } from '../runtime-session.js'
import type { WorkbookWorkerStateSnapshot } from '../worker-runtime.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'

const metrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function runtimeState(overrides: Partial<WorkbookWorkerStateSnapshot> = {}): WorkbookWorkerStateSnapshot {
  return {
    workbookName: 'doc-1',
    sheets: [{ id: 1, name: 'Sheet1', order: 0 }],
    sheetNames: ['Sheet1'],
    definedNames: [],
    metrics,
    syncState: 'live',
    localHistoryState: {
      canUndo: false,
      canRedo: false,
    },
    authoritativeRevision: 3,
    pendingMutationSummary: {
      activeCount: 1,
      failedCount: 0,
      firstFailed: null,
    },
    localPersistenceMode: 'ephemeral',
    ...overrides,
  }
}

class FakeRevisionLiveView {
  readonly listeners = new Set<(value: unknown) => void>()
  destroy = vi.fn()

  constructor(readonly data: unknown) {}

  addListener(listener: (value: unknown) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  emit(value: unknown): void {
    this.listeners.forEach((listener) => {
      listener(value)
    })
  }
}

describe('worker runtime session authoritative event loading', () => {
  let channel: MessageChannel | null = null
  let controller: WorkerRuntimeSessionController | null = null
  let host: { dispose(): void } | null = null

  afterEach(() => {
    controller?.dispose()
    host?.dispose()
    channel?.port1.close()
    channel?.port2.close()
    vi.unstubAllGlobals()
    channel = null
    controller = null
    host = null
  })

  it('publishes pending mutation state before enqueue resolves', async () => {
    channel = new MessageChannel()
    let state = runtimeState({
      authoritativeRevision: 0,
      pendingMutationSummary: {
        activeCount: 0,
        failedCount: 0,
        firstFailed: null,
      },
    })
    const runtimeStates: WorkbookWorkerStateSnapshot[] = []
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: state,
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 0
        },
        getRuntimeState() {
          return state
        },
        getCell(sheetName: string, address: string) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }
        },
        enqueuePendingMutation() {
          state = runtimeState({
            authoritativeRevision: 0,
            pendingMutationSummary: {
              activeCount: 1,
              failedCount: 0,
              firstFailed: null,
            },
          })
          return {
            id: 'doc-1:browser:test:pending:1',
            localSeq: 1,
            baseRevision: 0,
            method: 'setRangeStyle',
            args: [{ sheetName: 'Sheet1', startAddress: 'E6', endAddress: 'E6' }, { fill: { backgroundColor: '#00ff00' } }],
            enqueuedAtUnixMs: 1,
            submittedAtUnixMs: null,
            lastAttemptedAtUnixMs: null,
            ackedAtUnixMs: null,
            rebasedAtUnixMs: null,
            failedAtUnixMs: null,
            attemptCount: 0,
            failureMessage: null,
            status: 'local',
          }
        },
        subscribeViewportPatches() {
          return () => undefined
        },
      },
      channel.port1,
    )

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        authoritativeSyncEnabled: false,
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createWorker: () => channel!.port2,
      },
      {
        onError: (message) => {
          throw new Error(message)
        },
        onRuntimeState: (snapshot) => {
          runtimeStates.push(snapshot)
        },
        onSelection: () => undefined,
      },
    )

    await controller.invoke('enqueuePendingMutation', {
      method: 'setRangeStyle',
      args: [{ sheetName: 'Sheet1', startAddress: 'E6', endAddress: 'E6' }, { fill: { backgroundColor: '#00ff00' } }],
    })

    expect(runtimeStates.at(-1)?.pendingMutationSummary?.activeCount).toBe(1)
    expect(controller.runtimeState.pendingMutationSummary?.activeCount).toBe(1)
  })

  it('keeps the browser fetch receiver intact for default authoritative refreshes', async () => {
    channel = new MessageChannel()
    const applyAuthoritativeEvents = vi.fn()
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: runtimeState({
              authoritativeRevision: 0,
              pendingMutationSummary: {
                activeCount: 0,
                failedCount: 0,
                firstFailed: null,
              },
            }),
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 0
        },
        getCell(sheetName: string, address: string) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }
        },
        applyAuthoritativeEvents,
      },
      channel.port1,
    )
    const fetchImpl = vi.fn(async function (this: typeof globalThis, input: RequestInfo | URL, init?: RequestInit) {
      expect(this).toBe(globalThis)
      expect(input).toBe('/v2/documents/doc-1/events?afterRevision=0')
      expect(init).toEqual({
        headers: {
          accept: 'application/json',
        },
        cache: 'no-store',
      })
      return new Response(
        JSON.stringify({
          afterRevision: 0,
          headRevision: 0,
          calculatedRevision: 0,
          events: [],
        }),
        {
          status: 200,
          headers: { 'content-type': 'application/json' },
        },
      )
    })
    vi.stubGlobal('fetch', fetchImpl)

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError: vi.fn(),
      },
    )

    fetchImpl.mockClear()
    await expect(controller.invoke('refreshAuthoritativeEvents')).resolves.toEqual(undefined)
    expect(fetchImpl).toHaveBeenCalledWith('/v2/documents/doc-1/events?afterRevision=0', {
      headers: {
        accept: 'application/json',
      },
      cache: 'no-store',
    })
    expect(applyAuthoritativeEvents).not.toHaveBeenCalled()
  })

  it('rejects authoritative event responses whose cursor does not match the requested revision', async () => {
    channel = new MessageChannel()
    const applyAuthoritativeEvents = vi.fn()
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: runtimeState(),
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 3
        },
        getCell(sheetName: string, address: string) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }
        },
        applyAuthoritativeEvents,
      },
      channel.port1,
    )
    const fetchImpl = vi.fn(
      async () =>
        new Response(
          JSON.stringify({
            afterRevision: 0,
            headRevision: 5,
            calculatedRevision: 5,
            events: [
              {
                revision: 1,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A1',
                  value: 1,
                },
              },
              {
                revision: 2,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A2',
                  value: 2,
                },
              },
              {
                revision: 3,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A3',
                  value: 3,
                },
              },
              {
                revision: 4,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A4',
                  value: 4,
                },
              },
              {
                revision: 5,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A5',
                  value: 5,
                },
              },
            ],
          }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        ),
    )
    const onError = vi.fn()

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        fetchImpl,
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError,
      },
    )

    await expect(controller.invoke('refreshAuthoritativeEvents')).rejects.toThrow(
      'Authoritative event payload does not match the expected schema',
    )
    expect(fetchImpl).toHaveBeenCalledWith('/v2/documents/doc-1/events?afterRevision=3', {
      headers: {
        accept: 'application/json',
      },
      cache: 'no-store',
    })
    expect(applyAuthoritativeEvents).not.toHaveBeenCalled()
    expect(onError).toHaveBeenCalledWith('Authoritative event payload does not match the expected schema')
  })

  it('subscribes to live zero revisions only after authoritative snapshot bootstrap', async () => {
    channel = new MessageChannel()
    const calls: string[] = []
    const revisionView = new FakeRevisionLiveView({
      headRevision: 0,
      calculatedRevision: 0,
    })
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          calls.push('worker-bootstrap')
          return {
            runtimeState: runtimeState({
              authoritativeRevision: 0,
              pendingMutationSummary: {
                activeCount: 0,
                failedCount: 0,
                firstFailed: null,
              },
            }),
            restoredFromPersistence: false,
            requiresAuthoritativeHydrate: true,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 0
        },
        getCell(sheetName: string, address: string) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }
        },
        subscribeViewportPatches() {
          return () => undefined
        },
      },
      channel.port1,
    )
    const fetchImpl = vi.fn(async (url: string) => {
      calls.push(`fetch:${url}`)
      return new Response(null, { status: 204 })
    })
    const zero = {
      materialize: vi.fn(() => {
        calls.push('zero-materialize')
        return revisionView
      }),
    }

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        fetchImpl,
        zero,
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError: vi.fn(),
      },
    )

    expect(fetchImpl).toHaveBeenCalledWith('/v2/documents/doc-1/snapshot/latest', {
      headers: {
        accept: 'application/json, application/vnd.bilig.workbook+json',
      },
      cache: 'no-store',
    })
    expect(zero.materialize).toHaveBeenCalledOnce()
    expect(calls.indexOf('fetch:/v2/documents/doc-1/snapshot/latest')).toBeGreaterThanOrEqual(0)
    expect(calls.indexOf('fetch:/v2/documents/doc-1/snapshot/latest')).toBeLessThan(calls.indexOf('zero-materialize'))
  })

  it('keeps the local workbook visible when required authoritative snapshot loading fails', async () => {
    channel = new MessageChannel()
    const revisionView = new FakeRevisionLiveView({
      headRevision: 0,
      calculatedRevision: 0,
    })
    const getCell = vi.fn((sheetName: string, address: string) => ({
      sheetName,
      address,
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    }))
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: runtimeState({
              authoritativeRevision: 0,
              pendingMutationSummary: {
                activeCount: 0,
                failedCount: 0,
                firstFailed: null,
              },
            }),
            restoredFromPersistence: false,
            requiresAuthoritativeHydrate: true,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 0
        },
        getCell,
        subscribeViewportPatches() {
          return () => undefined
        },
      },
      channel.port1,
    )
    const fetchImpl = vi.fn(async (url: RequestInfo | URL) => {
      const requestUrl = typeof url === 'string' ? url : url instanceof URL ? url.href : url.url
      if (requestUrl.endsWith('/snapshot/latest')) {
        return new Response('server error', { status: 500 })
      }
      return new Response(
        JSON.stringify({
          afterRevision: 0,
          headRevision: 0,
          calculatedRevision: 0,
          events: [],
        }),
        {
          status: 200,
          headers: { 'content-type': 'application/json' },
        },
      )
    })
    const zero = {
      materialize: vi.fn(() => revisionView),
    }
    const onRuntimeState = vi.fn()
    const onError = vi.fn()

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        fetchImpl,
        zero,
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState,
        onSelection: vi.fn(),
        onError,
      },
    )

    expect(controller.runtimeState.sheetNames).toEqual(['Sheet1'])
    expect(controller.selection).toEqual({ sheetName: 'Sheet1', address: 'A1' })
    expect(getCell).toHaveBeenCalledWith('Sheet1', 'A1')
    expect(zero.materialize).toHaveBeenCalledOnce()
    expect(onRuntimeState).toHaveBeenCalledWith(expect.objectContaining({ sheetNames: ['Sheet1'] }))
    expect(onError).toHaveBeenCalledWith('Failed to load workbook snapshot (500)')
    await new Promise((resolve) => setTimeout(resolve, 20))
    expect(
      fetchImpl.mock.calls
        .map(([url]) => (typeof url === 'string' ? url : url instanceof URL ? url.href : url.url))
        .filter((url) => url.endsWith('/events?afterRevision=0')),
    ).toEqual([])
  })

  it('polls authoritative events after bootstrap when the live revision callback misses an update', async () => {
    channel = new MessageChannel()
    let authoritativeRevision = 0
    const applyAuthoritativeEvents = vi.fn((events: unknown[], headRevision: number) => {
      authoritativeRevision = headRevision
      expect(events).toHaveLength(1)
      return runtimeState({
        authoritativeRevision: headRevision,
        pendingMutationSummary: {
          activeCount: 0,
          failedCount: 0,
          firstFailed: null,
        },
      })
    })
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: runtimeState({
              authoritativeRevision: 0,
              pendingMutationSummary: {
                activeCount: 0,
                failedCount: 0,
                firstFailed: null,
              },
            }),
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return authoritativeRevision
        },
        getCell(sheetName: string, address: string) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }
        },
        applyAuthoritativeEvents,
        subscribeViewportPatches() {
          return () => undefined
        },
      },
      channel.port1,
    )
    const fetchImpl = vi.fn(async (url: string) => {
      if (url.endsWith('/snapshot/latest')) {
        return new Response(null, { status: 204 })
      }
      return new Response(
        JSON.stringify({
          afterRevision: authoritativeRevision,
          headRevision: authoritativeRevision === 0 ? 1 : authoritativeRevision,
          calculatedRevision: authoritativeRevision === 0 ? 1 : authoritativeRevision,
          events:
            authoritativeRevision === 0
              ? [
                  {
                    revision: 1,
                    clientMutationId: null,
                    payload: {
                      kind: 'setCellValue',
                      sheetName: 'Sheet1',
                      address: 'B2',
                      value: 'remote',
                    },
                  },
                ]
              : [],
        }),
        {
          status: 200,
          headers: { 'content-type': 'application/json' },
        },
      )
    })

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        fetchImpl,
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError: vi.fn(),
      },
    )

    await vi.waitFor(
      () => {
        expect(applyAuthoritativeEvents).toHaveBeenCalledTimes(1)
      },
      { timeout: 2_000 },
    )
    expect(fetchImpl).toHaveBeenCalledWith('/v2/documents/doc-1/events?afterRevision=0', {
      headers: {
        accept: 'application/json',
      },
      cache: 'no-store',
    })
  })

  it('refreshes a same-head authoritative snapshot when calculation catches up', async () => {
    channel = new MessageChannel()
    const revisionView = new FakeRevisionLiveView({
      headRevision: 5,
      calculatedRevision: 4,
    })
    const installAuthoritativeSnapshot = vi.fn((input: { authoritativeRevision: number }) =>
      runtimeState({
        authoritativeRevision: input.authoritativeRevision,
        pendingMutationSummary: {
          activeCount: 0,
          failedCount: 0,
          firstFailed: null,
        },
      }),
    )
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: runtimeState({
              authoritativeRevision: 0,
              pendingMutationSummary: {
                activeCount: 0,
                failedCount: 0,
                firstFailed: null,
              },
            }),
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 0
        },
        getCell(sheetName: string, address: string) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }
        },
        installAuthoritativeSnapshot,
        subscribeViewportPatches() {
          return () => undefined
        },
      },
      channel.port1,
    )
    const fetchImpl = vi.fn(async () => {
      return new Response(JSON.stringify(createEmptyWorkbookSnapshot('doc-1')), {
        status: 200,
        headers: {
          'content-type': 'application/json',
          'x-bilig-snapshot-cursor': '5',
        },
      })
    })

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        fetchImpl,
        zero: {
          materialize: vi.fn(() => revisionView),
        },
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError: vi.fn(),
      },
    )

    await vi.waitFor(() => {
      expect(installAuthoritativeSnapshot).toHaveBeenCalledTimes(1)
    })
    expect(installAuthoritativeSnapshot.mock.calls[0]?.[0]).toMatchObject({
      authoritativeRevision: 5,
      mode: 'bootstrap',
    })

    revisionView.emit({
      headRevision: 5,
      calculatedRevision: 5,
    })

    await vi.waitFor(() => {
      expect(installAuthoritativeSnapshot).toHaveBeenCalledTimes(2)
    })
    expect(installAuthoritativeSnapshot.mock.calls[1]?.[0]).toMatchObject({
      authoritativeRevision: 5,
      mode: 'reconcile',
    })
  })

  it('hydrates the selected cell after local undo supersedes an optimistic clear', async () => {
    channel = new MessageChannel()
    let selectedCell: CellSnapshot = {
      sheetName: 'Sheet1',
      address: 'D12',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    }
    let state = runtimeState({
      syncState: 'local-only',
      localHistoryState: { canUndo: true, canRedo: false },
      pendingMutationSummary: {
        activeCount: 2,
        failedCount: 0,
        firstFailed: null,
      },
    })
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: state,
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 0
        },
        getRuntimeState() {
          return state
        },
        getCell(sheetName: string, address: string) {
          return selectedCell.sheetName === sheetName && selectedCell.address === address
            ? selectedCell
            : {
                sheetName,
                address,
                value: { tag: ValueTag.Empty },
                flags: 0,
                version: 0,
              }
        },
        undoLocalChange() {
          selectedCell = {
            sheetName: 'Sheet1',
            address: 'D12',
            value: { tag: ValueTag.String, value: 'delete-undo-redo', stringId: 1 },
            input: 'delete-undo-redo',
            flags: 0,
            version: 1,
          }
          state = runtimeState({
            syncState: 'local-only',
            localHistoryState: { canUndo: true, canRedo: true },
            pendingMutationSummary: {
              activeCount: 2,
              failedCount: 0,
              firstFailed: null,
            },
          })
          return true
        },
        subscribeViewportPatches() {
          return () => undefined
        },
      },
      channel.port1,
    )

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        authoritativeSyncEnabled: false,
        initialSelection: { sheetName: 'Sheet1', address: 'D12' },
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError: vi.fn(),
      },
    )
    controller.handle.viewportStore.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D12',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 2,
    })

    await controller.invoke('undoLocalChange')

    expect(controller.handle.viewportStore.getCell('Sheet1', 'D12')).toMatchObject({
      input: 'delete-undo-redo',
      value: { tag: ValueTag.String, value: 'delete-undo-redo', stringId: 1 },
      flags: 0,
      version: 1,
    })
  })
})
