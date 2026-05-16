// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { ValueTag } from '@bilig/protocol'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ProjectedViewportStore } from '../projected-viewport-store.js'
import type { WorkerHandle } from '../runtime-session.js'
import { useWorkbookSync } from '../use-workbook-sync.js'
import type { PendingWorkbookMutation } from '../workbook-sync.js'

function createPendingMutation(): PendingWorkbookMutation {
  return {
    id: 'pending-1',
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 17],
    localSeq: 1,
    baseRevision: 0,
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
}

describe('useWorkbookSync', () => {
  afterEach(() => {
    vi.useRealTimers()
    document.body.innerHTML = ''
    vi.restoreAllMocks()
  })

  it('marks non-retryable mutation failures without escalating them as runtime errors', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let pendingMutation = createPendingMutation()
    let sync: ReturnType<typeof useWorkbookSync> | null = null
    const reportRuntimeError = vi.fn()
    const runtimeController = {
      invoke: vi.fn(async (method: string, ...args: unknown[]) => {
        switch (method) {
          case 'enqueuePendingMutation':
            return pendingMutation
          case 'listPendingMutations':
            return [pendingMutation]
          case 'recordPendingMutationAttempt':
            pendingMutation = {
              ...pendingMutation,
              attemptCount: 1,
              lastAttemptedAtUnixMs: 2,
            }
            return undefined
          case 'refreshAuthoritativeEvents':
            return undefined
          case 'markPendingMutationFailed':
            pendingMutation = {
              ...pendingMutation,
              status: 'failed',
              failedAtUnixMs: 3,
              failureMessage: typeof args[1] === 'string' ? args[1] : 'mutation rejected by server',
            }
            return undefined
          default:
            throw new Error(`Unexpected runtime invoke: ${method}`)
        }
      }),
    }

    function Harness() {
      sync = useWorkbookSync({
        documentId: 'doc-1',
        connectionStateName: 'connected',
        connectionStateRef: { current: 'connected' },
        runtimeController,
        workerHandleRef: { current: null },
        zeroRef: {
          current: {
            mutate() {
              return {
                server: Promise.resolve({
                  type: 'error',
                  error: {
                    type: 'app',
                    message: 'mutation rejected by server',
                  },
                }),
              }
            },
          },
        },
        reportRuntimeError,
      })
      return null
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<Harness />)
    })

    if (!sync) {
      throw new Error('Expected useWorkbookSync harness to initialize')
    }

    await act(async () => {
      await sync!.invokeMutation('setCellValue', 'Sheet1', 'A1', 17)
    })

    await vi.waitFor(() => {
      expect(runtimeController.invoke).toHaveBeenCalledWith('markPendingMutationFailed', 'pending-1', 'mutation rejected by server')
    })
    expect(reportRuntimeError).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  it('probes authoritative events when a Zero server observer does not settle', async () => {
    vi.useFakeTimers()
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let pendingMutation = createPendingMutation()
    let sync: ReturnType<typeof useWorkbookSync> | null = null
    const reportRuntimeError = vi.fn()
    const runtimeController = {
      invoke: vi.fn(async (method: string) => {
        switch (method) {
          case 'enqueuePendingMutation':
            return pendingMutation
          case 'listPendingMutations':
            return [pendingMutation]
          case 'recordPendingMutationAttempt':
            pendingMutation = {
              ...pendingMutation,
              attemptCount: 1,
              lastAttemptedAtUnixMs: 2,
            }
            return undefined
          case 'refreshAuthoritativeEvents':
            return undefined
          default:
            throw new Error(`Unexpected runtime invoke: ${method}`)
        }
      }),
    }

    function Harness() {
      sync = useWorkbookSync({
        documentId: 'doc-1',
        connectionStateName: 'connected',
        connectionStateRef: { current: 'connected' },
        runtimeController,
        workerHandleRef: { current: null },
        zeroRef: {
          current: {
            mutate() {
              return {
                server: new Promise(() => {}),
              }
            },
          },
        },
        reportRuntimeError,
      })
      return null
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<Harness />)
    })

    if (!sync) {
      throw new Error('Expected useWorkbookSync harness to initialize')
    }

    await act(async () => {
      await sync!.invokeMutation('setCellValue', 'Sheet1', 'A1', 17)
    })
    await vi.waitFor(() => {
      expect(runtimeController.invoke).toHaveBeenCalledWith('recordPendingMutationAttempt', 'pending-1')
    })

    await act(async () => {
      await vi.advanceTimersByTimeAsync(400)
    })

    expect(runtimeController.invoke).toHaveBeenCalledWith('refreshAuthoritativeEvents')
    expect(reportRuntimeError).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  it('persists deferred axis resizes once while avoiding stale viewport store writes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const frames = installAnimationFrameQueue()
    const initialStore = new ProjectedViewportStore()
    const replacementStore = new ProjectedViewportStore()
    const initialSetColumnWidth = vi.spyOn(initialStore, 'setColumnWidth')
    const replacementSetColumnWidth = vi.spyOn(replacementStore, 'setColumnWidth')
    const replacementSetRowHeight = vi.spyOn(replacementStore, 'setRowHeight')
    const workerHandleRef: { current: WorkerHandle | null } = {
      current: {
        viewportStore: initialStore,
      },
    }
    let sync: ReturnType<typeof useWorkbookSync> | null = null
    const runtimeController = {
      invoke: vi.fn(async (method: string, input?: unknown) => {
        if (method !== 'enqueuePendingMutation') {
          throw new Error(`Unexpected runtime invoke: ${method}`)
        }
        if (!isPendingMutationInput(input)) {
          throw new Error('Expected pending mutation input')
        }
        return {
          ...createPendingMutation(),
          method: input.method,
          args: input.args,
        }
      }),
    }

    function Harness() {
      sync = useWorkbookSync({
        documentId: 'doc-1',
        connectionStateName: 'disconnected',
        connectionStateRef: { current: 'disconnected' },
        runtimeController,
        workerHandleRef,
        zeroRef: {
          current: {
            mutate() {
              throw new Error('Deferred resize test should not attempt remote sync while disconnected')
            },
          },
        },
        reportRuntimeError: vi.fn(),
      })
      return null
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    await act(async () => {
      root.render(<Harness />)
    })
    if (!sync) {
      throw new Error('Expected useWorkbookSync harness to initialize')
    }

    const staleStorePromise = sync.invokeColumnWidthMutation('SheetA', 4, 152, {
      deferLocalApplication: true,
      deferPersistence: true,
    })
    workerHandleRef.current = {
      viewportStore: replacementStore,
    }

    await act(async () => {
      frames.flushNext()
      await Promise.resolve()
    })
    expect(initialSetColumnWidth).not.toHaveBeenCalled()
    expect(replacementSetColumnWidth).not.toHaveBeenCalled()

    await act(async () => {
      frames.flushNext()
      await staleStorePromise
    })
    expect(runtimeController.invoke).toHaveBeenCalledWith('enqueuePendingMutation', {
      method: 'updateColumnMetadata',
      args: ['SheetA', 4, 1, 152, null],
    })

    const stableStorePromise = sync.invokeRowHeightMutation('SheetB', 6, 44, {
      deferLocalApplication: true,
      deferPersistence: true,
    })
    await act(async () => {
      frames.flushNext()
      await Promise.resolve()
    })
    expect(replacementSetRowHeight).toHaveBeenCalledTimes(1)
    expect(replacementSetRowHeight).toHaveBeenCalledWith('SheetB', 6, 44)
    await act(async () => {
      frames.flushNext()
      await stableStorePromise
    })
    expect(runtimeController.invoke).toHaveBeenCalledWith('enqueuePendingMutation', {
      method: 'updateRowMetadata',
      args: ['SheetB', 6, 1, 44, null],
    })

    await act(async () => {
      root.unmount()
    })
    frames.restore()
  })

  it('applies simple cell mutations to the visible viewport before persistence catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const viewportStore = new ProjectedViewportStore()
    const workerHandleRef: { current: WorkerHandle | null } = {
      current: {
        viewportStore,
      },
    }
    let sync: ReturnType<typeof useWorkbookSync> | null = null
    const runtimeController = {
      invoke: vi.fn(async (method: string, input?: unknown) => {
        if (method !== 'enqueuePendingMutation') {
          throw new Error(`Unexpected runtime invoke: ${method}`)
        }
        if (!isPendingMutationInput(input)) {
          throw new Error('Expected pending mutation input')
        }
        return {
          ...createPendingMutation(),
          method: input.method,
          args: input.args,
        }
      }),
    }

    function Harness() {
      sync = useWorkbookSync({
        documentId: 'doc-1',
        connectionStateName: 'disconnected',
        connectionStateRef: { current: 'disconnected' },
        runtimeController,
        workerHandleRef,
        zeroRef: {
          current: {
            mutate() {
              throw new Error('Cell mutation visibility test should not attempt remote sync while disconnected')
            },
          },
        },
        reportRuntimeError: vi.fn(),
      })
      return null
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    await act(async () => {
      root.render(<Harness />)
    })
    if (!sync) {
      throw new Error('Expected useWorkbookSync harness to initialize')
    }

    await act(async () => {
      await sync!.invokeMutation('setCellValue', 'Sheet1', 'D53', 'Month 1')
    })

    expect(viewportStore.getCell('Sheet1', 'D53')).toMatchObject({
      input: 'Month 1',
      value: { tag: ValueTag.String, value: 'Month 1' },
    })

    await act(async () => {
      root.unmount()
    })
  })

  it('does not probe authoritative events for local-only mutations while disconnected', async () => {
    vi.useFakeTimers()
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let sync: ReturnType<typeof useWorkbookSync> | null = null
    const reportRuntimeError = vi.fn()
    const runtimeController = {
      invoke: vi.fn(async (method: string, input?: unknown) => {
        if (method !== 'enqueuePendingMutation') {
          throw new Error(`Unexpected runtime invoke: ${method}`)
        }
        if (!isPendingMutationInput(input)) {
          throw new Error('Expected pending mutation input')
        }
        return {
          ...createPendingMutation(),
          method: input.method,
          args: input.args,
        }
      }),
    }

    function Harness() {
      sync = useWorkbookSync({
        documentId: 'doc-1',
        connectionStateName: 'closed',
        connectionStateRef: { current: 'closed' },
        runtimeController,
        workerHandleRef: { current: null },
        zeroRef: {
          current: {
            mutate() {
              throw new Error('Local-only mutation must not attempt remote sync')
            },
          },
        },
        reportRuntimeError,
      })
      return null
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    await act(async () => {
      root.render(<Harness />)
    })
    if (!sync) {
      throw new Error('Expected useWorkbookSync harness to initialize')
    }

    await act(async () => {
      await sync!.invokeMutation('setCellValue', 'Sheet1', 'D12', 'local-only')
      await vi.advanceTimersByTimeAsync(10_000)
    })

    expect(runtimeController.invoke).toHaveBeenCalledTimes(1)
    expect(runtimeController.invoke).toHaveBeenCalledWith('enqueuePendingMutation', {
      method: 'setCellValue',
      args: ['Sheet1', 'D12', 'local-only'],
    })
    expect(reportRuntimeError).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  it('rejects malformed structural mutation args before persistence', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let sync: ReturnType<typeof useWorkbookSync> | null = null
    const runtimeController = {
      invoke: vi.fn(async () => {
        throw new Error('Invalid mutation args must not reach persistence')
      }),
    }

    function Harness() {
      sync = useWorkbookSync({
        documentId: 'doc-1',
        connectionStateName: 'closed',
        connectionStateRef: { current: 'closed' },
        runtimeController,
        workerHandleRef: { current: null },
        zeroRef: {
          current: {
            mutate() {
              throw new Error('Invalid mutation args must not reach Zero')
            },
          },
        },
        reportRuntimeError: vi.fn(),
      })
      return null
    }

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    await act(async () => {
      root.render(<Harness />)
    })
    if (!sync) {
      throw new Error('Expected useWorkbookSync harness to initialize')
    }

    await expect(sync.invokeMutation('insertRows', 'Sheet1', Number.NaN, 1)).rejects.toThrow('Invalid insertRows args')
    await expect(sync.invokeMutation('updateColumnMetadata', 'Sheet1', 1, 1, Number.POSITIVE_INFINITY, null)).rejects.toThrow(
      'Invalid updateColumnMetadata args',
    )
    await expect(sync.invokeMutation('setFreezePane', 'Sheet1', 1.5, 0)).rejects.toThrow('Invalid setFreezePane args')
    expect(runtimeController.invoke).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })
})

function isPendingMutationInput(
  value: unknown,
): value is { method: PendingWorkbookMutation['method']; args: PendingWorkbookMutation['args'] } {
  return typeof value === 'object' && value !== null && 'method' in value && 'args' in value && Array.isArray(value.args)
}

function installAnimationFrameQueue(): {
  readonly flushNext: () => void
  readonly restore: () => void
} {
  const callbacks: FrameRequestCallback[] = []
  const originalRequestAnimationFrame = window.requestAnimationFrame
  const originalCancelAnimationFrame = window.cancelAnimationFrame
  Object.defineProperty(window, 'requestAnimationFrame', {
    configurable: true,
    writable: true,
    value: vi.fn((callback: FrameRequestCallback) => {
      callbacks.push(callback)
      return callbacks.length
    }),
  })
  Object.defineProperty(window, 'cancelAnimationFrame', {
    configurable: true,
    writable: true,
    value: vi.fn(),
  })
  return {
    flushNext: () => {
      const callback = callbacks.shift()
      if (!callback) {
        throw new Error('Expected a queued animation frame')
      }
      callback(performance.now())
    },
    restore: () => {
      Object.defineProperty(window, 'requestAnimationFrame', {
        configurable: true,
        writable: true,
        value: originalRequestAnimationFrame,
      })
      Object.defineProperty(window, 'cancelAnimationFrame', {
        configurable: true,
        writable: true,
        value: originalCancelAnimationFrame,
      })
    },
  }
}
