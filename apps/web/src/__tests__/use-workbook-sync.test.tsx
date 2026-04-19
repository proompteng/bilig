// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
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
    document.body.innerHTML = ''
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
})
