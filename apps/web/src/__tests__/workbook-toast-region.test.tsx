// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import type { Action, ToastT, ToastToDismiss } from 'sonner'
import { toast } from 'sonner'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { WorkbookToastRegion } from '../WorkbookToastRegion.js'

async function flushToasts(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

function findActiveToast(id: string): ToastT | null {
  return toast.getToasts().find((entry: ToastT | ToastToDismiss): entry is ToastT => !('dismiss' in entry) && entry.id === id) ?? null
}

function getToastAction(activeToast: ToastT | null): Action {
  if (!activeToast || !activeToast.action || typeof activeToast.action !== 'object' || !('onClick' in activeToast.action)) {
    throw new Error('Expected toast action object')
  }
  return activeToast.action
}

describe('WorkbookToastRegion', () => {
  afterEach(() => {
    toast.dismiss()
    vi.restoreAllMocks()
    document.body.innerHTML = ''
  })

  it('renders sonner-managed toasts with action and dismiss handlers', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const onAction = vi.fn()
    const onDismiss = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToastRegion
          toasts={[
            {
              id: 'error-1',
              tone: 'error',
              message: 'Remote sync failed.',
              action: {
                label: 'Retry',
                onAction,
              },
              onDismiss,
            },
          ]}
        />,
      )
    })
    await flushToasts()

    const activeToast = findActiveToast('error-1')
    expect(activeToast?.title).toBe('Remote sync failed.')
    expect(activeToast?.classNames).toBeUndefined()
    const retryAction = getToastAction(activeToast)

    await act(async () => {
      Reflect.apply(retryAction.onClick, undefined, [new MouseEvent('click')])
    })

    expect(onAction).toHaveBeenCalledTimes(1)
    expect(typeof activeToast?.onDismiss).toBe('function')

    await act(async () => {
      activeToast?.onDismiss?.(activeToast)
    })
    await flushToasts()

    expect(onDismiss).toHaveBeenCalledTimes(1)

    await act(async () => {
      root.unmount()
    })
  })

  it('does not call dismiss handlers when toasts are removed programmatically', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const onDismiss = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToastRegion
          toasts={[
            {
              id: 'error-2',
              tone: 'error',
              message: 'Remote sync failed.',
              onDismiss,
            },
          ]}
        />,
      )
    })
    await flushToasts()

    await act(async () => {
      root.render(<WorkbookToastRegion toasts={[]} />)
    })
    await flushToasts()

    expect(onDismiss).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  it('uses stock Sonner styling instead of workbook-specific class overrides', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToastRegion
          toasts={[
            {
              id: 'error-3',
              tone: 'error',
              message: 'Failed to capture undo ops',
              onDismiss: () => {},
            },
          ]}
        />,
      )
    })
    await flushToasts()

    const activeToast = findActiveToast('error-3')
    expect(activeToast?.classNames).toBeUndefined()
    expect(activeToast?.unstyled).not.toBe(true)

    await act(async () => {
      root.unmount()
    })
  })
})
