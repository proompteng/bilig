// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { useZeroHealthReady } from '../use-zero-health-ready.js'

function TestProbe(props: { connectionStateName: string; runtimeReady: boolean }) {
  const zeroHealthReady = useZeroHealthReady(props)
  return <div data-testid="probe" data-zero-health-ready={String(zeroHealthReady)} />
}

async function flushMicrotasks() {
  await Promise.resolve()
  await Promise.resolve()
}

describe('useZeroHealthReady', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.unstubAllGlobals()
    document.body.innerHTML = ''
  })

  it('stays false while the runtime is not ready', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<TestProbe connectionStateName="connected" runtimeReady={false} />)
    })

    expect(host.querySelector("[data-testid='probe']")?.getAttribute('data-zero-health-ready')).toBe('false')
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('becomes true after a successful keepalive probe', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => ({ ok: true }) satisfies Pick<Response, 'ok'>),
    )
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<TestProbe connectionStateName="connected" runtimeReady />)
      await flushMicrotasks()
    })

    expect(host.querySelector("[data-testid='probe']")?.getAttribute('data-zero-health-ready')).toBe('true')
  })

  it('retries until the keepalive probe succeeds', async () => {
    vi.useFakeTimers()
    const fetchMock = vi.fn<() => Promise<Pick<Response, 'ok'>>>().mockResolvedValueOnce({ ok: false }).mockResolvedValueOnce({ ok: true })
    vi.stubGlobal('fetch', fetchMock)
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<TestProbe connectionStateName="connected" runtimeReady />)
      await flushMicrotasks()
    })

    expect(host.querySelector("[data-testid='probe']")?.getAttribute('data-zero-health-ready')).toBe('false')

    await act(async () => {
      vi.advanceTimersByTime(250)
      await flushMicrotasks()
    })

    expect(fetchMock).toHaveBeenCalledTimes(2)
    expect(host.querySelector("[data-testid='probe']")?.getAttribute('data-zero-health-ready')).toBe('true')
  })
})
