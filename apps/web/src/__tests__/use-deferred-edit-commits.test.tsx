// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { useDeferredEditCommits } from '../use-deferred-edit-commits.js'

vi.mock('../interaction-idle-scheduler.js', () => ({
  deferInteractionPersistence: () => new Promise<void>(() => {}),
}))

it('keeps edits pending until ordered durable writes finish and flushes each write once', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  let captured: ReturnType<typeof useDeferredEditCommits> | undefined
  function Harness() {
    captured = useDeferredEditCommits()
    return null
  }
  function state() {
    if (!captured) {
      throw new Error('Expected edit commit state')
    }
    return captured
  }
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  let releaseWrite: (() => void) | undefined
  const first = vi.fn(
    () =>
      new Promise<void>((resolve) => {
        releaseWrite = resolve
      }),
  )
  const second = vi.fn(async () => {})
  try {
    await act(async () => root.render(createElement(Harness)))
    await act(async () => {
      state().enqueueDeferredEditCommit(first)
      state().enqueueDeferredEditCommit(second)
    })
    expect(state().isEditCommitPending).toBe(true)
    expect(first).not.toHaveBeenCalled()
    let flush: Promise<void> | undefined
    await act(async () => {
      flush = state().flushPendingEditCommit()
      await Promise.resolve()
    })
    expect(first).toHaveBeenCalledTimes(1)
    expect(second).not.toHaveBeenCalled()
    expect(state().isEditCommitPending).toBe(true)
    await act(async () => {
      releaseWrite?.()
      await flush
    })
    expect(second).toHaveBeenCalledTimes(1)
    expect(state().isEditCommitPending).toBe(false)
    await act(async () => state().flushPendingEditCommit())
    expect(first).toHaveBeenCalledTimes(1)
    expect(second).toHaveBeenCalledTimes(1)
  } finally {
    await act(async () => root.unmount())
    host.remove()
    vi.unstubAllGlobals()
  }
})
