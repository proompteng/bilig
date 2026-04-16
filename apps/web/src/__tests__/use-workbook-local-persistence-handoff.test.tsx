// @vitest-environment jsdom
import { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { useWorkbookLocalPersistenceHandoff } from '../use-workbook-local-persistence-handoff.js'

class MockBroadcastChannel {
  static channels = new Map<string, Set<MockBroadcastChannel>>()

  readonly name: string
  #listeners = new Set<EventListener>()

  constructor(name: string) {
    this.name = name
    const peers = MockBroadcastChannel.channels.get(name) ?? new Set<MockBroadcastChannel>()
    peers.add(this)
    MockBroadcastChannel.channels.set(name, peers)
  }

  postMessage(data: unknown) {
    const peers = MockBroadcastChannel.channels.get(this.name)
    peers?.forEach((peer) => {
      if (peer === this) {
        return
      }
      const event = new MessageEvent<unknown>('message', { data })
      peer.#listeners.forEach((listener) => listener(event))
    })
  }

  addEventListener(type: string, listener: EventListenerOrEventListenerObject) {
    if (type !== 'message' || typeof listener !== 'function') {
      return
    }
    this.#listeners.add(listener)
  }

  removeEventListener(type: string, listener: EventListenerOrEventListenerObject) {
    if (type !== 'message' || typeof listener !== 'function') {
      return
    }
    this.#listeners.delete(listener)
  }

  close() {
    const peers = MockBroadcastChannel.channels.get(this.name)
    peers?.delete(this)
    if (peers?.size === 0) {
      MockBroadcastChannel.channels.delete(this.name)
    }
  }

  static reset() {
    MockBroadcastChannel.channels.clear()
  }
}

function TestProbe(props: {
  documentId: string
  localPersistenceMode: 'persistent' | 'ephemeral' | 'follower'
  onRetry: (persistState: boolean) => void
  testId: string
}) {
  const handoff = useWorkbookLocalPersistenceHandoff({
    documentId: props.documentId,
    localPersistenceMode: props.localPersistenceMode,
    retryRuntime: props.onRetry,
  })

  return (
    <div
      data-testid={props.testId}
      data-pending-request={String(handoff.pendingTransferRequest !== null)}
      data-transfer-requested={String(handoff.transferRequested)}
    >
      <button onClick={handoff.requestPersistenceTransfer} type="button">
        Request
      </button>
      <button onClick={handoff.approvePersistenceTransfer} type="button">
        Approve
      </button>
      <button onClick={handoff.dismissPersistenceTransferRequest} type="button">
        Dismiss
      </button>
    </div>
  )
}

function mountProbe(): {
  host: HTMLDivElement
  root: Root
  render: (nextProps: Parameters<typeof TestProbe>[0]) => Promise<void>
} {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  const render = async (nextProps: Parameters<typeof TestProbe>[0]) => {
    await act(async () => {
      root.render(<TestProbe {...nextProps} />)
    })
  }
  return { host, root, render }
}

describe('useWorkbookLocalPersistenceHandoff', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
    MockBroadcastChannel.reset()
    document.body.innerHTML = ''
  })

  it('coordinates a deliberate writer handoff between tabs', async () => {
    vi.stubGlobal('BroadcastChannel', MockBroadcastChannel)
    const writerRetry = vi.fn()
    const followerRetry = vi.fn()
    const writer = mountProbe()
    const follower = mountProbe()

    await writer.render({
      documentId: 'doc-1',
      localPersistenceMode: 'persistent',
      onRetry: writerRetry,
      testId: 'writer',
    })
    await follower.render({
      documentId: 'doc-1',
      localPersistenceMode: 'follower',
      onRetry: followerRetry,
      testId: 'follower',
    })

    await act(async () => {
      follower.host.querySelector('button')?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(followerRetry).toHaveBeenCalledWith(true)
    expect(writer.host.querySelector("[data-testid='writer']")?.getAttribute('data-pending-request')).toBe('true')

    await act(async () => {
      writer.host.querySelectorAll('button')[1]?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(writerRetry).toHaveBeenCalledWith(false)

    await writer.render({
      documentId: 'doc-1',
      localPersistenceMode: 'ephemeral',
      onRetry: writerRetry,
      testId: 'writer',
    })

    expect(followerRetry).toHaveBeenNthCalledWith(2, true)

    await act(async () => {
      writer.root.unmount()
      follower.root.unmount()
    })
  })

  it('clears a pending writer request when the writer dismisses it', async () => {
    vi.stubGlobal('BroadcastChannel', MockBroadcastChannel)
    const writerRetry = vi.fn()
    const followerRetry = vi.fn()
    const writer = mountProbe()
    const follower = mountProbe()

    await writer.render({
      documentId: 'doc-2',
      localPersistenceMode: 'persistent',
      onRetry: writerRetry,
      testId: 'writer',
    })
    await follower.render({
      documentId: 'doc-2',
      localPersistenceMode: 'follower',
      onRetry: followerRetry,
      testId: 'follower',
    })

    await act(async () => {
      follower.host.querySelector('button')?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(writer.host.querySelector("[data-testid='writer']")?.getAttribute('data-pending-request')).toBe('true')

    await act(async () => {
      writer.host.querySelectorAll('button')[2]?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(writer.host.querySelector("[data-testid='writer']")?.getAttribute('data-pending-request')).toBe('false')

    await act(async () => {
      writer.root.unmount()
      follower.root.unmount()
    })
  })
})
