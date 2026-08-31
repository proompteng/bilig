// @vitest-environment jsdom
import {
  act,
  AgentHarness,
  agentStorageKey,
  createMockZeroAgentHarness,
  createRoot,
  createSnapshot,
  createThreadSummary,
  describe,
  expect,
  it,
  MockEventSource,
  requestMethod,
  requestUrl,
  vi,
} from './workbook-agent-pane-test-helpers.js'

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((resolvePromise) => {
    resolve = resolvePromise
  })
  return { promise, resolve }
}

function assistantEntry(id: string, turnId: string, text: string) {
  return {
    id,
    kind: 'assistant',
    turnId,
    text,
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
  }
}

describe('workbook agent session races', () => {
  it('does not resurrect a session whose creation was abandoned for a new thread', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const createdSession = deferred<Response>()
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return Promise.resolve(new Response(JSON.stringify([]), { status: 200, headers: { 'content-type': 'application/json' } }))
      }
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'POST') {
        return createdSession.promise
      }
      if (url.endsWith('/turns')) {
        return Promise.resolve(new Response(JSON.stringify(createSnapshot()), { status: 200 }))
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => root.render(<AgentHarness showNewThreadControl />))
    const input = host.querySelector("[data-testid='workbook-agent-input']")
    const submit = host.querySelector("[data-testid='workbook-agent-send']")
    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement) || !(submit instanceof HTMLButtonElement)) {
        throw new Error('Assistant composer not found')
      }
      const valueSetter = Reflect.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set
      if (!valueSetter) {
        throw new Error('Textarea value setter not found')
      }
      Reflect.apply(valueSetter, input, ['Abandoned prompt'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
      submit.click()
    })
    act(() => host.querySelector<HTMLButtonElement>("[data-testid='test-start-new-thread']")?.click())
    await act(async () => {
      const valueSetter = Reflect.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set
      if (!(input instanceof HTMLTextAreaElement) || !(submit instanceof HTMLButtonElement) || !valueSetter) {
        throw new Error('Assistant composer not found after starting a new thread')
      }
      Reflect.apply(valueSetter, input, ['Fresh prompt'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(submit.disabled).toBe(false)

    await act(async () => {
      createdSession.resolve(
        new Response(JSON.stringify(createSnapshot({ threadId: 'thr-abandoned' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        }),
      )
      await createdSession.promise
      await Promise.resolve()
    })

    expect(fetchSpy.mock.calls.filter(([requestInput]) => requestUrl(requestInput).endsWith('/turns'))).toHaveLength(0)
    expect(sessionStorage.getItem(agentStorageKey())).toBeNull()

    await act(async () => root.unmount())
  })

  it('does not let stale stream recovery replace a newly selected thread', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const staleRecovery = deferred<Response>()
    const zero = createMockZeroAgentHarness({
      initialThreadSummaries: [
        createThreadSummary({ threadId: 'thr-1', latestEntryText: 'First thread' }),
        createThreadSummary({ threadId: 'thr-2', latestEntryText: 'Second thread' }),
      ],
      initialWorkflowRuns: [],
    })
    let threadOneSnapshotCalls = 0
    const responseFor = (snapshot: Record<string, unknown>) =>
      new Response(JSON.stringify(createSnapshot(snapshot)), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        threadOneSnapshotCalls += 1
        return threadOneSnapshotCalls === 1
          ? Promise.resolve(
              responseFor({
                threadId: 'thr-1',
                entries: [assistantEntry('first', 'turn-first', 'First thread snapshot')],
              }),
            )
          : staleRecovery.promise
      }
      if (url.endsWith('/chat/threads/thr-2') && requestMethod(init) === 'GET') {
        return Promise.resolve(
          responseFor({
            threadId: 'thr-2',
            entries: [assistantEntry('second', 'turn-second', 'Second thread snapshot')],
          }),
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)
    sessionStorage.setItem(agentStorageKey(), JSON.stringify({ threadId: 'thr-1' }))
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => root.render(<AgentHarness zero={zero.zero} zeroEnabled />))
    const firstSource = MockEventSource.latest
    expect(firstSource?.url).toBe('/v2/documents/doc-1/chat/threads/thr-1/events')

    await act(async () => {
      firstSource?.emitError()
      await vi.waitFor(() => expect(threadOneSnapshotCalls).toBe(2))
    })

    await act(async () => {
      host.querySelector<HTMLButtonElement>("[data-testid='workbook-agent-thread-thr-2']")?.click()
    })
    expect(MockEventSource.latest?.url).toBe('/v2/documents/doc-1/chat/threads/thr-2/events')
    expect(host.textContent).toContain('Second thread snapshot')

    await act(async () => {
      staleRecovery.resolve(
        responseFor({
          threadId: 'thr-1',
          entries: [assistantEntry('stale', 'turn-stale', 'Stale recovery snapshot')],
        }),
      )
      await staleRecovery.promise
      await Promise.resolve()
    })

    expect(MockEventSource.latest?.url).toBe('/v2/documents/doc-1/chat/threads/thr-2/events')
    expect(host.textContent).toContain('Second thread snapshot')
    expect(host.textContent).not.toContain('Stale recovery snapshot')

    await act(async () => root.unmount())
  })
})
