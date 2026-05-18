// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createWorkbookAgentClient } from '../workbook-agent-client.js'

function requestUrl(input: RequestInfo | URL): string {
  return typeof input === 'string' ? input : input instanceof URL ? input.href : input.url
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function requestBody(init: RequestInit | undefined): unknown {
  if (typeof init?.body !== 'string') {
    throw new Error('Expected JSON request body')
  }
  return JSON.parse(init.body) as unknown
}

function createContext(address = 'A1') {
  return {
    selection: {
      sheetName: 'Sheet1',
      address,
    },
    viewport: {
      rowStart: 0,
      rowEnd: 10,
      colStart: 0,
      colEnd: 5,
    },
  }
}

function createSnapshot(overrides: Record<string, unknown> = {}) {
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
    context: createContext(),
    entries: [],
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
  }
}

function createThreadSummary(overrides: Record<string, unknown> = {}) {
  return {
    threadId: 'thr-1',
    scope: 'private',
    ownerUserId: 'alex@example.com',
    updatedAtUnixMs: 100,
    entryCount: 1,
    reviewQueueItemCount: 0,
    latestEntryText: null,
    ...overrides,
  }
}

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('workbook agent client', () => {
  it('builds thread urls and decodes successful responses', async () => {
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads')) {
        return new Response(JSON.stringify([createThreadSummary()]), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      return new Response(JSON.stringify(createSnapshot()), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })
    vi.stubGlobal('fetch', fetchSpy)

    const client = createWorkbookAgentClient('doc-1')
    expect(client.threadEventsUrl('thr/1')).toBe('/v2/documents/doc-1/chat/threads/thr%2F1/events')
    await expect(client.loadThreadSummaries()).resolves.toEqual([expect.objectContaining({ threadId: 'thr-1' })])
    await expect(client.loadThreadSnapshot('thr-1')).resolves.toEqual(expect.objectContaining({ threadId: 'thr-1' }))
  })

  it('surfaces server-provided error messages', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(JSON.stringify({ message: 'Prompt rejected by server' }), {
            status: 422,
            headers: { 'content-type': 'application/json' },
          }),
      ),
    )

    const client = createWorkbookAgentClient('doc-1')
    await expect(client.sendPrompt('thr-1', 'Continue working', createContext())).rejects.toThrow('Prompt rejected by server')
  })

  it('surfaces fallback status when failed responses are not JSON', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response('service unavailable', {
            status: 503,
            headers: { 'content-type': 'text/plain' },
          }),
      ),
    )

    const client = createWorkbookAgentClient('doc-1')
    await expect(client.sendPrompt('thr-1', 'Continue working', createContext())).rejects.toThrow(
      'Workbook agent request failed with status 503',
    )
  })

  it('surfaces malformed JSON for successful responses', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response('{', {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
      ),
    )

    const client = createWorkbookAgentClient('doc-1')
    await expect(client.loadThreadSnapshot('thr-1')).rejects.toThrow('Workbook agent request returned malformed JSON')
  })

  it('surfaces invalid successful response shapes with stable copy', async () => {
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads')) {
        return new Response(JSON.stringify([{ threadId: 'thr-1' }]), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      return new Response(JSON.stringify({ documentId: 'doc-1' }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })
    vi.stubGlobal('fetch', fetchSpy)

    const client = createWorkbookAgentClient('doc-1')
    await expect(client.loadThreadSummaries()).rejects.toThrow('Workbook agent request returned invalid thread summaries')
    await expect(client.loadThreadSnapshot('thr-1')).rejects.toThrow('Workbook agent request returned invalid thread snapshot')
  })

  it('rejects failed context sync responses instead of treating them as synced', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(JSON.stringify({ message: 'Context rejected by server' }), {
            status: 503,
            headers: { 'content-type': 'application/json' },
          }),
      ),
    )

    const client = createWorkbookAgentClient('doc-1')
    await expect(client.syncThreadContext('thr-1', createContext())).rejects.toThrow('Context rejected by server')
  })

  it('skips duplicate semantic context syncs after the context is already current', async () => {
    const fetchSpy = vi.fn(
      async () =>
        new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        }),
    )
    vi.stubGlobal('fetch', fetchSpy)

    const client = createWorkbookAgentClient('doc-1')
    await client.syncThreadContext('thr-1', createContext('A1'))
    await client.syncThreadContext('thr-1', createContext('A1'))
    await client.syncThreadContext('thr-1', {
      ...createContext('A1'),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 7,
        batchId: 11,
        selection: null,
        visibleRange: null,
      },
    })
    await client.syncThreadContext('thr-1', {
      ...createContext('A1'),
      rendered: {
        capturedAtUnixMs: Date.now() + 500,
        capturedRevision: 7,
        batchId: 12,
        selection: null,
        visibleRange: null,
      },
    })

    const contextCalls = fetchSpy.mock.calls.filter(([input]) => requestUrl(input).endsWith('/chat/threads/thr-1/context'))
    expect(contextCalls).toHaveLength(2)
    expect(requestBody(contextCalls[0]?.[1])).toMatchObject({
      context: {
        selection: { sheetName: 'Sheet1', address: 'A1' },
      },
    })
    const firstRequestBody = requestBody(contextCalls[0]?.[1])
    expect(
      isRecord(firstRequestBody) && isRecord(firstRequestBody['context']) ? firstRequestBody['context']['rendered'] : undefined,
    ).toBeUndefined()
    expect(requestBody(contextCalls[1]?.[1])).toMatchObject({
      context: {
        rendered: {
          capturedRevision: 7,
        },
      },
    })
  })

  it('keeps context sync single-flight and posts only the latest pending context', async () => {
    const contextResponses: Array<() => void> = []
    const fetchSpy = vi.fn(
      () =>
        new Promise<Response>((resolve) => {
          contextResponses.push(() => {
            resolve(
              new Response(JSON.stringify({ ok: true }), {
                status: 200,
                headers: { 'content-type': 'application/json' },
              }),
            )
          })
        }),
    )
    vi.stubGlobal('fetch', fetchSpy)

    const client = createWorkbookAgentClient('doc-1')
    const first = client.syncThreadContext('thr-1', createContext('A1'))
    const second = client.syncThreadContext('thr-1', createContext('A2'))
    const third = client.syncThreadContext('thr-1', createContext('A3'))

    expect(fetchSpy).toHaveBeenCalledTimes(1)
    expect(requestBody(fetchSpy.mock.calls[0]?.[1])).toMatchObject({
      context: {
        selection: { sheetName: 'Sheet1', address: 'A1' },
      },
    })

    contextResponses[0]?.()
    await first
    await Promise.resolve()

    expect(fetchSpy).toHaveBeenCalledTimes(2)
    expect(requestBody(fetchSpy.mock.calls[1]?.[1])).toMatchObject({
      context: {
        selection: { sheetName: 'Sheet1', address: 'A3' },
      },
    })

    contextResponses[1]?.()
    await expect(Promise.all([second, third])).resolves.toEqual([undefined, undefined])
  })

  it('does not mark failed context syncs as current', async () => {
    let attempt = 0
    const fetchSpy = vi.fn(async () => {
      attempt += 1
      if (attempt === 1) {
        return new Response(JSON.stringify({ message: 'Context rejected by server' }), {
          status: 503,
          headers: { 'content-type': 'application/json' },
        })
      }
      return new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })
    vi.stubGlobal('fetch', fetchSpy)

    const client = createWorkbookAgentClient('doc-1')
    await expect(client.syncThreadContext('thr-1', createContext('A1'))).rejects.toThrow('Context rejected by server')
    await expect(client.syncThreadContext('thr-1', createContext('A1'))).resolves.toBeUndefined()

    expect(fetchSpy).toHaveBeenCalledTimes(2)
  })
})
