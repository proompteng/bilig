import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { encodeAgentFrame, decodeAgentFrame, type AgentFrame } from '@bilig/agent-api'
import type { AgentFrameContext, DocumentControlService } from '@bilig/runtime-kernel'
import { createHttpWorksheetExecutor, createInProcessWorksheetExecutor, type HttpWorksheetExecutorOptions } from './worksheet-executor.js'

describe('worksheet executor', () => {
  it('posts agent frames to the monolith worksheet endpoint', async () => {
    const requestFrame = {
      kind: 'request',
      request: {
        kind: 'getMetrics',
        id: 'metrics-1',
        sessionId: 'doc-1:replica-1',
      },
    } satisfies AgentFrame
    const responseFrame = {
      kind: 'response',
      response: {
        kind: 'metrics',
        id: 'metrics-1',
        value: {
          service: 'bilig-app',
          documentSessions: 1,
        },
      },
    } satisfies AgentFrame
    const fetchImpl: NonNullable<HttpWorksheetExecutorOptions['fetchImpl']> = vi.fn(async (url: RequestInfo | URL, init?: RequestInit) => {
      const target = typeof url === 'string' ? url : url instanceof URL ? url.toString() : url.url
      expect(target).toBe('https://bilig.proompteng.ai/v2/agent/frames')
      expect(init?.method).toBe('POST')
      expect(init?.headers).toEqual({
        'content-type': 'application/octet-stream',
      })
      if (!(init?.body instanceof Uint8Array)) {
        throw new TypeError('expected binary request body')
      }
      expect(decodeAgentFrame(init.body)).toEqual(requestFrame)
      return new Response(Buffer.from(encodeAgentFrame(responseFrame)), {
        status: 200,
      })
    })

    const executor = createHttpWorksheetExecutor({
      baseUrl: 'https://bilig.proompteng.ai/',
      fetchImpl,
    })

    await expect(executor.execute(requestFrame)).resolves.toEqual(responseFrame)
    expect(fetchImpl).toHaveBeenCalledTimes(1)
  })

  it('forwards agent frames to the in-process worksheet host with configured context', async () => {
    const requestFrame = {
      kind: 'request',
      request: {
        kind: 'openWorkbookSession',
        id: 'open-1',
        documentId: 'doc-1',
        replicaId: 'replica-1',
      },
    } satisfies AgentFrame
    const responseFrame = {
      kind: 'response',
      response: {
        kind: 'ok',
        id: 'open-1',
        sessionId: 'doc-1:replica-1',
      },
    } satisfies AgentFrame
    const handleAgentFrame: DocumentControlService['handleAgentFrame'] = vi.fn((_frame: AgentFrame, _context?: AgentFrameContext) =>
      Effect.succeed(responseFrame),
    )
    const documentService: DocumentControlService = {
      attachBrowser: vi.fn(),
      openBrowserSession: vi.fn(),
      handleSyncFrame: vi.fn(),
      handleAgentFrame,
      getDocumentState: vi.fn(),
      getLatestSnapshot: vi.fn(),
    }
    const executor = createInProcessWorksheetExecutor({
      documentService,
      serverUrl: 'https://bilig.proompteng.ai',
      browserAppBaseUrl: 'https://bilig.proompteng.ai',
    })

    await expect(executor.execute(requestFrame)).resolves.toEqual(responseFrame)
    expect(handleAgentFrame).toHaveBeenCalledWith(requestFrame, {
      serverUrl: 'https://bilig.proompteng.ai',
      browserAppBaseUrl: 'https://bilig.proompteng.ai',
    })
  })
})
