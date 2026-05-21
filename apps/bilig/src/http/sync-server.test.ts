import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookAgentThreadSnapshot } from '@bilig/contracts'
import { toWorkbookAgentReviewQueueItem, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { Effect } from 'effect'
import type { DocumentControlService } from '@bilig/runtime-kernel'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookAgentService } from '../codex-app/workbook-agent-service.js'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import { createSyncServer } from './sync-server.js'

type TestServer = Awaited<ReturnType<typeof startHttpServer>>

async function startHttpServer(handler: (request: IncomingMessage, response: ServerResponse) => void) {
  const server = createServer(handler)
  await new Promise<void>((resolve, reject) => {
    server.listen(0, '127.0.0.1', () => resolve())
    server.once('error', reject)
  })
  const address = server.address()
  if (!address || typeof address === 'string') {
    throw new Error('Expected TCP test server address')
  }
  return {
    server,
    origin: `http://127.0.0.1:${String(address.port)}`,
    async close() {
      await new Promise<void>((resolve, reject) => {
        server.close((error) => {
          if (error) {
            reject(error)
            return
          }
          resolve()
        })
      })
    },
  }
}

const upstreamServers: TestServer[] = []

afterEach(async () => {
  delete process.env['BILIG_ZERO_PROXY_UPSTREAM']
  delete process.env['BILIG_PERSIST_STATE']
  delete process.env['BILIG_REMOTE_MCP_ALLOWED_ORIGINS']
  await Promise.all(upstreamServers.splice(0).map((server) => server.close()))
})

function createZeroSyncStub(overrides: Partial<ZeroSyncService> = {}): ZeroSyncService {
  return {
    enabled: true,
    async initialize() {},
    async close() {},
    async handleQuery() {
      throw new Error('not used')
    },
    async handleMutate() {
      throw new Error('not used')
    },
    async inspectWorkbook<T>(_documentId: string, _task: (runtime: never) => T | Promise<T>) {
      throw new Error('not used')
    },
    async applyServerMutator() {
      throw new Error('not used')
    },
    async applyAgentCommandBundle() {
      throw new Error('not used')
    },
    async listWorkbookChanges() {
      return []
    },
    async listWorkbookAgentRuns() {
      return []
    },
    async listWorkbookAgentThreadRuns() {
      return []
    },
    async appendWorkbookAgentRun() {
      throw new Error('not used')
    },
    async listWorkbookAgentThreadSummaries() {
      return []
    },
    async loadWorkbookAgentThreadState() {
      return null
    },
    async saveWorkbookAgentThreadState() {
      throw new Error('not used')
    },
    async listWorkbookThreadWorkflowRuns() {
      return []
    },
    async upsertWorkbookWorkflowRun() {
      throw new Error('not used')
    },
    async getWorkbookHeadRevision() {
      return 1
    },
    async loadAuthoritativeEvents() {
      throw new Error('not used')
    },
    ...overrides,
  }
}

function createWorkbookAgentServiceStub(overrides: Partial<WorkbookAgentService> = {}): WorkbookAgentService {
  return {
    enabled: true,
    async createSession() {
      throw new Error('not used')
    },
    async updateContext() {
      throw new Error('not used')
    },
    async startTurn() {
      throw new Error('not used')
    },
    async startWorkflow() {
      throw new Error('not used')
    },
    async cancelWorkflow() {
      throw new Error('not used')
    },
    async interruptTurn() {
      throw new Error('not used')
    },
    async applyReviewItem() {
      throw new Error('not used')
    },
    async reviewReviewItem() {
      throw new Error('not used')
    },
    async dismissReviewItem() {
      throw new Error('not used')
    },
    async replayExecutionRecord() {
      throw new Error('not used')
    },
    async listThreads() {
      return []
    },
    getObservabilitySnapshot() {
      return {
        enabled: true,
        generatedAtUnixMs: 1,
        featureFlags: {
          sharedThreadsEnabled: true,
          workflowRunnerEnabled: true,
          autoApplyLowRiskEnabled: true,
          formulaWorkflowFamilyEnabled: true,
          formattingWorkflowFamilyEnabled: true,
          importWorkflowFamilyEnabled: true,
          rollupWorkflowFamilyEnabled: true,
          structuralWorkflowFamilyEnabled: true,
          allowlistedUserCount: 0,
          allowlistedDocumentCount: 0,
        },
        sessions: {
          sessionCount: 0,
          subscriberThreadCount: 0,
          subscriberCount: 0,
          activeTurnCount: 0,
          runningWorkflowCount: 0,
          reviewQueueSessionCount: 0,
          sharedPendingReviewCount: 0,
        },
        pool: {
          slotCount: 0,
          boundThreadCount: 0,
          activeTurnCount: 0,
          queuedTurnCount: 0,
          maxClients: 0,
          maxConcurrentTurnsPerClient: 0,
          maxQueuedTurnsPerClient: 0,
        },
        counters: {
          turnBackpressureCount: 0,
          workflowStartedCount: 0,
          workflowCompletedCount: 0,
          workflowFailedCount: 0,
          workflowCancelledCount: 0,
          sharedReviewApprovedCount: 0,
          sharedReviewRejectedCount: 0,
          sharedRecommendationApprovedCount: 0,
          sharedRecommendationRejectedCount: 0,
        },
      }
    },
    getSnapshot() {
      throw new Error('not used')
    },
    subscribe() {
      return () => {}
    },
    async close() {},
    ...overrides,
  }
}

function createDocumentServiceStub(overrides: Partial<DocumentControlService> = {}): DocumentControlService {
  return {
    attachBrowser() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    openBrowserSession() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    handleSyncFrame() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    handleAgentFrame() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    getDocumentState() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    getLatestSnapshot() {
      return Effect.succeed(null)
    },
    ...overrides,
  }
}

function createAgentSessionSnapshot(overrides: Partial<WorkbookAgentThreadSnapshot> = {}): WorkbookAgentThreadSnapshot {
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    executionPolicy: 'autoApplyAll',
    scope: 'private',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
    context: {
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    },
    entries: [],
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
  }
}

function createReviewQueueItem(bundle: WorkbookAgentCommandBundle) {
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: bundle.sharedReview ? 'ownerReview' : 'manual',
    ...(bundle.sharedReview ? { sharedReview: bundle.sharedReview } : {}),
  })
}

function readMcpToolNames(responseBody: unknown): string[] {
  const tools = readMcpTools(responseBody)
  return tools.map((tool) => {
    if (!isRecord(tool) || typeof tool['name'] !== 'string') {
      throw new Error(`Expected MCP tool definition, received ${JSON.stringify(tool)}`)
    }
    return tool['name']
  })
}

function readMcpToolAnnotations(responseBody: unknown, toolName: string): Record<string, unknown> {
  const tool = readMcpTools(responseBody).find((candidate) => isRecord(candidate) && candidate['name'] === toolName)
  if (!isRecord(tool) || !isRecord(tool['annotations'])) {
    throw new Error(`Expected MCP tool annotations for ${toolName}, received ${JSON.stringify(responseBody)}`)
  }
  return tool['annotations']
}

function readMcpTools(responseBody: unknown): unknown[] {
  const result = isRecord(responseBody) ? responseBody['result'] : undefined
  const tools = isRecord(result) ? result['tools'] : undefined
  if (!Array.isArray(tools)) {
    throw new Error(`Expected tools/list response, received ${JSON.stringify(responseBody)}`)
  }
  return tools
}

function readMcpServerCardToolNames(card: unknown): string[] {
  const tools = isRecord(card) ? card['tools'] : undefined
  if (!Array.isArray(tools)) {
    throw new Error(`Expected MCP server-card tools, received ${JSON.stringify(card)}`)
  }
  return tools.map((tool) => {
    if (!isRecord(tool) || typeof tool['name'] !== 'string') {
      throw new Error(`Expected MCP server-card tool, received ${JSON.stringify(tool)}`)
    }
    return tool['name']
  })
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

describe('sync-server zero keepalive', () => {
  it('proxies a healthy keepalive response without using the generic zero proxy route', async () => {
    const upstream = await startHttpServer((request, response) => {
      expect(request.url).toBe('/keepalive')
      response.writeHead(200, { 'content-type': 'text/plain; charset=utf-8' })
      response.end('ok')
    })
    upstreamServers.push(upstream)
    process.env['BILIG_ZERO_PROXY_UPSTREAM'] = upstream.origin

    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/zero/keepalive',
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.body).toBe('ok')
    } finally {
      await app.close()
    }
  })

  it('returns 503 when the upstream resets the keepalive connection', async () => {
    const upstream = await startHttpServer((request) => {
      expect(request.url).toBe('/keepalive')
      request.socket.destroy()
    })
    upstreamServers.push(upstream)
    process.env['BILIG_ZERO_PROXY_UPSTREAM'] = upstream.origin

    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/zero/keepalive',
      })

      expect(response.statusCode).toBe(503)
      expect(response.json()).toEqual({
        error: 'ZERO_CACHE_UNAVAILABLE',
        message: 'Zero cache keepalive probe failed',
        retryable: true,
      })
    } finally {
      await app.close()
    }
  })
})

describe('sync-server cross-origin isolation', () => {
  it('serves runtime responses with isolation and CSP headers required for the workbook browser runtime', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/runtime-config.json',
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['cross-origin-opener-policy']).toBe('same-origin')
      expect(response.headers['cross-origin-embedder-policy']).toBe('require-corp')
      expect(response.headers['origin-agent-cluster']).toBe('?1')
      expect(response.headers['content-security-policy']).toEqual(
        expect.stringContaining("default-src 'self'; object-src 'none'; base-uri 'none'"),
      )
      expect(response.headers['content-security-policy']).toEqual(expect.stringContaining("script-src 'self' 'wasm-unsafe-eval'"))
      expect(response.headers['content-security-policy']).toEqual(expect.stringContaining("worker-src 'self' blob:"))
    } finally {
      await app.close()
    }
  })
})

describe('sync-server remote WorkPaper MCP', () => {
  it('serves same-origin MCP server cards for hosted directory scanners', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const responses = await Promise.all(
        ['/.well-known/mcp/server-card.json', '/.well-known/mcp.json', '/.well-known/mcp-server-card.json'].map((url) =>
          app.inject({
            method: 'GET',
            url,
          }),
        ),
      )

      for (const response of responses) {
        expect(response.statusCode).toBe(200)
        expect(response.headers['content-type']).toContain('application/json')
        expect(response.headers['access-control-allow-origin']).toBe('*')
        const card = response.json()
        expect(card).toMatchObject({
          protocolVersion: '2025-11-25',
          serverInfo: {
            name: 'Bilig WorkPaper Remote Demo',
            version: expect.any(String),
          },
          authentication: {
            required: false,
          },
          transport: {
            type: 'streamable-http',
            url: 'https://bilig.proompteng.ai/mcp',
            stateless: true,
          },
        })
        expect(readMcpServerCardToolNames(card)).toEqual([
          'list_sheets',
          'read_range',
          'read_cell',
          'set_cell_contents',
          'get_cell_display_value',
          'export_workpaper_document',
          'validate_formula',
        ])
        expect(card.resources).toEqual(
          expect.arrayContaining([
            expect.objectContaining({
              uri: 'bilig://workpaper/agent-handoff',
            }),
          ]),
        )
        expect(card.prompts).toEqual(
          expect.arrayContaining([
            expect.objectContaining({
              name: 'edit_and_verify_workpaper',
            }),
          ]),
        )
      }
    } finally {
      await app.close()
    }
  })

  it('initializes a stateless Streamable HTTP MCP endpoint with WorkPaper tools', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/mcp',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2025-11-25',
          origin: 'https://claude.ai',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          id: 1,
          method: 'initialize',
          params: {
            protocolVersion: '2025-11-25',
            capabilities: {},
            clientInfo: {
              name: 'test-client',
              version: '1.0.0',
            },
          },
        }),
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['content-type']).toContain('application/json')
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.headers['mcp-protocol-version']).toBe('2025-11-25')
      expect(response.headers['access-control-allow-origin']).toBe('https://claude.ai')
      expect(response.headers['mcp-session-id']).toBeUndefined()
      expect(response.headers['set-cookie']).toBeUndefined()
      expect(response.json()).toMatchObject({
        jsonrpc: '2.0',
        id: 1,
        result: {
          protocolVersion: '2025-11-25',
          capabilities: {
            tools: {
              listChanged: false,
            },
            resources: {
              listChanged: false,
              subscribe: false,
            },
            prompts: {
              listChanged: false,
            },
          },
          serverInfo: {
            name: 'bilig-workpaper-remote-demo',
            title: 'Bilig WorkPaper Remote Demo',
          },
        },
      })
    } finally {
      await app.close()
    }
  })

  it('allows ChatGPT Apps to call the hosted WorkPaper MCP endpoint', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const preflight = await app.inject({
        method: 'OPTIONS',
        url: '/mcp',
        headers: {
          origin: 'https://chatgpt.com',
          'access-control-request-method': 'POST',
          'access-control-request-headers': 'accept, content-type, mcp-protocol-version',
        },
      })

      expect(preflight.statusCode).toBe(204)
      expect(preflight.headers['access-control-allow-origin']).toBe('https://chatgpt.com')
      expect(preflight.headers['vary']).toBe('Origin')

      const response = await app.inject({
        method: 'POST',
        url: '/mcp',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2025-11-25',
          origin: 'https://chatgpt.com',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          id: 'tools',
          method: 'tools/list',
        }),
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['access-control-allow-origin']).toBe('https://chatgpt.com')
      expect(readMcpToolNames(response.json())).toContain('set_cell_contents')
    } finally {
      await app.close()
    }
  })

  it('lists directory-friendly WorkPaper MCP tools over the HTTP alias', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/mcp/workpaper',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2025-11-25',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          id: 'tools',
          method: 'tools/list',
        }),
      })

      expect(response.statusCode).toBe(200)
      const body = response.json()
      expect(readMcpToolNames(body)).toEqual([
        'list_sheets',
        'read_range',
        'read_cell',
        'set_cell_contents',
        'get_cell_display_value',
        'export_workpaper_document',
        'validate_formula',
      ])
      expect(readMcpToolAnnotations(body, 'set_cell_contents')).toMatchObject({
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false,
      })
    } finally {
      await app.close()
    }
  })

  it('edits the demo WorkPaper in request-local memory and returns readback proof', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/mcp',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2025-11-25',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          id: 2,
          method: 'tools/call',
          params: {
            name: 'set_cell_contents',
            arguments: {
              sheetName: 'Inputs',
              address: 'B3',
              value: 0.4,
            },
          },
        }),
      })

      expect(response.statusCode).toBe(200)
      expect(response.json()).toMatchObject({
        jsonrpc: '2.0',
        id: 2,
        result: {
          isError: false,
          structuredContent: {
            editedCell: 'Inputs!B3',
            before: {
              serialized: 0.25,
            },
            after: {
              serialized: 0.4,
            },
            persistence: {
              persisted: false,
            },
            checks: {
              persisted: false,
              restoredMatchesAfter: true,
            },
          },
        },
      })
    } finally {
      await app.close()
    }
  })

  it('accepts notifications with HTTP 202 and no body', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/mcp',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2025-11-25',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          method: 'notifications/initialized',
        }),
      })

      expect(response.statusCode).toBe(202)
      expect(response.body).toBe('')
    } finally {
      await app.close()
    }
  })

  it('handles ping requests and JSON-RPC response messages over POST', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const ping = await app.inject({
        method: 'POST',
        url: '/mcp',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2025-11-25',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          id: 'ping-1',
          method: 'ping',
        }),
      })

      expect(ping.statusCode).toBe(200)
      expect(ping.json()).toEqual({
        jsonrpc: '2.0',
        id: 'ping-1',
        result: {},
      })

      const responseMessage = await app.inject({
        method: 'POST',
        url: '/mcp',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2025-11-25',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          id: 'client-response',
          result: {},
        }),
      })

      expect(responseMessage.statusCode).toBe(202)
      expect(responseMessage.body).toBe('')
    } finally {
      await app.close()
    }
  })

  it('rejects invalid Streamable HTTP headers without minting a session cookie', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const invalidOrigin = await app.inject({
        method: 'POST',
        url: '/mcp',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2025-11-25',
          origin: 'https://attacker.example',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          id: 1,
          method: 'tools/list',
        }),
      })

      expect(invalidOrigin.statusCode).toBe(403)
      expect(invalidOrigin.headers['set-cookie']).toBeUndefined()
      expect(invalidOrigin.json()).toMatchObject({
        jsonrpc: '2.0',
        id: null,
        error: {
          message: 'Forbidden Origin header',
        },
      })

      const invalidProtocol = await app.inject({
        method: 'POST',
        url: '/mcp',
        headers: {
          accept: 'application/json, text/event-stream',
          'content-type': 'application/json',
          'mcp-protocol-version': '2024-11-05',
        },
        payload: JSON.stringify({
          jsonrpc: '2.0',
          id: 1,
          method: 'tools/list',
        }),
      })

      expect(invalidProtocol.statusCode).toBe(400)
      expect(invalidProtocol.json()).toMatchObject({
        jsonrpc: '2.0',
        id: null,
        error: {
          message: expect.stringContaining('Unsupported MCP-Protocol-Version'),
        },
      })
    } finally {
      await app.close()
    }
  })

  it('returns 405 for GET because the stateless endpoint does not offer SSE', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/mcp',
        headers: {
          accept: 'text/event-stream',
        },
      })

      expect(response.statusCode).toBe(405)
      expect(response.headers['allow']).toBe('POST, GET, DELETE, OPTIONS')
      expect(response.json()).toMatchObject({
        jsonrpc: '2.0',
        id: null,
        error: {
          message: 'Method not allowed; this stateless endpoint returns JSON over POST',
        },
      })
    } finally {
      await app.close()
    }
  })
})

describe('sync-server runtime config', () => {
  it('resolves explicit persist-state controls for the browser runtime', async () => {
    process.env['BILIG_PERSIST_STATE'] = 'false'
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/runtime-config.json',
      })

      expect(response.statusCode).toBe(200)
      expect(response.json()).toMatchObject({
        persistState: false,
      })
    } finally {
      await app.close()
    }
  })

  it('rejects ambiguous persist-state controls before serving runtime config', () => {
    process.env['BILIG_PERSIST_STATE'] = ' FALSE '

    expect(() => createSyncServer({ logger: false })).toThrow(
      'BILIG_PERSIST_STATE must be "1", "true", "0", or "false" when set, got  FALSE ',
    )
  })
})

describe('sync-server snapshots', () => {
  it('returns 204 when no latest snapshot exists', async () => {
    const { app } = createSyncServer({
      logger: false,
      documentService: createDocumentServiceStub({
        getLatestSnapshot(documentId: string) {
          expect(documentId).toBe('doc-1')
          return Effect.succeed(null)
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/doc-1/snapshot/latest',
      })

      expect(response.statusCode).toBe(204)
      expect(response.body).toBe('')
      expect(response.headers['content-type']).toBeUndefined()
    } finally {
      await app.close()
    }
  })

  it('falls back to the authoritative zero workbook snapshot when the live session has no snapshot', async () => {
    const calls: string[] = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'doc-1',
      },
      sheets: [
        {
          id: 1,
          name: 'Prepaid Template',
          order: 0,
          cells: [
            {
              address: 'C6',
              value: 42,
            },
          ],
        },
      ],
    }
    const { app } = createSyncServer({
      logger: false,
      documentService: createDocumentServiceStub({
        getLatestSnapshot(documentId: string) {
          expect(documentId).toBe('doc-1')
          return Effect.succeed(null)
        },
      }),
      zeroSyncService: createZeroSyncStub({
        async ensureWorkbookDocument(documentId, ownerUserId) {
          calls.push(`ensure:${documentId}:${ownerUserId ?? ''}`)
        },
        async loadLatestWorkbookSnapshot(documentId) {
          calls.push(`load:${documentId}`)
          expect(documentId).toBe('doc-1')
          return {
            revision: 451,
            calculatedRevision: 449,
            snapshot,
          }
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/doc-1/snapshot/latest',
        headers: {
          'x-bilig-user-id': 'owner-1',
        },
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['x-bilig-snapshot-cursor']).toBe('451')
      expect(response.headers['x-bilig-calculated-cursor']).toBe('449')
      expect(response.headers['content-type']).toContain('application/vnd.bilig.workbook+json')
      expect(JSON.parse(response.body)).toEqual(snapshot)
      expect(calls).toEqual(['ensure:doc-1:owner-1', 'load:doc-1'])
    } finally {
      await app.close()
    }
  })
})

describe('sync-server authoritative events', () => {
  it('returns authoritative workbook events from the zero sync service', async () => {
    const calls: string[] = []
    const { app } = createSyncServer({
      logger: false,
      zeroSyncService: createZeroSyncStub({
        async ensureWorkbookDocument(documentId, ownerUserId) {
          calls.push(`ensure:${documentId}:${ownerUserId ?? ''}`)
        },
        async loadAuthoritativeEvents(documentId, afterRevision) {
          calls.push(`events:${documentId}:${String(afterRevision)}`)
          expect(documentId).toBe('doc-1')
          expect(afterRevision).toBe(4)
          return {
            afterRevision,
            headRevision: 6,
            calculatedRevision: 6,
            events: [
              {
                revision: 5,
                clientMutationId: 'doc-1:pending:5',
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A1',
                  value: 42,
                },
              },
              {
                revision: 6,
                clientMutationId: 'doc-1:pending:6',
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'B1',
                  value: 84,
                },
              },
            ],
          }
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/doc-1/events?afterRevision=4',
        headers: {
          'x-bilig-user-id': 'owner-1',
        },
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.json()).toEqual({
        afterRevision: 4,
        headRevision: 6,
        calculatedRevision: 6,
        events: [
          {
            revision: 5,
            clientMutationId: 'doc-1:pending:5',
            payload: {
              kind: 'setCellValue',
              sheetName: 'Sheet1',
              address: 'A1',
              value: 42,
            },
          },
          {
            revision: 6,
            clientMutationId: 'doc-1:pending:6',
            payload: {
              kind: 'setCellValue',
              sheetName: 'Sheet1',
              address: 'B1',
              value: 84,
            },
          },
        ],
      })
      expect(calls).toEqual(['ensure:doc-1:owner-1', 'events:doc-1:4'])
    } finally {
      await app.close()
    }
  })

  it('rejects invalid afterRevision values', async () => {
    const { app } = createSyncServer({
      logger: false,
      zeroSyncService: createZeroSyncStub({
        async loadAuthoritativeEvents() {
          throw new Error('not used')
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/doc-1/events?afterRevision=nope',
      })

      expect(response.statusCode).toBe(400)
      expect(response.json()).toEqual({
        error: 'INVALID_AFTER_REVISION',
        message: 'afterRevision must be a non-negative integer',
        retryable: false,
      })
    } finally {
      await app.close()
    }
  })
})

describe('sync-server workbook agent', () => {
  it('lists durable workbook chat threads through the public route', async () => {
    const listThreads = vi.fn(async () => [
      {
        threadId: 'thr-2',
        scope: 'shared' as const,
        ownerUserId: 'alex@example.com',
        updatedAtUnixMs: 200,
        entryCount: 3,
        reviewQueueItemCount: 0,
        latestEntryText: 'Applied shared cleanup at revision r7',
      },
      {
        threadId: 'thr-1',
        scope: 'private' as const,
        ownerUserId: 'alex@example.com',
        updatedAtUnixMs: 100,
        entryCount: 1,
        reviewQueueItemCount: 1,
        latestEntryText: 'Review item queued',
      },
    ])

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        listThreads,
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/doc-1/chat/threads',
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['cache-control']).toBe('no-store')
      expect(listThreads).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
        }),
      )
      expect(response.json()).toEqual([
        {
          threadId: 'thr-2',
          scope: 'shared',
          ownerUserId: 'alex@example.com',
          updatedAtUnixMs: 200,
          entryCount: 3,
          reviewQueueItemCount: 0,
          latestEntryText: 'Applied shared cleanup at revision r7',
        },
        {
          threadId: 'thr-1',
          scope: 'private',
          ownerUserId: 'alex@example.com',
          updatedAtUnixMs: 100,
          entryCount: 1,
          reviewQueueItemCount: 1,
          latestEntryText: 'Review item queued',
        },
      ])
    } finally {
      await app.close()
    }
  })

  it('creates or resumes workbook chat threads through the public route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-shared',
        scope: 'shared',
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads',
        payload: {
          threadId: 'thr-shared',
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B2',
            },
            viewport: {
              rowStart: 0,
              rowEnd: 10,
              colStart: 0,
              colEnd: 5,
            },
          },
        },
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: expect.objectContaining({
            threadId: 'thr-shared',
          }),
        }),
      )
      expect(response.json()).toEqual(
        expect.objectContaining({
          threadId: 'thr-shared',
          scope: 'shared',
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('loads workbook chat thread snapshots through a thread-specific route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-shared',
        scope: 'shared',
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/doc-1/chat/threads/thr-shared',
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-shared',
          },
        }),
      )
      expect(response.json()).toEqual(
        expect.objectContaining({
          threadId: 'thr-shared',
          scope: 'shared',
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('starts workbook chat turns through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const startTurn = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        status: 'inProgress',
        activeTurnId: 'turn-1',
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        startTurn,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/turns',
        payload: {
          prompt: 'Summarize this thread',
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'A1',
            },
            viewport: {
              rowStart: 0,
              rowEnd: 10,
              colStart: 0,
              colEnd: 5,
            },
          },
        },
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: expect.objectContaining({
            threadId: 'thr-2',
          }),
        }),
      )
      expect(startTurn).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          body: expect.objectContaining({
            prompt: 'Summarize this thread',
          }),
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('starts workbook chat workflows through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const startWorkflow = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        workflowRuns: [
          {
            runId: 'wf-2',
            threadId: 'thr-2',
            startedByUserId: 'alex@example.com',
            workflowTemplate: 'describeRecentChanges',
            title: 'Describe Recent Changes',
            summary: 'Summarized 3 recent workbook changes.',
            status: 'completed' as const,
            createdAtUnixMs: 1,
            updatedAtUnixMs: 3,
            completedAtUnixMs: 3,
            errorMessage: null,
            steps: [
              {
                stepId: 'load-revisions',
                label: 'Load durable revisions',
                status: 'completed' as const,
                summary: 'Loaded 3 durable workbook revisions.',
                updatedAtUnixMs: 2,
              },
              {
                stepId: 'draft-change-report',
                label: 'Draft change report',
                status: 'completed' as const,
                summary: 'Prepared the durable recent change report for the thread.',
                updatedAtUnixMs: 3,
              },
            ],
            artifact: {
              kind: 'markdown' as const,
              title: 'Recent Changes',
              text: '## Recent Changes',
            },
          },
        ],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        startWorkflow,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/workflows',
        payload: {
          workflowTemplate: 'describeRecentChanges',
        },
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(startWorkflow).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          body: {
            workflowTemplate: 'describeRecentChanges',
          },
        }),
      )
      expect(response.json()).toEqual(
        expect.objectContaining({
          threadId: 'thr-2',
          workflowRuns: [
            expect.objectContaining({
              runId: 'wf-2',
              workflowTemplate: 'describeRecentChanges',
            }),
          ],
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('cancels workbook chat workflows through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const cancelWorkflow = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        workflowRuns: [
          {
            runId: 'wf-running-2',
            threadId: 'thr-2',
            startedByUserId: 'alex@example.com',
            workflowTemplate: 'describeRecentChanges',
            title: 'Describe Recent Changes',
            summary: 'Cancelled workflow: Describe Recent Changes',
            status: 'cancelled' as const,
            createdAtUnixMs: 1,
            updatedAtUnixMs: 4,
            completedAtUnixMs: 4,
            errorMessage: 'Cancelled by alex@example.com.',
            steps: [
              {
                stepId: 'load-revisions',
                label: 'Load durable revisions',
                status: 'cancelled' as const,
                summary: 'Workflow cancelled before this step completed.',
                updatedAtUnixMs: 4,
              },
            ],
            artifact: null,
          },
        ],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        cancelWorkflow,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/workflows/wf-running-2/cancel',
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(cancelWorkflow).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          runId: 'wf-running-2',
        }),
      )
      expect(response.json()).toEqual(
        expect.objectContaining({
          threadId: 'thr-2',
          workflowRuns: [
            expect.objectContaining({
              runId: 'wf-running-2',
              status: 'cancelled',
            }),
          ],
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('passes query input through workbook search workflows', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-search',
      }),
    )
    const startWorkflow = vi.fn(async () =>
      createAgentSessionSnapshot({
        workflowRuns: [
          {
            runId: 'wf-search-1',
            threadId: 'thr-1',
            startedByUserId: 'alex@example.com',
            workflowTemplate: 'searchWorkbookQuery',
            title: 'Search Workbook',
            summary: 'Found 1 workbook match for "revenue".',
            status: 'completed' as const,
            createdAtUnixMs: 1,
            updatedAtUnixMs: 2,
            completedAtUnixMs: 2,
            errorMessage: null,
            steps: [
              {
                stepId: 'search-workbook',
                label: 'Search workbook',
                status: 'completed' as const,
                summary: 'Searched workbook sheets, formulas, values, and addresses for "revenue" and found 1 match.',
                updatedAtUnixMs: 1,
              },
              {
                stepId: 'draft-search-report',
                label: 'Draft search report',
                status: 'completed' as const,
                summary: 'Prepared the durable workbook search report for the thread.',
                updatedAtUnixMs: 2,
              },
            ],
            artifact: {
              kind: 'markdown' as const,
              title: 'Workbook Search',
              text: '## Workbook Search',
            },
          },
        ],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        startWorkflow,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-search/workflows',
        payload: {
          workflowTemplate: 'searchWorkbookQuery',
          query: 'revenue',
          limit: 5,
        },
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-search',
          },
        }),
      )
      expect(startWorkflow).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-search',
          body: {
            workflowTemplate: 'searchWorkbookQuery',
            query: 'revenue',
            limit: 5,
          },
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('reviews workbook agent bundles through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const reviewReviewItem = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        reviewQueueItems: [
          createReviewQueueItem({
            id: 'bundle-1',
            documentId: 'doc-1',
            threadId: 'thr-2',
            turnId: 'turn-1',
            goalText: 'Normalize shared workbook',
            summary: 'Normalize shared workbook',
            scope: 'workbook',
            riskClass: 'high',
            baseRevision: 4,
            createdAtUnixMs: 10,
            context: null,
            commands: [],
            affectedRanges: [],
            estimatedAffectedCells: 0,
            sharedReview: {
              ownerUserId: 'alex@example.com',
              status: 'approved',
              decidedByUserId: 'alex@example.com',
              decidedAtUnixMs: 12,
              recommendations: [],
            },
          }),
        ],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        reviewReviewItem,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/review-items/bundle-1/review',
        payload: {
          decision: 'approved',
        },
      })

      expect(response.statusCode).toBe(200)
      expect(reviewReviewItem).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          reviewItemId: 'bundle-1',
          body: {
            decision: 'approved',
          },
        }),
      )
      expect(response.json()).toEqual(
        expect.objectContaining({
          reviewQueueItems: [
            expect.objectContaining({
              reviewMode: 'ownerReview',
              status: 'approved',
              decidedByUserId: 'alex@example.com',
            }),
          ],
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('updates workbook agent context through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const updateContext = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        updateContext,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/context',
        payload: {
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B2',
            },
            viewport: {
              rowStart: 1,
              rowEnd: 11,
              colStart: 1,
              colEnd: 6,
            },
          },
        },
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(updateContext).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          body: {
            context: {
              selection: {
                sheetName: 'Sheet1',
                address: 'B2',
              },
              viewport: {
                rowStart: 1,
                rowEnd: 11,
                colStart: 1,
                colEnd: 6,
              },
            },
          },
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('interrupts workbook agent turns through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const interruptTurn = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        status: 'idle',
        activeTurnId: null,
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        interruptTurn,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/interrupt',
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(interruptTurn).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('applies staged workbook bundles through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const applyReviewItem = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        reviewQueueItems: [],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        applyReviewItem,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/review-items/bundle-1/apply',
        payload: {
          commandIndexes: [1],
        },
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(applyReviewItem).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          reviewItemId: 'bundle-1',
          appliedBy: 'user',
          commandIndexes: [1],
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('returns a structured conflict envelope when agent apply rejects a stale preview', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-stale',
      }),
    )
    const applyReviewItem = vi.fn(async () => {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_PREVIEW_STALE',
        message: 'Workbook changed after preview. Replay the plan to stage a fresh review item.',
        statusCode: 409,
        retryable: true,
      })
    })

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        applyReviewItem,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-stale/review-items/bundle-1/apply',
        payload: {},
      })

      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-stale',
          },
        }),
      )
      expect(response.statusCode).toBe(409)
      expect(response.json()).toEqual(
        expect.objectContaining({
          error: 'WORKBOOK_AGENT_PREVIEW_STALE',
          message: 'Workbook changed after preview. Replay the plan to stage a fresh review item.',
          retryable: true,
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('dismisses staged workbook bundles through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const dismissReviewItem = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        dismissReviewItem,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/review-items/bundle-1/dismiss',
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(dismissReviewItem).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          reviewItemId: 'bundle-1',
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('replays prior execution records through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const replayExecutionRecord = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        reviewQueueItems: [
          createReviewQueueItem({
            id: 'bundle-replay-1',
            documentId: 'doc-1',
            threadId: 'thr-2',
            turnId: 'replay:run-1:10',
            goalText: 'Reapply formatting',
            summary: 'Format Sheet1!A1',
            scope: 'selection',
            riskClass: 'low',
            baseRevision: 4,
            createdAtUnixMs: 10,
            context: {
              selection: {
                sheetName: 'Sheet1',
                address: 'A1',
              },
              viewport: {
                rowStart: 0,
                rowEnd: 10,
                colStart: 0,
                colEnd: 5,
              },
            },
            commands: [
              {
                kind: 'formatRange',
                range: {
                  sheetName: 'Sheet1',
                  startAddress: 'A1',
                  endAddress: 'A1',
                },
                patch: {
                  font: {
                    bold: true,
                  },
                },
              },
            ],
            affectedRanges: [
              {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'A1',
                role: 'target',
              },
            ],
            estimatedAffectedCells: 1,
            sharedReview: null,
          }),
        ],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        replayExecutionRecord,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/runs/run-1/replay',
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(replayExecutionRecord).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          recordId: 'run-1',
        }),
      )
      expect(response.json()).toEqual(
        expect.objectContaining({
          reviewQueueItems: [expect.objectContaining({ id: 'bundle-replay-1' })],
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('returns a structured not-found envelope when the chat thread event stream is stale', async () => {
    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        async createSession() {
          throw createWorkbookAgentServiceError({
            code: 'WORKBOOK_AGENT_SESSION_NOT_FOUND',
            message: 'Workbook agent session not found',
            statusCode: 404,
            retryable: true,
          })
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/doc-1/chat/threads/thr-1/events',
      })

      expect(response.statusCode).toBe(404)
      expect(response.json()).toEqual(
        expect.objectContaining({
          error: 'WORKBOOK_AGENT_SESSION_NOT_FOUND',
          message: 'Workbook agent session not found',
          retryable: true,
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('includes workbook agent observability in healthz when the service is enabled', async () => {
    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        getObservabilitySnapshot() {
          return {
            enabled: true,
            generatedAtUnixMs: 42,
            featureFlags: {
              sharedThreadsEnabled: true,
              workflowRunnerEnabled: true,
              autoApplyLowRiskEnabled: false,
              formulaWorkflowFamilyEnabled: true,
              formattingWorkflowFamilyEnabled: true,
              importWorkflowFamilyEnabled: true,
              rollupWorkflowFamilyEnabled: true,
              structuralWorkflowFamilyEnabled: true,
              allowlistedUserCount: 2,
              allowlistedDocumentCount: 1,
            },
            sessions: {
              sessionCount: 3,
              subscriberThreadCount: 2,
              subscriberCount: 4,
              activeTurnCount: 1,
              runningWorkflowCount: 1,
              reviewQueueSessionCount: 1,
              sharedPendingReviewCount: 1,
            },
            pool: {
              slotCount: 1,
              boundThreadCount: 2,
              activeTurnCount: 1,
              queuedTurnCount: 0,
              maxClients: 4,
              maxConcurrentTurnsPerClient: 1,
              maxQueuedTurnsPerClient: 8,
            },
            counters: {
              turnBackpressureCount: 1,
              workflowStartedCount: 2,
              workflowCompletedCount: 1,
              workflowFailedCount: 0,
              workflowCancelledCount: 0,
              sharedReviewApprovedCount: 0,
              sharedReviewRejectedCount: 0,
              sharedRecommendationApprovedCount: 1,
              sharedRecommendationRejectedCount: 0,
            },
          }
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/healthz',
      })

      expect(response.statusCode).toBe(200)
      expect(response.json()).toEqual(
        expect.objectContaining({
          ok: true,
          workbookAgent: expect.objectContaining({
            enabled: true,
            generatedAtUnixMs: 42,
            featureFlags: expect.objectContaining({
              allowlistedUserCount: 2,
              allowlistedDocumentCount: 1,
            }),
            sessions: expect.objectContaining({
              sessionCount: 3,
              sharedPendingReviewCount: 1,
            }),
          }),
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('exposes the workbook agent observability snapshot route', async () => {
    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        getObservabilitySnapshot() {
          return {
            enabled: true,
            generatedAtUnixMs: 99,
            featureFlags: {
              sharedThreadsEnabled: true,
              workflowRunnerEnabled: true,
              autoApplyLowRiskEnabled: true,
              formulaWorkflowFamilyEnabled: true,
              formattingWorkflowFamilyEnabled: true,
              importWorkflowFamilyEnabled: true,
              rollupWorkflowFamilyEnabled: true,
              structuralWorkflowFamilyEnabled: true,
              allowlistedUserCount: 0,
              allowlistedDocumentCount: 0,
            },
            sessions: {
              sessionCount: 0,
              subscriberThreadCount: 0,
              subscriberCount: 0,
              activeTurnCount: 0,
              runningWorkflowCount: 0,
              reviewQueueSessionCount: 0,
              sharedPendingReviewCount: 0,
            },
            pool: {
              slotCount: 0,
              boundThreadCount: 0,
              activeTurnCount: 0,
              queuedTurnCount: 0,
              maxClients: 4,
              maxConcurrentTurnsPerClient: 1,
              maxQueuedTurnsPerClient: 8,
            },
            counters: {
              turnBackpressureCount: 0,
              workflowStartedCount: 0,
              workflowCompletedCount: 0,
              workflowFailedCount: 0,
              workflowCancelledCount: 0,
              sharedReviewApprovedCount: 0,
              sharedReviewRejectedCount: 0,
              sharedRecommendationApprovedCount: 0,
              sharedRecommendationRejectedCount: 0,
            },
          }
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/agent/observability',
        headers: {
          cookie: 'bilig_session=test',
        },
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.json()).toEqual(
        expect.objectContaining({
          enabled: true,
          generatedAtUnixMs: 99,
          pool: expect.objectContaining({
            maxClients: 4,
          }),
        }),
      )
    } finally {
      await app.close()
    }
  })
})
