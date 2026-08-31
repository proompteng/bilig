import { encodeAgentFrame } from '@bilig/agent-api'

import {
  Effect,
  createDocumentServiceStub,
  createSignedProxyTestSession,
  createSyncServer,
  createWorkbookAgentServiceError,
  createWorkbookAgentServiceStub,
  createZeroSyncStub,
  describe,
  expect,
  it,
  vi,
} from './sync-server-test-helpers.js'

describe('sync-server lifecycle ownership', () => {
  it('closes the workbook agent exactly once when HTTP close and explicit draining overlap', async () => {
    const close = vi.fn(async () => {})
    const { app, closeWorkbookAgent } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({ close }),
    })

    await Promise.all([app.close(), closeWorkbookAgent()])

    expect(close).toHaveBeenCalledOnce()
  })
})

describe('sync-server agent import ingress', () => {
  const largeUploadFrame = Buffer.from(
    encodeAgentFrame({
      kind: 'request',
      request: {
        kind: 'loadWorkbookFile',
        id: 'upload-large',
        replicaId: 'agent-local',
        openMode: 'create',
        fileName: 'large.csv',
        contentType: 'text/csv',
        bytesBase64: Buffer.alloc(800 * 1024).toString('base64'),
      },
    }),
  )

  it('accepts configured workbook uploads larger than Fastify default body limit', async () => {
    const handleAgentFrame = vi.fn(() =>
      Effect.succeed({
        kind: 'response' as const,
        response: {
          kind: 'error' as const,
          id: 'upload-large',
          code: 'TEST_STOP',
          message: 'Accepted by ingress',
          retryable: false,
        },
      }),
    )
    const { app } = createSyncServer({
      documentService: createDocumentServiceStub({ handleAgentFrame }),
      logger: false,
      maxImportBytes: 1024 * 1024,
    })

    try {
      expect(largeUploadFrame.byteLength).toBeGreaterThan(1024 * 1024)
      const response = await app.inject({
        method: 'POST',
        url: '/v2/agent/frames',
        headers: { 'content-type': 'application/octet-stream' },
        payload: largeUploadFrame,
      })

      expect(response.statusCode).toBe(200)
      expect(handleAgentFrame).toHaveBeenCalledOnce()
    } finally {
      await app.close()
    }
  })

  it('rejects uploads above the configured decoded workbook budget at ingress', async () => {
    const handleAgentFrame = vi.fn(() => Effect.die('must not execute'))
    const { app } = createSyncServer({
      documentService: createDocumentServiceStub({ handleAgentFrame }),
      logger: false,
      maxImportBytes: 512 * 1024,
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/agent/frames',
        headers: { 'content-type': 'application/octet-stream' },
        payload: largeUploadFrame,
      })

      expect(response.statusCode).toBe(413)
      expect(handleAgentFrame).not.toHaveBeenCalled()
    } finally {
      await app.close()
    }
  })
})

describe('sync-server remote MCP origin policy', () => {
  it('treats an explicit remote MCP origin list as the complete allowlist', async () => {
    process.env['BILIG_REMOTE_MCP_ALLOWED_ORIGINS'] = 'https://trusted.example'
    const { app } = createSyncServer({ logger: false })

    try {
      const [trusted, defaultOrigin] = await Promise.all([
        app.inject({
          method: 'OPTIONS',
          url: '/mcp',
          headers: { origin: 'https://trusted.example' },
        }),
        app.inject({
          method: 'OPTIONS',
          url: '/mcp',
          headers: { origin: 'https://chatgpt.com' },
        }),
      ])

      expect(trusted.statusCode).toBe(204)
      expect(trusted.headers['access-control-allow-origin']).toBe('https://trusted.example')
      expect(defaultOrigin.statusCode).toBe(403)
    } finally {
      await app.close()
    }
  })

  it('can disable local remote MCP origins explicitly', async () => {
    process.env['BILIG_REMOTE_MCP_ALLOW_LOCAL_ORIGINS'] = 'false'
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'OPTIONS',
        url: '/mcp',
        headers: { origin: 'http://localhost:5173' },
      })

      expect(response.statusCode).toBe(403)
    } finally {
      await app.close()
    }
  })

  it('rejects malformed remote MCP origin configuration during startup', () => {
    process.env['BILIG_REMOTE_MCP_ALLOWED_ORIGINS'] = 'https://trusted.example/not-an-origin'

    expect(() => createSyncServer({ logger: false })).toThrow(
      'BILIG_REMOTE_MCP_ALLOWED_ORIGINS entries must be HTTP(S) origins without credentials, paths, queries, or fragments',
    )
  })
})

describe('sync-server request authentication', () => {
  it('rejects demo authentication during production server startup', () => {
    vi.stubEnv('NODE_ENV', 'production')
    vi.stubEnv('BILIG_AUTH_MODE', 'demo')
    vi.stubEnv('BILIG_SESSION_SECRET', 'production-demo-secret-that-is-at-least-32-bytes')

    try {
      expect(() => createSyncServer({ logger: false })).toThrow('BILIG_AUTH_MODE=demo is not allowed in production; use signed-proxy')
    } finally {
      vi.unstubAllEnvs()
    }
  })

  it('uses a signed anonymous demo cookie and ignores spoofed identity headers', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/runtime-config.json',
        headers: {
          authorization: 'Bearer admin@example.com',
          'x-bilig-user-id': 'admin@example.com',
        },
      })

      expect(response.statusCode).toBe(200)
      expect(response.json().currentUserId).toMatch(/^guest:/u)
      expect(response.json().currentUserId).not.toBe('admin@example.com')
      expect(response.headers['set-cookie']).toContain('bilig_session=')
      expect(response.headers['set-cookie']).toContain('HttpOnly')
      expect(response.headers['set-cookie']).toContain('SameSite=Lax')
    } finally {
      await app.close()
    }
  })

  it('rejects requests without a signed assertion in private mode', async () => {
    const auth = createSignedProxyTestSession('owner-1')
    const { app } = createSyncServer({ logger: false, sessionResolver: auth.sessionResolver })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/session',
      })

      expect(response.statusCode).toBe(401)
      expect(response.json()).toMatchObject({
        code: 'REQUEST_AUTHENTICATION_REQUIRED',
        message: 'Missing signed proxy assertion',
      })
    } finally {
      await app.close()
    }
  })
})

describe('sync-server workbook authorization', () => {
  it('denies a signed user before reading another user workbook', async () => {
    const auth = createSignedProxyTestSession('mallory@example.com')
    const getDocumentState = vi.fn(() =>
      Effect.succeed({
        documentId: 'private-book',
        cursor: 0,
        owner: null,
        sessions: [],
        latestSnapshotCursor: null,
      }),
    )
    const { app } = createSyncServer({
      logger: false,
      sessionResolver: auth.sessionResolver,
      documentService: createDocumentServiceStub({ getDocumentState }),
      zeroSyncService: createZeroSyncStub({
        async assertWorkbookAccess(_documentId, _session, _mode, options) {
          expect(options).toEqual({ createIfMissing: false })
          throw Object.assign(new Error('Workbook access denied'), {
            statusCode: 403,
            code: 'WORKBOOK_ACCESS_DENIED',
          })
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/private-book/state',
        headers: auth.headers,
      })

      expect(response.statusCode).toBe(403)
      expect(getDocumentState).not.toHaveBeenCalled()
    } finally {
      await app.close()
    }
  })

  it('creates ownership metadata only after an authorized new-workbook import', async () => {
    const auth = createSignedProxyTestSession('owner@example.com')
    const assertWorkbookAccess = vi.fn(async () => {})
    const { app } = createSyncServer({
      logger: false,
      sessionResolver: auth.sessionResolver,
      documentService: createDocumentServiceStub({
        handleAgentFrame() {
          return Effect.succeed({
            kind: 'response',
            response: {
              kind: 'workbookLoaded',
              id: 'upload-1',
              documentId: 'csv:new-book',
              sessionId: 'csv:new-book:agent-local',
              workbookName: 'new-book',
              sheetNames: ['Sheet1'],
              serverUrl: 'https://bilig.example.test',
              warnings: [],
            },
          })
        },
      }),
      zeroSyncService: createZeroSyncStub({ assertWorkbookAccess }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/agent/frames',
        headers: { ...auth.headers, 'content-type': 'application/octet-stream' },
        payload: Buffer.from(
          encodeAgentFrame({
            kind: 'request',
            request: {
              kind: 'loadWorkbookFile',
              id: 'upload-1',
              replicaId: 'agent-local',
              openMode: 'create',
              fileName: 'new-book.csv',
              contentType: 'text/csv',
              bytesBase64: 'YSxi',
            },
          }),
        ),
      })

      expect(response.statusCode).toBe(200)
      expect(assertWorkbookAccess).toHaveBeenCalledOnce()
      expect(assertWorkbookAccess).toHaveBeenCalledWith(
        'csv:new-book',
        { userID: 'owner@example.com', roles: ['editor'] },
        'signed-proxy',
        { createIfMissing: true },
      )
    } finally {
      await app.close()
    }
  })

  it('rejects private workbook creation before import when authoritative authorization is unavailable', async () => {
    const auth = createSignedProxyTestSession('owner@example.com')
    const handleAgentFrame = vi.fn(() => Effect.die('must not import without authoritative authorization'))
    const { app } = createSyncServer({
      logger: false,
      sessionResolver: auth.sessionResolver,
      documentService: createDocumentServiceStub({ handleAgentFrame }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/agent/frames',
        headers: { ...auth.headers, 'content-type': 'application/octet-stream' },
        payload: Buffer.from(
          encodeAgentFrame({
            kind: 'request',
            request: {
              kind: 'loadWorkbookFile',
              id: 'upload-1',
              replicaId: 'agent-local',
              openMode: 'create',
              fileName: 'new-book.csv',
              contentType: 'text/csv',
              bytesBase64: 'YSxi',
            },
          }),
        ),
      })

      expect(response.statusCode).toBe(503)
      expect(response.json()).toMatchObject({
        code: 'WORKBOOK_AUTHORIZATION_UNAVAILABLE',
      })
      expect(handleAgentFrame).not.toHaveBeenCalled()
    } finally {
      await app.close()
    }
  })

  it('keeps signed identity on chat event-stream session lookup', async () => {
    const auth = createSignedProxyTestSession('owner@example.com')
    const createSession = vi.fn(async () => {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_SESSION_NOT_FOUND',
        message: 'Workbook agent session not found',
        statusCode: 404,
        retryable: true,
      })
    })
    const { app } = createSyncServer({
      logger: false,
      sessionResolver: auth.sessionResolver,
      workbookAgentService: createWorkbookAgentServiceStub({ createSession }),
      zeroSyncService: createZeroSyncStub(),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/private-book/chat/threads/thread-1/events',
        headers: auth.headers,
      })

      expect(response.statusCode).toBe(404)
      expect(createSession).toHaveBeenCalledWith({
        documentId: 'private-book',
        session: { userID: 'owner@example.com', roles: ['editor'] },
        body: { threadId: 'thread-1' },
      })
    } finally {
      await app.close()
    }
  })
})
