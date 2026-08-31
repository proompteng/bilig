import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

const deps = vi.hoisted(() => {
  const pool = {
    query: vi.fn(),
    connect: vi.fn(),
    end: vi.fn(async () => undefined),
  }
  return {
    pool,
    ensureZeroSyncSchema: vi.fn(async () => undefined),
    ensureWorkbookPresenceSchema: vi.fn(async () => undefined),
    ensureWorkbookChangeSchema: vi.fn(async () => undefined),
    ensureWorkbookAgentRunSchema: vi.fn(async () => undefined),
    ensureWorkbookChatThreadSchema: vi.fn(async () => undefined),
    ensureWorkbookWorkflowRunSchema: vi.fn(async () => undefined),
    ensureZeroPublication: vi.fn(async () => undefined),
    ensureZeroDataMigrationSchema: vi.fn(async () => undefined),
    runPendingZeroDataMigrations: vi.fn(async () => undefined),
    assertZeroDataMigrationsReady: vi.fn(async () => undefined),
    recalcStart: vi.fn(),
    recalcStop: vi.fn(),
    runtimeClose: vi.fn(async () => undefined),
    handleQueryRequest: vi.fn(),
    handleMutateRequest: vi.fn(),
    loadWorkbookEventRecordsAfter: vi.fn(async () => []),
    ensureWorkbookDocumentExists: vi.fn(async () => undefined),
    loadWorkbookRuntimeMetadata: vi.fn(async () => ({
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: 'system',
    })),
    loadWorkbookState: vi.fn(async () => ({
      snapshot: {
        version: 1,
        workbook: {
          name: 'book-1',
        },
        sheets: [],
      },
      replicaSnapshot: null,
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: 'system',
    })),
  }
})

vi.mock('@rocicorp/zero/server', () => ({
  handleQueryRequest: deps.handleQueryRequest,
  handleMutateRequest: deps.handleMutateRequest,
}))

vi.mock('../db.js', () => ({
  resolveZeroDatabaseUrl: () => 'postgres://example.test/bilig',
  createZeroPool: () => deps.pool,
  createZeroDbProvider: () => ({}),
  createWorkbookRuntimeStoreConnection: () => ({
    query: deps.pool.query,
    run: vi.fn(async () => undefined),
  }),
}))

vi.mock('../zero-schema-store.js', () => ({
  ensureZeroSyncSchema: deps.ensureZeroSyncSchema,
}))

vi.mock('../presence-store.js', () => ({
  ensureWorkbookPresenceSchema: deps.ensureWorkbookPresenceSchema,
}))

vi.mock('../workbook-change-store.js', () => ({
  ensureWorkbookChangeSchema: deps.ensureWorkbookChangeSchema,
  listWorkbookChanges: vi.fn(async () => []),
}))

vi.mock('../workbook-agent-run-store.js', () => ({
  ensureWorkbookAgentRunSchema: deps.ensureWorkbookAgentRunSchema,
  appendWorkbookAgentRun: vi.fn(async () => undefined),
  listWorkbookAgentThreadRuns: vi.fn(async () => []),
  listWorkbookAgentRuns: vi.fn(async () => []),
}))

vi.mock('../workbook-chat-thread-store.js', () => ({
  ensureWorkbookChatThreadSchema: deps.ensureWorkbookChatThreadSchema,
  listWorkbookAgentThreadSummaries: vi.fn(async () => []),
  loadWorkbookAgentThreadState: vi.fn(async () => null),
  saveWorkbookAgentThreadState: vi.fn(async () => undefined),
}))

vi.mock('../workbook-workflow-run-store.js', () => ({
  ensureWorkbookWorkflowRunSchema: deps.ensureWorkbookWorkflowRunSchema,
  listWorkbookThreadWorkflowRuns: vi.fn(async () => []),
  upsertWorkbookWorkflowRun: vi.fn(async () => undefined),
}))

vi.mock('../publication-store.js', () => ({
  ensureZeroPublication: deps.ensureZeroPublication,
}))

vi.mock('../data-migration-runner.js', () => ({
  ensureZeroDataMigrationSchema: deps.ensureZeroDataMigrationSchema,
  runPendingZeroDataMigrations: deps.runPendingZeroDataMigrations,
  assertZeroDataMigrationsReady: deps.assertZeroDataMigrationsReady,
  resolveRunDataMigrationsOnBoot: () => process.env['BILIG_RUN_DATA_MIGRATIONS_ON_BOOT'] === 'true',
  resolveAllowPendingCleanupMigrations: () => process.env['BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS'] === 'true',
}))

vi.mock('../../workbook-runtime/runtime-manager.js', () => ({
  WorkbookRuntimeManager: class {
    async close() {
      await deps.runtimeClose()
    }
  },
}))

vi.mock('../recalc-worker.js', () => ({
  ZeroRecalcWorker: class {
    start() {
      deps.recalcStart()
    }

    stop() {
      deps.recalcStop()
    }
  },
}))

vi.mock('../server-mutators.js', () => ({
  handleServerMutator: vi.fn(async () => undefined),
}))

vi.mock('../store.js', () => ({
  loadWorkbookEventRecordsAfter: deps.loadWorkbookEventRecordsAfter,
}))

vi.mock('../workbook-mutation-store.js', () => ({
  persistWorkbookMutation: vi.fn(async () => {
    throw new Error('not used')
  }),
}))

vi.mock('../workbook-migration-store.js', () => ({
  ensureWorkbookDocumentExists: deps.ensureWorkbookDocumentExists,
}))

vi.mock('../workbook-runtime-store.js', () => ({
  acquireWorkbookMutationLock: vi.fn(async () => undefined),
  loadWorkbookRuntimeMetadata: deps.loadWorkbookRuntimeMetadata,
  loadWorkbookState: deps.loadWorkbookState,
}))

describe('zero sync service startup', () => {
  const ownerSession = { userID: 'user-1', roles: ['editor'] }

  beforeEach(() => {
    vi.clearAllMocks()
    deps.loadWorkbookEventRecordsAfter.mockResolvedValue([])
    deps.loadWorkbookRuntimeMetadata.mockResolvedValue({
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: 'system',
    })
    deps.loadWorkbookState.mockResolvedValue({
      snapshot: {
        version: 1,
        workbook: {
          name: 'book-1',
        },
        sheets: [],
      },
      replicaSnapshot: null,
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: 'system',
    })
    delete process.env['BILIG_RUN_DATA_MIGRATIONS_ON_BOOT']
    delete process.env['BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS']
  })

  afterEach(() => {
    delete process.env['BILIG_RUN_DATA_MIGRATIONS_ON_BOOT']
    delete process.env['BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS']
  })

  it('gates startup on migration readiness without auto-running by default', async () => {
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    expect(service.isReady()).toBe(false)
    await service.initialize()

    expect(deps.ensureZeroSyncSchema).toHaveBeenCalledWith(deps.pool)
    expect(deps.ensureWorkbookPresenceSchema).toHaveBeenCalledWith(deps.pool)
    expect(deps.ensureWorkbookChangeSchema).toHaveBeenCalledWith(deps.pool)
    expect(deps.ensureWorkbookAgentRunSchema).toHaveBeenCalledWith(deps.pool)
    expect(deps.ensureWorkbookChatThreadSchema).toHaveBeenCalledWith(deps.pool)
    expect(deps.ensureWorkbookWorkflowRunSchema).toHaveBeenCalledWith(deps.pool)
    expect(deps.ensureZeroPublication).toHaveBeenCalledWith(deps.pool)
    expect(deps.ensureZeroDataMigrationSchema).toHaveBeenCalledWith(deps.pool)
    expect(deps.runPendingZeroDataMigrations).not.toHaveBeenCalled()
    expect(deps.assertZeroDataMigrationsReady).toHaveBeenCalledWith(deps.pool, {
      allowPendingCleanup: false,
    })
    expect(deps.recalcStart).toHaveBeenCalledOnce()
    expect(service.isReady()).toBe(true)
  }, 15_000)

  it('coalesces concurrent initialization and makes close idempotent', async () => {
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await Promise.all([service.initialize(), service.initialize()])
    await Promise.all([service.close(), service.close()])

    expect(deps.recalcStart).toHaveBeenCalledOnce()
    expect(deps.recalcStop).toHaveBeenCalledOnce()
    expect(deps.runtimeClose).toHaveBeenCalledOnce()
    expect(deps.pool.end).toHaveBeenCalledOnce()
    expect(service.isReady()).toBe(false)
  }, 15_000)

  it('remains unready when initialization fails and still releases resources on close', async () => {
    const startupError = new Error('schema unavailable')
    deps.ensureZeroSyncSchema.mockRejectedValueOnce(startupError)
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await expect(service.initialize()).rejects.toBe(startupError)
    expect(service.isReady()).toBe(false)

    await service.close()

    expect(deps.recalcStart).not.toHaveBeenCalled()
    expect(deps.recalcStop).toHaveBeenCalledOnce()
    expect(deps.runtimeClose).toHaveBeenCalledOnce()
    expect(deps.pool.end).toHaveBeenCalledOnce()
  }, 15_000)

  it('attempts every shutdown boundary before reporting close failures', async () => {
    deps.runtimeClose.mockRejectedValueOnce(new Error('runtime close failed'))
    deps.pool.end.mockRejectedValueOnce(new Error('pool close failed'))
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await expect(service.close()).rejects.toMatchObject({
      errors: expect.arrayContaining([expect.any(Error), expect.any(Error)]),
    })
    expect(deps.recalcStop).toHaveBeenCalledOnce()
    expect(deps.runtimeClose).toHaveBeenCalledOnce()
    expect(deps.pool.end).toHaveBeenCalledOnce()
    expect(service.isReady()).toBe(false)
  }, 15_000)

  it('auto-runs pending migrations on boot when explicitly enabled', async () => {
    process.env['BILIG_RUN_DATA_MIGRATIONS_ON_BOOT'] = 'true'
    process.env['BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS'] = 'true'
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await service.initialize()

    expect(deps.runPendingZeroDataMigrations).toHaveBeenCalledWith(deps.pool)
    expect(deps.assertZeroDataMigrationsReady).toHaveBeenCalledWith(deps.pool, {
      allowPendingCleanup: true,
    })
    expect(deps.recalcStart).toHaveBeenCalledOnce()
  }, 15_000)

  it('preserves the trusted request protocol and canonical host for Zero query requests', async () => {
    deps.handleQueryRequest.mockImplementationOnce(async (_executeQuery, _schema, request: Request) => ({
      method: request.method,
      url: request.url,
    }))
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    const result = await service.handleQuery(
      {
        protocol: 'https',
        method: 'POST',
        url: '/zero/query?hash=abc',
        headers: {
          host: ['   ', ' sheets.example.com:8443 '],
        },
        body: {
          name: 'workbook.get',
        },
      },
      ownerSession,
      'demo',
    )

    expect(result).toEqual({
      method: 'POST',
      url: 'https://sheets.example.com:8443/zero/query?hash=abc',
    })
  })

  it('serves every exported Zero query alias through the shared transform registry', async () => {
    deps.handleQueryRequest.mockImplementationOnce(async (executeQuery: (name: string, args: unknown) => unknown) => {
      const workbookArgs = { documentId: 'doc-1' }
      const workbookCellArgs = { ...workbookArgs, sheetName: 'Sheet1', address: 'A1' }
      const workbookTileArgs = { ...workbookArgs, sheetName: 'Sheet1', rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 2 }
      const workbookRowTileArgs = { ...workbookArgs, sheetName: 'Sheet1', rowStart: 0, rowEnd: 2 }
      const workbookColumnTileArgs = { ...workbookArgs, sheetName: 'Sheet1', colStart: 0, colEnd: 2 }
      const workbookThreadArgs = { ...workbookArgs, threadId: 'thr-1' }
      const queryRequests: ReadonlyArray<readonly [string, unknown]> = [
        ['workbook.get', workbookArgs],
        ['workbooks.get', workbookArgs],
        ['sheet.byWorkbook', workbookArgs],
        ['sheets.byWorkbook', workbookArgs],
        ['cellInput.one', workbookCellArgs],
        ['cellInput.tile', workbookTileArgs],
        ['cells.one', workbookCellArgs],
        ['cells.tile', workbookTileArgs],
        ['cellEval.one', workbookCellArgs],
        ['cellEval.tile', workbookTileArgs],
        ['cellRender.one', workbookCellArgs],
        ['cellRender.tile', workbookTileArgs],
        ['sheetRow.tile', workbookRowTileArgs],
        ['rowMetadata.tile', workbookRowTileArgs],
        ['sheetCol.tile', workbookColumnTileArgs],
        ['columnMetadata.tile', workbookColumnTileArgs],
        ['cellStyle.byWorkbook', workbookArgs],
        ['numberFormat.byWorkbook', workbookArgs],
        ['presenceCoarse.byWorkbook', workbookArgs],
        ['presence.byWorkbook', workbookArgs],
        ['workbookChange.byWorkbook', workbookArgs],
        ['workbookChanges.byWorkbook', workbookArgs],
        ['workbookChatThread.byWorkbook', workbookArgs],
        ['workbookChatThread.visibleByWorkbook', workbookArgs],
        ['workbookChatItem.byThread', workbookThreadArgs],
        ['workbookChatToolCall.byThread', workbookThreadArgs],
        ['workbookReviewQueueItem.byThread', workbookThreadArgs],
        ['workbookAgentThread.byWorkbook', workbookArgs],
        ['workbookAgentThread.visibleByWorkbook', workbookArgs],
        ['workbookAgentRun.byWorkbook', workbookArgs],
        ['workbookAgentRun.byThread', workbookThreadArgs],
        ['workbookAgentRun.visibleByThread', workbookThreadArgs],
        ['workbookWorkflowRun.byThread', workbookThreadArgs],
        ['workbookWorkflowRun.visibleByThread', workbookThreadArgs],
        ['workbookWorkflowStep.byThread', workbookThreadArgs],
        ['workbookWorkflowArtifact.byThread', workbookThreadArgs],
        ['workbookAgentWorkflowRun.byThread', workbookThreadArgs],
        ['workbookAgentWorkflowRun.visibleByThread', workbookThreadArgs],
      ]
      for (const [name, args] of queryRequests) {
        expect(() => executeQuery(name, args), name).not.toThrow()
      }
      return { ok: true }
    })
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await expect(
      service.handleQuery(
        {
          protocol: 'https',
          method: 'POST',
          url: '/zero/query',
          headers: {
            host: 'sheets.example.com',
          },
          body: {},
        },
        { userID: 'alex@example.com', roles: ['editor'] },
        'demo',
      ),
    ).resolves.toEqual({ ok: true })
  })

  it('rejects malformed request protocols before forwarding Zero mutations', async () => {
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await expect(
      service.handleMutate(
        {
          protocol: 'ftp',
          method: 'POST',
          url: '/zero/mutate',
          headers: {
            host: 'sheets.example.com',
          },
          body: {
            mutations: [],
          },
        },
        ownerSession,
        'demo',
      ),
    ).rejects.toThrow('request protocol must be "http" or "https", got ftp')
    expect(deps.handleMutateRequest).not.toHaveBeenCalled()
  })

  it('rejects authoritative event batches that do not cover the advertised head revision', async () => {
    deps.loadWorkbookRuntimeMetadata.mockResolvedValueOnce({
      headRevision: 3,
      calculatedRevision: 3,
      ownerUserId: 'system',
    })
    deps.loadWorkbookEventRecordsAfter.mockResolvedValueOnce([
      {
        revision: 3,
        clientMutationId: null,
        payload: {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 3,
        },
      },
    ])
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await expect(service.loadAuthoritativeEvents('book-1', 1)).rejects.toThrow('Invalid authoritative workbook event batch for book-1')
  })

  it('allows only the workbook owner or an administrator in private mode', async () => {
    deps.loadWorkbookRuntimeMetadata.mockResolvedValue({
      headRevision: 4,
      calculatedRevision: 4,
      ownerUserId: 'owner@example.com',
    })
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await expect(
      service.assertWorkbookAccess?.('book-1', { userID: 'owner@example.com', roles: ['editor'] }, 'signed-proxy'),
    ).resolves.toBeUndefined()
    await expect(
      service.assertWorkbookAccess?.('book-1', { userID: 'admin@example.com', roles: ['admin'] }, 'signed-proxy'),
    ).resolves.toBeUndefined()
    await expect(
      service.assertWorkbookAccess?.('book-1', { userID: 'mallory@example.com', roles: ['editor'] }, 'signed-proxy'),
    ).rejects.toMatchObject({
      code: 'WORKBOOK_ACCESS_DENIED',
      statusCode: 403,
    })
    expect(deps.ensureWorkbookDocumentExists).not.toHaveBeenCalled()
  })

  it('creates private workbook ownership only when the caller marks an explicit creation boundary', async () => {
    deps.loadWorkbookRuntimeMetadata.mockResolvedValue({
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: 'owner@example.com',
    })
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await service.assertWorkbookAccess?.('new-book', { userID: 'owner@example.com', roles: ['editor'] }, 'signed-proxy', {
      createIfMissing: true,
    })

    expect(deps.ensureWorkbookDocumentExists).toHaveBeenCalledOnce()
    expect(deps.ensureWorkbookDocumentExists).toHaveBeenCalledWith(deps.pool, 'new-book', 'owner@example.com')
  })

  it('keeps explicit demo mode public without consulting private ownership', async () => {
    const { createZeroSyncService } = await import('../service.js')
    const service = createZeroSyncService()

    await expect(service.assertWorkbookAccess?.('book-1', ownerSession, 'demo')).resolves.toBeUndefined()
    expect(deps.loadWorkbookRuntimeMetadata).not.toHaveBeenCalled()
  })
})
