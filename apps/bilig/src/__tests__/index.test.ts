import { describe, expect, it, vi } from 'vitest'

const testState = vi.hoisted(() => {
  const zeroSyncService = {
    enabled: true,
    initialize: vi.fn(async () => {}),
    close: vi.fn(async () => {}),
  }
  const workbookAgentService = {
    enabled: true,
  }
  const syncApp = {
    close: vi.fn(async () => {}),
    listen: vi.fn(async () => {}),
    log: {
      info: vi.fn(),
    },
  }
  const closeWorkbookAgent = vi.fn(async () => {})

  return {
    closeWorkbookAgent,
    createSyncServer: vi.fn(() => ({ app: syncApp, closeWorkbookAgent })),
    createWorkbookAgentService: vi.fn(() => workbookAgentService),
    createZeroSyncService: vi.fn(() => zeroSyncService),
    logError: vi.fn(),
    resolveBiligAppRuntimeConfig: vi.fn(() => ({
      host: '127.0.0.1',
      appPort: 3000,
      publicServerUrl: 'http://127.0.0.1:3000',
      browserAppBaseUrl: 'http://127.0.0.1:3000',
    })),
    syncApp,
    workbookAgentService,
    zeroSyncService,
  }
})

vi.mock('../http/sync-server.js', () => ({
  createSyncServer: testState.createSyncServer,
}))
vi.mock('../workbook-runtime/document-session-manager.js', () => ({
  DocumentSessionManager: class DocumentSessionManager {
    readonly isTestDouble = true
  },
}))
vi.mock('../workbook-runtime/sync-document-supervisor.js', () => ({
  SyncDocumentSupervisor: class SyncDocumentSupervisor {
    readonly isTestDouble = true
  },
}))
vi.mock('../workbook-runtime/local-document-supervisor.js', () => ({
  LocalDocumentSupervisor: class LocalDocumentSupervisor {
    readonly isTestDouble = true
  },
}))
vi.mock('../workbook-runtime/local-workbook-session-manager.js', () => ({
  LocalWorkbookSessionManager: class LocalWorkbookSessionManager {
    readonly isTestDouble = true
  },
}))
vi.mock('../workbook-runtime/worksheet-executor.js', () => ({
  createInProcessWorksheetExecutor: vi.fn(() => ({})),
}))
vi.mock('../zero/service.js', () => ({
  createZeroSyncService: testState.createZeroSyncService,
}))
vi.mock('../codex-app/workbook-agent-service.js', () => ({
  createWorkbookAgentService: testState.createWorkbookAgentService,
}))
vi.mock('../runtime-logger.js', () => ({
  logError: testState.logError,
}))
vi.mock('../app-runtime-config.js', () => ({
  resolveBiligAppRuntimeConfig: testState.resolveBiligAppRuntimeConfig,
}))

describe('bilig startup', () => {
  it('cleans up exactly once and preserves the initialization error', async () => {
    const previousExitCode = process.exitCode
    const startupError = new Error('persistence unavailable')
    testState.zeroSyncService.initialize.mockRejectedValueOnce(startupError)

    try {
      await import('../index.js')
      await vi.waitFor(() => expect(testState.zeroSyncService.close).toHaveBeenCalledOnce())
      await vi.waitFor(() => expect(testState.logError).toHaveBeenCalledWith(startupError))

      expect(testState.createSyncServer).toHaveBeenCalledOnce()
      expect(testState.zeroSyncService.initialize).toHaveBeenCalledOnce()
      expect(testState.syncApp.listen).not.toHaveBeenCalled()
      expect(testState.syncApp.close).toHaveBeenCalledOnce()
      expect(testState.closeWorkbookAgent).toHaveBeenCalledOnce()
      expect(testState.zeroSyncService.close).toHaveBeenCalledOnce()
      expect(testState.logError).toHaveBeenCalledExactlyOnceWith(startupError)
    } finally {
      process.exitCode = previousExitCode
    }
  })
})
