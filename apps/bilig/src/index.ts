import { createSyncServer } from './http/sync-server.js'
import { DocumentSessionManager } from './workbook-runtime/document-session-manager.js'
import { SyncDocumentSupervisor } from './workbook-runtime/sync-document-supervisor.js'
import { LocalDocumentSupervisor } from './workbook-runtime/local-document-supervisor.js'
import { LocalWorkbookSessionManager } from './workbook-runtime/local-workbook-session-manager.js'
import { createInProcessWorksheetExecutor } from './workbook-runtime/worksheet-executor.js'
import { createZeroSyncService } from './zero/service.js'
import { createWorkbookAgentService } from './codex-app/workbook-agent-service.js'
import { logError } from './runtime-logger.js'
import { resolveBiligAppRuntimeConfig } from './app-runtime-config.js'
import { createApplicationShutdown, registerApplicationShutdownSignals } from './app-shutdown.js'

async function main() {
  const { host, appPort, publicServerUrl, browserAppBaseUrl, maxImportBytes } = resolveBiligAppRuntimeConfig()
  const worksheetHostSessionManager = new LocalWorkbookSessionManager({
    publicServerUrl,
    browserAppBaseUrl,
    ...(maxImportBytes !== undefined ? { maxImportBytes } : {}),
  })
  const worksheetHostDocumentService = new LocalDocumentSupervisor(worksheetHostSessionManager)

  const sessionManager = new DocumentSessionManager(
    undefined,
    undefined,
    createInProcessWorksheetExecutor({
      documentService: worksheetHostDocumentService,
      serverUrl: publicServerUrl,
      browserAppBaseUrl,
    }),
    {
      publicServerUrl,
      browserAppBaseUrl,
      ...(maxImportBytes !== undefined ? { maxImportBytes } : {}),
    },
  )
  const documentService = new SyncDocumentSupervisor(sessionManager)
  const zeroSyncService = createZeroSyncService()
  const workbookAgentService = createWorkbookAgentService(zeroSyncService)

  const { app: syncApp, closeWorkbookAgent } = createSyncServer({
    sessionManager,
    documentService,
    zeroSyncService,
    workbookAgentService,
    ...(maxImportBytes !== undefined ? { maxImportBytes } : {}),
  })

  const shutdown = createApplicationShutdown({
    closeHttpServer: async () => await syncApp.close(),
    closeWorkbookAgent,
    closePersistence: async () => await zeroSyncService.close(),
  })
  const removeShutdownSignals = registerApplicationShutdownSignals({
    shutdown,
    onError: (error) => {
      process.exitCode = 1
      logError('Failed to shut down Bilig cleanly', error)
    },
  })

  try {
    await zeroSyncService.initialize()
    await syncApp.listen({ host, port: appPort })
    syncApp.log.info({ host, appPort, zeroSync: zeroSyncService.enabled }, 'bilig app listening')
  } catch (error) {
    removeShutdownSignals()
    try {
      await shutdown('startup-error')
    } catch (closeError) {
      logError('Failed to clean up after Bilig startup error', closeError)
    }
    throw error
  }
}

void (async () => {
  try {
    await main()
  } catch (error) {
    logError(error)
    process.exitCode = 1
  }
})()
