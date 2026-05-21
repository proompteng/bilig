import Fastify from 'fastify'

import { type DocumentControlService, resolveServerRuntimeConfig } from '@bilig/runtime-kernel'

import { DocumentSessionManager } from '../workbook-runtime/document-session-manager.js'
import { SyncDocumentSupervisor } from '../workbook-runtime/sync-document-supervisor.js'
import { registerSyncServerDocumentRoutes } from './sync-server-document-routes.js'
import { registerWorkbookAgentRoutes } from './workbook-agent-routes.js'
import { registerSyncServerRuntimeRoutes } from './sync-server-runtime-routes.js'
import { applySyncServerSecurityHeaders } from './sync-server-security-headers.js'
import { resolveSyncServerWebDistRoot, registerSyncServerSpaRoutes } from './sync-server-spa.js'
import { registerSyncServerZeroProxyRoutes, resolveZeroProxyUpstream } from './sync-server-zero-proxy.js'
import { registerWorkPaperMcpRemoteRoutes } from './workpaper-mcp-remote-routes.js'
import { registerWorkPaperN8nRoutes } from './workpaper-n8n-routes.js'
import type { WorksheetExecutor } from '../workbook-runtime/worksheet-executor.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookAgentService } from '../codex-app/workbook-agent-service.js'

export interface SyncServerOptions {
  sessionManager?: DocumentSessionManager
  documentService?: DocumentControlService
  worksheetExecutor?: WorksheetExecutor | null
  zeroSyncService?: ZeroSyncService
  workbookAgentService?: WorkbookAgentService
  logger?: boolean
}

export function createSyncServer(options: SyncServerOptions = {}) {
  const runtimeConfig = resolveServerRuntimeConfig(process.env)
  const webDistRoot = resolveSyncServerWebDistRoot()
  const zeroProxyUpstream = resolveZeroProxyUpstream(process.env)
  const sessionManager = options.sessionManager ?? new DocumentSessionManager(undefined, undefined, options.worksheetExecutor ?? null)
  const documentService = options.documentService ?? new SyncDocumentSupervisor(sessionManager)
  const zeroSyncService = options.zeroSyncService
  const workbookAgentService = options.workbookAgentService
  const app = Fastify({ logger: options.logger ?? true })

  app.addHook('onSend', async (_request, reply, payload) => {
    applySyncServerSecurityHeaders(reply)
    return payload
  })

  app.addContentTypeParser('application/octet-stream', { parseAs: 'buffer' }, (_request, body, done) => {
    done(null, body)
  })

  registerWorkbookAgentRoutes(app, workbookAgentService)

  if (zeroProxyUpstream) {
    registerSyncServerZeroProxyRoutes(app, zeroProxyUpstream)
  }

  registerSyncServerRuntimeRoutes(app, {
    documentService,
    env: process.env,
    runtimeConfig,
    webEnabled: webDistRoot !== null,
    ...(workbookAgentService ? { workbookAgentService } : {}),
  })

  registerSyncServerDocumentRoutes(app, {
    documentService,
    ...(zeroSyncService ? { zeroSyncService } : {}),
  })

  registerWorkPaperMcpRemoteRoutes(app)
  registerWorkPaperN8nRoutes(app)

  registerSyncServerSpaRoutes(app, webDistRoot)

  return { app, sessionManager, documentService }
}
