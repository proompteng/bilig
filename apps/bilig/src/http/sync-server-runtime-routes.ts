import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import { MAX_AGENT_WORKBOOK_IMPORT_BYTES, decodeAgentFrame, encodeAgentFrame } from '@bilig/agent-api'
import type { RuntimeSession } from '@bilig/contracts'
import { createRuntimeSession, type DocumentControlService, resolveRequestBaseUrl, runPromise } from '@bilig/runtime-kernel'
import type { BiligRuntimeConfig } from '@bilig/zero-sync'
import type { WorkbookAgentService } from '../codex-app/workbook-agent-service.js'
import { resolveAgentDocumentId } from '../workbook-runtime/document-supervisor-shared.js'
import type { ZeroSyncService } from '../zero/service.js'
import { resolveRequestSession, resolveSessionIdentity, type RequestSessionResolver } from './session.js'
import { resolveAuthorizedWorkbookSession, resolveWorkbookSessionWithAuthority } from './workbook-access.js'

const DEFAULT_MAX_IMPORT_BYTES = 10 * 1024 * 1024
const AGENT_FRAME_ENVELOPE_BYTES = 64 * 1024

function resolveAgentFrameBodyLimit(maxImportBytes = DEFAULT_MAX_IMPORT_BYTES): number {
  if (!Number.isSafeInteger(maxImportBytes) || maxImportBytes <= 0) {
    throw new Error('maxImportBytes must be a positive safe integer')
  }
  if (maxImportBytes > MAX_AGENT_WORKBOOK_IMPORT_BYTES) {
    throw new Error(`maxImportBytes must not exceed ${MAX_AGENT_WORKBOOK_IMPORT_BYTES}`)
  }
  const base64Bytes = 4 * Math.ceil(maxImportBytes / 3)
  const bodyLimit = base64Bytes + AGENT_FRAME_ENVELOPE_BYTES
  if (!Number.isSafeInteger(bodyLimit)) {
    throw new Error('maxImportBytes is too large to derive a safe request body limit')
  }
  return bodyLimit
}

function resolveBooleanEnv(value: string | undefined, fallback: boolean, name: string): boolean {
  if (value === undefined || value.length === 0) {
    return fallback
  }
  if (value === 'true' || value === '1') {
    return true
  }
  if (value === 'false' || value === '0') {
    return false
  }
  throw new Error(`${name} must be "1", "true", "0", or "false" when set, got ${value}`)
}

function resolveWebRuntimeConfig(env: Record<string, string | undefined>): Omit<BiligRuntimeConfig, 'currentUserId'> {
  const zeroCacheUrl = env['BILIG_ZERO_CACHE_URL']?.trim() || '/zero'
  const defaultDocumentId = env['BILIG_DEFAULT_DOCUMENT_ID']?.trim() || 'bilig-demo'

  return {
    zeroCacheUrl,
    defaultDocumentId,
    persistState: resolveBooleanEnv(env['BILIG_PERSIST_STATE'], true, 'BILIG_PERSIST_STATE'),
  }
}

export function registerSyncServerRuntimeRoutes(
  app: FastifyInstance,
  options: {
    documentService: DocumentControlService
    workbookAgentService?: WorkbookAgentService
    env: Record<string, string | undefined>
    runtimeConfig: {
      readonly browserAppBaseUrl?: string
    }
    webEnabled: boolean
    sessionResolver: RequestSessionResolver
    zeroSyncService?: ZeroSyncService
    maxImportBytes?: number
  },
): void {
  const webRuntimeConfig = resolveWebRuntimeConfig(options.env)
  const healthSnapshot = () => ({
    service: 'bilig-app',
    zeroSync: options.zeroSyncService?.enabled ?? false,
    web: options.webEnabled,
    workbookAgent: options.workbookAgentService?.enabled ?? false,
  })
  const persistenceReady = () => {
    const zeroSyncService = options.zeroSyncService
    return !zeroSyncService || !zeroSyncService.enabled || zeroSyncService.isReady()
  }

  app.get('/healthz', async (_request, reply) => {
    reply.header('cache-control', 'no-store')
    return { ok: true, ...healthSnapshot() }
  })

  app.get('/readyz', async (_request, reply) => {
    reply.header('cache-control', 'no-store')
    const ready = persistenceReady()
    if (!ready) {
      reply.code(503)
    }
    return { ok: ready, ready, ...healthSnapshot() }
  })

  app.get('/runtime-config.json', async (request, reply) => {
    const session = resolveSessionIdentity(request, reply, options.sessionResolver)
    reply.header('cache-control', 'no-store')
    return {
      ...webRuntimeConfig,
      currentUserId: session.userID,
      workbookAgentEnabled: options.workbookAgentService?.enabled ?? false,
    } satisfies BiligRuntimeConfig
  })

  const handleSessionRequest = async (request: FastifyRequest, reply: FastifyReply) => {
    const requestSession = resolveRequestSession(request, options.sessionResolver)
    options.sessionResolver.persist(reply, requestSession)
    return createRuntimeSession({
      authToken: requestSession.userId,
      userId: requestSession.userId,
      roles: requestSession.roles,
      isAuthenticated: requestSession.isAuthenticated,
      authSource: requestSession.authSource,
    }) satisfies RuntimeSession
  }
  app.get('/v2/session', handleSessionRequest)

  app.post(
    '/v2/agent/frames',
    { bodyLimit: resolveAgentFrameBodyLimit(options.maxImportBytes) },
    async (request: FastifyRequest<{ Body: Buffer }>, reply: FastifyReply) => {
      const frame = decodeAgentFrame(request.body)
      const createsWorkbook = frame.kind === 'request' && frame.request.kind === 'loadWorkbookFile' && frame.request.openMode === 'create'
      const documentId = resolveAgentDocumentId(frame)
      if (documentId) {
        await resolveAuthorizedWorkbookSession({
          request,
          reply,
          documentId,
          sessionResolver: options.sessionResolver,
          ...(options.zeroSyncService ? { zeroSyncService: options.zeroSyncService } : {}),
        })
      } else if (createsWorkbook) {
        resolveWorkbookSessionWithAuthority({
          request,
          reply,
          sessionResolver: options.sessionResolver,
          ...(options.zeroSyncService ? { zeroSyncService: options.zeroSyncService } : {}),
        })
      } else {
        resolveSessionIdentity(request, reply, options.sessionResolver)
      }
      const response = await runPromise(
        options.documentService.handleAgentFrame(frame, {
          serverUrl: resolveRequestBaseUrl(request, '127.0.0.1:4321'),
          ...(options.runtimeConfig.browserAppBaseUrl ? { browserAppBaseUrl: options.runtimeConfig.browserAppBaseUrl } : {}),
        }),
      )
      if (response.kind === 'response' && response.response.kind === 'workbookLoaded') {
        await resolveAuthorizedWorkbookSession({
          request,
          reply,
          documentId: response.response.documentId,
          sessionResolver: options.sessionResolver,
          createIfMissing: createsWorkbook,
          ...(options.zeroSyncService ? { zeroSyncService: options.zeroSyncService } : {}),
        })
      }
      reply.header('content-type', 'application/octet-stream')
      return Buffer.from(encodeAgentFrame(response))
    },
  )
}
