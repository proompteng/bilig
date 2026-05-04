import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import type { WorkbookAgentStreamEvent } from '@bilig/contracts'
import { createErrorEnvelope } from '@bilig/runtime-kernel'
import type { SessionIdentity } from './session.js'
import { resolveSessionIdentity } from './session.js'
import type { WorkbookAgentService } from '../codex-app/workbook-agent-service.js'
import { isWorkbookAgentServiceError } from '../workbook-agent-errors.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

async function loadWorkbookAgentThreadSession(
  service: WorkbookAgentService,
  documentId: string,
  threadId: string,
  session: SessionIdentity,
) {
  return await service.createSession({
    documentId,
    session,
    body: {
      threadId,
    },
  })
}

function readWorkbookAgentThreadRouteParams(request: FastifyRequest): {
  documentId: string
  threadId: string
} {
  const params: unknown = request.params
  if (!isRecord(params) || typeof params['documentId'] !== 'string' || typeof params['threadId'] !== 'string') {
    throw new Error('Expected workbook agent thread route params')
  }
  return {
    documentId: params['documentId'],
    threadId: params['threadId'],
  }
}

export function registerWorkbookAgentRoutes(app: FastifyInstance, workbookAgentService?: WorkbookAgentService): void {
  const handleWorkbookAgentRequest = async <T>(
    request: FastifyRequest,
    reply: FastifyReply,
    task: (service: WorkbookAgentService, session: SessionIdentity) => Promise<T>,
  ): Promise<T | ReturnType<typeof createErrorEnvelope>> => {
    if (!workbookAgentService?.enabled) {
      reply.code(503)
      return createErrorEnvelope('WORKBOOK_AGENT_DISABLED', 'Workbook agent service is not configured', true)
    }
    const session = resolveSessionIdentity(request, reply)
    reply.header('cache-control', 'no-store')
    try {
      return await task(workbookAgentService, session)
    } catch (error) {
      if (isWorkbookAgentServiceError(error)) {
        reply.code(error.statusCode)
        return createErrorEnvelope(error.code, error.message, error.retryable)
      }
      throw error
    }
  }

  const handleWorkbookAgentThreadRequest = async <T>(
    request: FastifyRequest,
    reply: FastifyReply,
    task: (
      service: WorkbookAgentService,
      session: SessionIdentity,
      sessionSnapshot: Awaited<ReturnType<WorkbookAgentService['createSession']>>,
    ) => Promise<T>,
  ): Promise<T | ReturnType<typeof createErrorEnvelope>> => {
    return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
      const { documentId, threadId } = readWorkbookAgentThreadRouteParams(request)
      const sessionSnapshot = await loadWorkbookAgentThreadSession(service, documentId, threadId, session)
      return await task(service, session, sessionSnapshot)
    })
  }

  app.addHook('onClose', async () => {
    await workbookAgentService?.close().catch(() => undefined)
  })

  app.get('/v2/agent/observability', async (request, reply) => {
    return await handleWorkbookAgentRequest(request, reply, async (service) => service.getObservabilitySnapshot())
  })

  app.get(
    '/v2/documents/:documentId/chat/threads',
    async (
      request: FastifyRequest<{
        Params: { documentId: string }
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        return await service.listThreads({
          documentId: request.params.documentId,
          session,
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads',
    async (
      request: FastifyRequest<{
        Params: { documentId: string }
        Body: Record<string, unknown>
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        return await service.createSession({
          documentId: request.params.documentId,
          session,
          body: request.body ?? {},
        })
      })
    },
  )

  app.get(
    '/v2/documents/:documentId/chat/threads/:threadId',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string }
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (_service, _session, snapshot) => {
        return snapshot
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/turns',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string }
        Body: Record<string, unknown>
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        return await service.startTurn({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          session,
          body: request.body ?? {},
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/context',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string }
        Body: Record<string, unknown>
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        return await service.updateContext({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          session,
          body: request.body ?? {},
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/workflows',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string }
        Body: Record<string, unknown>
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        return await service.startWorkflow({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          session,
          body: request.body ?? {},
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/workflows/:runId/cancel',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string; runId: string }
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        return await service.cancelWorkflow({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          runId: request.params.runId,
          session,
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/interrupt',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string }
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        return await service.interruptTurn({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          session,
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/review-items/:reviewItemId/apply',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string; reviewItemId: string }
        Body: {
          appliedBy?: 'user' | 'auto'
          commandIndexes?: number[]
          preview?: unknown
        }
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        const commandIndexes =
          request.body && typeof request.body === 'object' && Array.isArray(request.body.commandIndexes)
            ? request.body.commandIndexes
            : undefined
        return await service.applyReviewItem({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          reviewItemId: request.params.reviewItemId,
          session,
          appliedBy: request.body && request.body.appliedBy === 'auto' ? 'auto' : 'user',
          ...(commandIndexes ? { commandIndexes } : {}),
          preview: request.body && typeof request.body === 'object' && 'preview' in request.body ? (request.body.preview ?? null) : null,
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/review-items/:reviewItemId/review',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string; reviewItemId: string }
        Body: {
          decision?: 'approved' | 'rejected'
        }
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        return await service.reviewReviewItem({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          reviewItemId: request.params.reviewItemId,
          session,
          body: request.body ?? {},
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/review-items/:reviewItemId/dismiss',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string; reviewItemId: string }
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        return await service.dismissReviewItem({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          reviewItemId: request.params.reviewItemId,
          session,
        })
      })
    },
  )

  app.post(
    '/v2/documents/:documentId/chat/threads/:threadId/runs/:recordId/replay',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string; recordId: string }
      }>,
      reply,
    ) => {
      return await handleWorkbookAgentThreadRequest(request, reply, async (service, session, sessionSnapshot) => {
        return await service.replayExecutionRecord({
          documentId: request.params.documentId,
          threadId: sessionSnapshot.threadId,
          recordId: request.params.recordId,
          session,
        })
      })
    },
  )

  app.get(
    '/v2/documents/:documentId/chat/threads/:threadId/events',
    async (
      request: FastifyRequest<{
        Params: { documentId: string; threadId: string }
      }>,
      reply,
    ) => {
      if (!workbookAgentService?.enabled) {
        reply.code(503)
        return createErrorEnvelope('WORKBOOK_AGENT_DISABLED', 'Workbook agent service is not configured', true)
      }
      const session = resolveSessionIdentity(request, reply)
      let sessionSnapshot: Awaited<ReturnType<typeof workbookAgentService.createSession>>
      try {
        sessionSnapshot = await loadWorkbookAgentThreadSession(
          workbookAgentService,
          request.params.documentId,
          request.params.threadId,
          session,
        )
      } catch (error) {
        if (isWorkbookAgentServiceError(error)) {
          reply.code(error.statusCode)
          return createErrorEnvelope(error.code, error.message, error.retryable)
        }
        throw error
      }

      reply.hijack()
      const raw = reply.raw
      raw.writeHead(200, {
        'content-type': 'text/event-stream; charset=utf-8',
        'cache-control': 'no-store',
        connection: 'keep-alive',
      })

      const writeEvent = (event: WorkbookAgentStreamEvent) => {
        raw.write(`data: ${JSON.stringify(event)}\n\n`)
      }

      writeEvent({
        type: 'snapshot',
        snapshot: sessionSnapshot,
      })

      const unsubscribe = workbookAgentService.subscribe(sessionSnapshot.threadId, (event) => {
        writeEvent(event)
      })
      const keepalive = setInterval(() => {
        raw.write(':keepalive\n\n')
      }, 15_000)

      request.raw.on('close', () => {
        clearInterval(keepalive)
        unsubscribe()
      })
      return reply
    },
  )
}
