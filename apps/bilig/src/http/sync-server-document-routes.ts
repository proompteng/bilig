import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import { decodeFrame, encodeFrame } from '@bilig/binary-protocol'
import { createErrorEnvelope, type DocumentControlService, runPromise } from '@bilig/runtime-kernel'
import { resolveSessionIdentity } from './session.js'
import type { ZeroSyncService } from '../zero/service.js'

export function parseAfterRevisionQuery(value: string | undefined): number | null {
  const trimmed = value?.trim() ?? '0'
  if (!/^(0|[1-9]\d*)$/u.test(trimmed)) {
    return null
  }
  const revision = Number(trimmed)
  return Number.isSafeInteger(revision) ? revision : null
}

export function registerSyncServerDocumentRoutes(
  app: FastifyInstance,
  options: {
    documentService: DocumentControlService
    zeroSyncService?: ZeroSyncService
  },
): void {
  const { documentService, zeroSyncService } = options

  app.get('/v2/documents/:documentId/state', async (request: FastifyRequest<{ Params: { documentId: string } }>) => {
    return await runPromise(documentService.getDocumentState(request.params.documentId))
  })

  app.get(
    '/v2/documents/:documentId/snapshot/latest',
    async (request: FastifyRequest<{ Params: { documentId: string } }>, reply: FastifyReply) => {
      const snapshot = await runPromise(documentService.getLatestSnapshot(request.params.documentId))
      if (snapshot) {
        reply.header('x-bilig-snapshot-cursor', String(snapshot.cursor))
        reply.header('content-type', snapshot.contentType)
        return Buffer.from(snapshot.bytes)
      }

      const zeroSnapshot = zeroSyncService?.enabled ? await zeroSyncService.loadLatestWorkbookSnapshot?.(request.params.documentId) : null
      if (!zeroSnapshot) {
        reply.code(204)
        return reply.send()
      }

      reply.header('cache-control', 'no-store')
      reply.header('x-bilig-snapshot-cursor', String(zeroSnapshot.revision))
      reply.header('content-type', 'application/vnd.bilig.workbook+json')
      return JSON.stringify(zeroSnapshot.snapshot)
    },
  )

  app.get(
    '/v2/documents/:documentId/events',
    async (
      request: FastifyRequest<{
        Params: { documentId: string }
        Querystring: { afterRevision?: string }
      }>,
      reply: FastifyReply,
    ) => {
      if (!zeroSyncService?.enabled) {
        reply.code(503)
        return createErrorEnvelope('ZERO_SYNC_DISABLED', 'Authoritative workbook events require Zero sync', true)
      }
      const afterRevision = parseAfterRevisionQuery(request.query.afterRevision)
      if (afterRevision === null) {
        reply.code(400)
        return createErrorEnvelope('INVALID_AFTER_REVISION', 'afterRevision must be a non-negative integer', false)
      }
      reply.header('cache-control', 'no-store')
      return await zeroSyncService.loadAuthoritativeEvents(request.params.documentId, afterRevision)
    },
  )

  app.post(
    '/v2/documents/:documentId/frames',
    async (request: FastifyRequest<{ Params: { documentId: string }; Body: Buffer }>, reply: FastifyReply) => {
      const frame = decodeFrame(request.body)
      if (frame.documentId !== request.params.documentId) {
        reply.code(400)
        return createErrorEnvelope('DOCUMENT_ID_MISMATCH', 'Frame document id does not match route document id', false)
      }
      const response = await runPromise(documentService.handleSyncFrame(frame))
      reply.header('content-type', 'application/octet-stream')
      return Buffer.from(encodeFrame(Array.isArray(response) ? (response[0] ?? frame) : response))
    },
  )

  const handleZeroQuery = async (request: FastifyRequest, reply: FastifyReply) => {
    if (!zeroSyncService?.enabled) {
      reply.code(503)
      return createErrorEnvelope('ZERO_SYNC_DISABLED', 'Zero sync is not configured', true)
    }
    resolveSessionIdentity(request, reply)
    return await zeroSyncService.handleQuery(request)
  }

  const handleZeroMutate = async (request: FastifyRequest, reply: FastifyReply) => {
    if (!zeroSyncService?.enabled) {
      reply.code(503)
      return createErrorEnvelope('ZERO_SYNC_DISABLED', 'Zero sync is not configured', true)
    }
    resolveSessionIdentity(request, reply)
    return await zeroSyncService.handleMutate(request)
  }

  app.post('/api/zero/v2/query', handleZeroQuery)
  app.post('/api/zero/v2/mutate', handleZeroMutate)
}
