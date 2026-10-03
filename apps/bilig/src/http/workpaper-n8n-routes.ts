import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import {
  createN8nForecastProof,
  createN8nWorkPaperEvaluationProof,
  type N8nForecastRequestBody,
  type N8nWorkPaperEvaluationRequestBody,
} from '@bilig/workpaper'

export function registerWorkPaperN8nRoutes(app: FastifyInstance): void {
  app.post('/api/workpaper/n8n/forecast', handleN8nForecastRequest)
  app.post('/api/workpaper/n8n/evaluate', handleN8nEvaluationRequest)
}

function handleN8nForecastRequest(request: FastifyRequest<{ Body: N8nForecastRequestBody }>, reply: FastifyReply) {
  try {
    reply.header('cache-control', 'no-store')
    reply.header('content-type', 'application/json; charset=utf-8')
    return createN8nForecastProof(request.body ?? {})
  } catch (error) {
    reply.code(400)
    return {
      verified: false,
      error: error instanceof Error ? error.message : 'Invalid n8n WorkPaper forecast request',
    }
  }
}

function handleN8nEvaluationRequest(request: FastifyRequest<{ Body: N8nWorkPaperEvaluationRequestBody }>, reply: FastifyReply) {
  try {
    reply.header('cache-control', 'no-store')
    reply.header('content-type', 'application/json; charset=utf-8')
    return createN8nWorkPaperEvaluationProof(request.body ?? {})
  } catch (error) {
    reply.code(400)
    return {
      verified: false,
      error: error instanceof Error ? error.message : 'Invalid n8n WorkPaper evaluation request',
    }
  }
}
