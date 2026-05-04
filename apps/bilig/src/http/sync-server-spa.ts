import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { FastifyInstance } from 'fastify'
import fastifyStatic from '@fastify/static'
import { createErrorEnvelope } from '@bilig/runtime-kernel'

const SPA_FALLBACK_PREFIXES = ['/api/', '/v1/', '/v2/', '/zero', '/healthz', '/runtime-config.json'] as const

export function resolveSyncServerWebDistRoot(): string | null {
  const candidate = join(dirname(fileURLToPath(import.meta.url)), '../../public')
  return existsSync(candidate) ? candidate : null
}

function shouldServeSpaFallback(method: string, url: string): boolean {
  if (method !== 'GET' && method !== 'HEAD') {
    return false
  }

  const pathname = url.split('?', 1)[0] ?? url
  if (pathname.includes('.', pathname.lastIndexOf('/') + 1)) {
    return false
  }

  return !SPA_FALLBACK_PREFIXES.some((prefix) => pathname === prefix || pathname.startsWith(prefix))
}

export function registerSyncServerSpaRoutes(app: FastifyInstance, webDistRoot: string | null): void {
  if (!webDistRoot) {
    return
  }

  app.register(fastifyStatic, {
    root: webDistRoot,
    prefix: '/',
    maxAge: '30d',
    immutable: true,
  })

  app.get('/', async (_request, reply) => {
    reply.header('cache-control', 'no-store')
    return reply.sendFile('index.html', { maxAge: 0, immutable: false })
  })

  app.setNotFoundHandler(async (request, reply) => {
    if (!shouldServeSpaFallback(request.method, request.url)) {
      reply.code(404)
      return createErrorEnvelope('NOT_FOUND', 'Route not found', false)
    }

    reply.header('cache-control', 'no-store')
    return reply.sendFile('index.html', { maxAge: 0, immutable: false })
  })
}
