import type { FastifyInstance } from 'fastify'
import httpProxy from '@fastify/http-proxy'
import { createErrorEnvelope } from '@bilig/runtime-kernel'

function resolveZeroKeepaliveUrl(upstream: string): URL {
  const url = new URL(upstream)
  url.pathname = '/keepalive'
  url.search = ''
  url.hash = ''
  return url
}

export function registerSyncServerZeroProxyRoutes(app: FastifyInstance, zeroProxyUpstream: string): void {
  app.route({
    method: ['GET', 'HEAD'],
    url: '/zero/keepalive',
    async handler(request, reply) {
      try {
        const upstreamResponse = await fetch(resolveZeroKeepaliveUrl(zeroProxyUpstream), {
          method: request.method,
          cache: 'no-store',
          signal: AbortSignal.timeout(2_000),
        })
        reply.code(upstreamResponse.status)
        reply.header('cache-control', 'no-store')
        const contentType = upstreamResponse.headers.get('content-type')
        if (contentType) {
          reply.header('content-type', contentType)
        }
        if (request.method === 'HEAD') {
          return reply.send()
        }
        return Buffer.from(await upstreamResponse.arrayBuffer())
      } catch {
        reply.code(503)
        reply.header('cache-control', 'no-store')
        return createErrorEnvelope('ZERO_CACHE_UNAVAILABLE', 'Zero cache keepalive probe failed', true)
      }
    },
  })
  app.get('/zero', async (_request, reply) => {
    return reply.redirect('/zero/')
  })
  app.register(httpProxy, {
    upstream: zeroProxyUpstream,
    prefix: '/zero/',
    rewritePrefix: '/',
    websocket: true,
    http2: false,
  })
}
