import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { FastifyInstance } from 'fastify'
import fastifyStatic from '@fastify/static'
import { createErrorEnvelope } from '@bilig/runtime-kernel'

const SPA_FALLBACK_PREFIXES = ['/api/', '/v1/', '/v2/', '/zero', '/healthz', '/readyz', '/runtime-config.json'] as const
const CANONICAL_DOCS_SITE_ROOT = 'https://proompteng.github.io/bilig/'
const DOCS_REDIRECT_EXACT_PATHS = new Set(['/AGENTS.md', '/agent.json', '/llms-full.txt', '/llms.txt', '/skill.txt'])
const DOCS_REDIRECT_EXTENSIONS = new Set(['.html', '.md', '.ts', '.txt'])

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

export function resolveCanonicalDocsRedirectUrl(method: string, url: string): string | null {
  if (method !== 'GET' && method !== 'HEAD') {
    return null
  }

  const requestUrl = new URL(url, 'https://bilig.proompteng.ai')
  const pathname = requestUrl.pathname
  if (
    pathname === '/' ||
    pathname.startsWith('/assets/') ||
    SPA_FALLBACK_PREFIXES.some((prefix) => pathname === prefix || pathname.startsWith(prefix))
  ) {
    return null
  }

  const lastSegment = pathname.slice(pathname.lastIndexOf('/') + 1)
  const extensionStart = lastSegment.lastIndexOf('.')
  const extension = extensionStart === -1 ? '' : lastSegment.slice(extensionStart)
  const isDocsPath =
    DOCS_REDIRECT_EXACT_PATHS.has(pathname) || pathname.startsWith('/.well-known/') || DOCS_REDIRECT_EXTENSIONS.has(extension)
  if (!isDocsPath) {
    return null
  }

  const target = new URL(pathname.replace(/^\/+/, ''), CANONICAL_DOCS_SITE_ROOT)
  target.search = requestUrl.search
  return target.toString()
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
    const docsRedirectUrl = resolveCanonicalDocsRedirectUrl(request.method, request.url)
    if (docsRedirectUrl !== null) {
      return reply.redirect(docsRedirectUrl, 302)
    }

    if (!shouldServeSpaFallback(request.method, request.url)) {
      reply.code(404)
      return createErrorEnvelope('NOT_FOUND', 'Route not found', false)
    }

    reply.header('cache-control', 'no-store')
    return reply.sendFile('index.html', { maxAge: 0, immutable: false })
  })
}
