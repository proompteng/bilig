import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import { WORKPAPER_VERSION } from '@bilig/headless'
import {
  WORKPAPER_MCP_PROTOCOL_VERSION,
  WORKPAPER_MCP_SUPPORTED_PROTOCOL_VERSIONS,
  buildDemoWorkPaper,
  createFileBackedWorkPaperMcpToolServer,
  createWorkPaperMcpJsonRpcError,
  dispatchWorkPaperMcpJsonRpc,
  isWorkPaperMcpProtocolVersion,
} from '@bilig/headless/mcp'

const MCP_ENDPOINTS = ['/mcp', '/mcp/workpaper'] as const
const MCP_ALLOWED_METHODS = 'POST, GET, DELETE, OPTIONS'
const DEFAULT_ALLOWED_ORIGINS = ['https://claude.ai', 'https://www.claude.ai', 'https://claude.com', 'https://www.claude.com'] as const

interface WorkPaperMcpRemoteRouteOptions {
  env?: Record<string, string | undefined>
}

function createRequestServer() {
  return createFileBackedWorkPaperMcpToolServer({
    workbook: buildDemoWorkPaper(),
    writable: false,
  })
}

function handleWorkPaperMcpPost(request: FastifyRequest<{ Body: unknown }>, reply: FastifyReply, env: Record<string, string | undefined>) {
  const originStatus = applyCorsHeaders(request, reply, env)
  if (originStatus === 'forbidden') {
    return sendHttpJsonRpcError(reply, 403, -32000, 'Forbidden Origin header')
  }

  if (!isSupportedAcceptHeader(request.headers.accept)) {
    return sendHttpJsonRpcError(reply, 406, -32000, 'Accept header must allow application/json responses')
  }

  const protocolVersion = readProtocolVersionHeader(request.headers['mcp-protocol-version'])
  if (protocolVersion === null) {
    return sendHttpJsonRpcError(
      reply,
      400,
      -32000,
      `Unsupported MCP-Protocol-Version. Supported versions: ${WORKPAPER_MCP_SUPPORTED_PROTOCOL_VERSIONS.join(', ')}`,
    )
  }

  const result = dispatchWorkPaperMcpJsonRpc(request.body, {
    server: createRequestServer(),
    protocolVersion,
    serverName: 'bilig-workpaper-remote-demo',
    serverTitle: 'Bilig WorkPaper Remote Demo',
    serverVersion: WORKPAPER_VERSION,
  })

  applyCommonMcpHeaders(reply, protocolVersion)
  if (result.kind === 'notification') {
    return reply.code(202).send()
  }

  reply.header('content-type', 'application/json; charset=utf-8')
  return result.response
}

function handleWorkPaperMcpOptions(request: FastifyRequest, reply: FastifyReply, env: Record<string, string | undefined>) {
  const originStatus = applyCorsHeaders(request, reply, env)
  if (originStatus === 'forbidden') {
    return sendHttpJsonRpcError(reply, 403, -32000, 'Forbidden Origin header')
  }
  reply.header('access-control-allow-methods', MCP_ALLOWED_METHODS)
  reply.header('access-control-allow-headers', 'accept, content-type, mcp-protocol-version, mcp-session-id')
  reply.header('access-control-max-age', '600')
  return reply.code(204).send()
}

function handleWorkPaperMcpUnsupportedMethod(request: FastifyRequest, reply: FastifyReply, env: Record<string, string | undefined>) {
  const originStatus = applyCorsHeaders(request, reply, env)
  if (originStatus === 'forbidden') {
    return sendHttpJsonRpcError(reply, 403, -32000, 'Forbidden Origin header')
  }
  applyCommonMcpHeaders(reply, WORKPAPER_MCP_PROTOCOL_VERSION)
  reply.header('allow', MCP_ALLOWED_METHODS)
  return sendHttpJsonRpcError(reply, 405, -32000, 'Method not allowed; this stateless endpoint returns JSON over POST')
}

function sendHttpJsonRpcError(reply: FastifyReply, statusCode: number, code: number, message: string) {
  reply.code(statusCode)
  reply.header('cache-control', 'no-store')
  reply.header('content-type', 'application/json; charset=utf-8')
  return createWorkPaperMcpJsonRpcError(null, code, message)
}

function applyCommonMcpHeaders(reply: FastifyReply, protocolVersion: string): void {
  reply.header('cache-control', 'no-store')
  reply.header('mcp-protocol-version', protocolVersion)
}

function applyCorsHeaders(request: FastifyRequest, reply: FastifyReply, env: Record<string, string | undefined>): 'allowed' | 'forbidden' {
  const origin = readSingleHeader(request.headers.origin)
  if (!origin) {
    return 'allowed'
  }

  if (!isAllowedOrigin(origin, env)) {
    return 'forbidden'
  }

  reply.header('access-control-allow-origin', origin)
  reply.header('vary', 'Origin')
  return 'allowed'
}

function isAllowedOrigin(origin: string, env: Record<string, string | undefined>): boolean {
  if (isLocalOrigin(origin)) {
    return true
  }

  return new Set([...DEFAULT_ALLOWED_ORIGINS, ...parseAllowedOrigins(env['BILIG_REMOTE_MCP_ALLOWED_ORIGINS'])]).has(origin)
}

function parseAllowedOrigins(value: string | undefined): string[] {
  return (value ?? '')
    .split(',')
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
}

function isLocalOrigin(origin: string): boolean {
  try {
    const url = new URL(origin)
    return (
      (url.protocol === 'http:' || url.protocol === 'https:') &&
      (url.hostname === 'localhost' || url.hostname === '127.0.0.1' || url.hostname === '[::1]')
    )
  } catch {
    return false
  }
}

function readProtocolVersionHeader(value: string | string[] | undefined): string | null {
  const header = readSingleHeader(value)?.trim()
  if (!header) {
    return '2025-03-26'
  }
  return isWorkPaperMcpProtocolVersion(header) ? header : null
}

function isSupportedAcceptHeader(value: string | string[] | undefined): boolean {
  const accept = readSingleHeader(value)
  if (!accept) {
    return true
  }

  const acceptedTypes = accept
    .split(',')
    .map((entry) => entry.split(';', 1)[0]?.trim().toLowerCase() ?? '')
    .filter((entry) => entry.length > 0)
  const acceptedTypeSet = new Set(acceptedTypes)
  return acceptedTypeSet.has('*/*') || acceptedTypeSet.has('application/*') || acceptedTypeSet.has('application/json')
}

function readSingleHeader(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value
}

export function registerWorkPaperMcpRemoteRoutes(app: FastifyInstance, options: WorkPaperMcpRemoteRouteOptions = {}): void {
  const env = options.env ?? process.env
  for (const endpoint of MCP_ENDPOINTS) {
    app.options(endpoint, async (request, reply) => handleWorkPaperMcpOptions(request, reply, env))
    app.get(endpoint, async (request, reply) => handleWorkPaperMcpUnsupportedMethod(request, reply, env))
    app.delete(endpoint, async (request, reply) => handleWorkPaperMcpUnsupportedMethod(request, reply, env))
    app.post(endpoint, async (request: FastifyRequest<{ Body: unknown }>, reply) => handleWorkPaperMcpPost(request, reply, env))
  }
}
