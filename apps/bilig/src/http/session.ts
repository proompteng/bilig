import { createHmac, randomUUID, timingSafeEqual } from 'node:crypto'
import type { FastifyReply } from 'fastify'

const SESSION_COOKIE_NAME = 'bilig_session'
const SESSION_COOKIE_MAX_AGE = 60 * 60 * 24 * 365
const MINIMUM_SECRET_BYTES = 32
const PROXY_ASSERTION_MAX_AGE_SECONDS = 60
const DEVELOPMENT_SESSION_SECRET = 'bilig-development-session-secret-do-not-use-in-production'

export type BiligAuthMode = 'demo' | 'signed-proxy'

export interface BiligRequestSession {
  userId: string
  roles: string[]
  authSource: 'header' | 'cookie' | 'guest'
  isAuthenticated: boolean
  setCookie: boolean
}

export interface SessionIdentity {
  userID: string
  roles: string[]
}

interface HeadersRequestLike {
  readonly headers: {
    readonly [key: string]: string | string[] | undefined
    readonly cookie?: string | string[] | undefined
  }
}

export interface RequestSessionResolver {
  readonly mode: BiligAuthMode
  resolve(request: HeadersRequestLike): BiligRequestSession
  persist(reply: FastifyReply, session: BiligRequestSession): void
  serializeCookie(session: BiligRequestSession): string
}

interface CreateRequestSessionResolverOptions {
  readonly env?: Readonly<Record<string, string | undefined>>
  readonly now?: () => number
  readonly randomUUID?: () => string
}

export class RequestAuthenticationError extends Error {
  readonly statusCode = 401
  readonly code = 'REQUEST_AUTHENTICATION_REQUIRED'

  constructor(message: string) {
    super(message)
    this.name = 'RequestAuthenticationError'
  }
}

function parseCookieHeader(header: string | undefined): ReadonlyMap<string, string> {
  if (!header) {
    return new Map()
  }
  const cookies = new Map<string, string>()
  for (const rawEntry of header.split(';')) {
    const entry = rawEntry.trim()
    if (entry.length === 0) {
      continue
    }
    const separator = entry.indexOf('=')
    const name = separator < 0 ? entry : entry.slice(0, separator)
    const rawValue = separator < 0 ? '' : entry.slice(separator + 1)
    try {
      cookies.set(name, decodeURIComponent(rawValue))
    } catch {
      cookies.set(name, '')
    }
  }
  return cookies
}

function firstHeaderValue(value: string | string[] | undefined): string | undefined {
  if (Array.isArray(value)) {
    return value.find((entry) => entry.length > 0)
  }
  return typeof value === 'string' && value.length > 0 ? value : undefined
}

function parseRoleHeader(value: string): string[] {
  const roles = value
    .split(',')
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
  if (roles.length === 0 || roles.some((role) => role.length > 128)) {
    throw new RequestAuthenticationError('Invalid signed proxy assertion')
  }
  return [...new Set(roles)]
}

function requireSecret(value: string | undefined, name: string, fallback?: string): string {
  const secret = value ?? fallback
  if (!secret || Buffer.byteLength(secret) < MINIMUM_SECRET_BYTES) {
    throw new Error(`${name} must contain at least ${String(MINIMUM_SECRET_BYTES)} bytes`)
  }
  return secret
}

function parseBoolean(value: string | undefined, fallback: boolean, name: string): boolean {
  if (value === undefined) {
    return fallback
  }
  if (value === 'true' || value === '1') {
    return true
  }
  if (value === 'false' || value === '0') {
    return false
  }
  throw new Error(`${name} must be "1", "true", "0", or "false" when set`)
}

function resolveAuthMode(env: Readonly<Record<string, string | undefined>>): BiligAuthMode {
  const configured = env['BILIG_AUTH_MODE']
  const nodeEnv = env['NODE_ENV']
  const isExplicitLocalEnvironment = nodeEnv === 'development' || nodeEnv === 'test'
  if (configured === undefined) {
    if (!isExplicitLocalEnvironment) {
      if (nodeEnv === 'production') {
        throw new Error('BILIG_AUTH_MODE must be explicitly configured in production')
      }
      throw new Error('BILIG_AUTH_MODE must be explicitly configured outside development and test')
    }
    return 'demo'
  }
  if (configured === 'demo' && !isExplicitLocalEnvironment) {
    if (nodeEnv === 'production') {
      throw new Error('BILIG_AUTH_MODE=demo is not allowed in production; use signed-proxy')
    }
    throw new Error('BILIG_AUTH_MODE=demo is only allowed when NODE_ENV is explicitly "development" or "test"')
  }
  if (configured !== 'demo' && configured !== 'signed-proxy') {
    throw new Error(`BILIG_AUTH_MODE must be "demo" or "signed-proxy", got ${configured}`)
  }
  return configured
}

function sign(secret: string, value: string): string {
  return createHmac('sha256', secret).update(value).digest('base64url')
}

function signaturesMatch(actual: string, expected: string): boolean {
  const actualBytes = Buffer.from(actual)
  const expectedBytes = Buffer.from(expected)
  return actualBytes.length === expectedBytes.length && timingSafeEqual(actualBytes, expectedBytes)
}

function encodeSessionCookieValue(userId: string, secret: string): string {
  const payload = Buffer.from(userId).toString('base64url')
  const unsigned = `v1.${payload}`
  return `${unsigned}.${sign(secret, unsigned)}`
}

function decodeSessionCookieValue(value: string | undefined, secret: string): string | undefined {
  if (!value) {
    return undefined
  }
  const parts = value.split('.')
  if (parts.length !== 3 || parts[0] !== 'v1' || !parts[1] || !parts[2]) {
    return undefined
  }
  const unsigned = `v1.${parts[1]}`
  if (!signaturesMatch(parts[2], sign(secret, unsigned))) {
    return undefined
  }
  try {
    const userId = Buffer.from(parts[1], 'base64url').toString('utf8')
    return /^guest:[A-Za-z0-9-]{1,128}$/u.test(userId) ? userId : undefined
  } catch {
    return undefined
  }
}

function resolveSignedProxySession(request: HeadersRequestLike, proxySecret: string, now: () => number): BiligRequestSession {
  const userId = firstHeaderValue(request.headers['x-bilig-auth-user'])?.trim()
  const roleHeader = firstHeaderValue(request.headers['x-bilig-auth-roles'])?.trim()
  const timestampHeader = firstHeaderValue(request.headers['x-bilig-auth-timestamp'])?.trim()
  const signature = firstHeaderValue(request.headers['x-bilig-auth-signature'])?.trim()
  if (!userId || userId.length > 256 || !roleHeader || !timestampHeader || !signature) {
    throw new RequestAuthenticationError('Missing signed proxy assertion')
  }
  if (!/^(?:0|[1-9]\d*)$/u.test(timestampHeader)) {
    throw new RequestAuthenticationError('Invalid signed proxy assertion')
  }
  const timestamp = Number(timestampHeader)
  const currentTimestamp = Math.floor(now() / 1_000)
  if (!Number.isSafeInteger(timestamp) || Math.abs(currentTimestamp - timestamp) > PROXY_ASSERTION_MAX_AGE_SECONDS) {
    throw new RequestAuthenticationError('Expired signed proxy assertion')
  }
  const expected = sign(proxySecret, `${timestampHeader}\n${userId}\n${roleHeader}`)
  if (!signaturesMatch(signature, expected)) {
    throw new RequestAuthenticationError('Invalid signed proxy assertion')
  }
  return {
    userId,
    roles: parseRoleHeader(roleHeader),
    authSource: 'header',
    isAuthenticated: true,
    setCookie: false,
  }
}

export function createRequestSessionResolver(options: CreateRequestSessionResolverOptions = {}): RequestSessionResolver {
  const env = options.env ?? process.env
  const mode = resolveAuthMode(env)
  const isProduction = env['NODE_ENV'] === 'production'
  const sessionSecret = requireSecret(
    env['BILIG_SESSION_SECRET'],
    'BILIG_SESSION_SECRET',
    isProduction ? undefined : DEVELOPMENT_SESSION_SECRET,
  )
  const proxySecret = mode === 'signed-proxy' ? requireSecret(env['BILIG_AUTH_PROXY_SECRET'], 'BILIG_AUTH_PROXY_SECRET') : undefined
  const secureCookie = parseBoolean(env['BILIG_SESSION_COOKIE_SECURE'], isProduction, 'BILIG_SESSION_COOKIE_SECURE')
  const now = options.now ?? Date.now
  const createUuid = options.randomUUID ?? randomUUID
  const sessionByRequest = new WeakMap<object, BiligRequestSession>()

  const serializeCookie = (session: BiligRequestSession): string => {
    const value = encodeSessionCookieValue(session.userId, sessionSecret)
    return `${SESSION_COOKIE_NAME}=${encodeURIComponent(value)}; Path=/; Max-Age=${SESSION_COOKIE_MAX_AGE}; HttpOnly; SameSite=Lax${secureCookie ? '; Secure' : ''}`
  }

  return {
    mode,
    resolve(request) {
      const cached = sessionByRequest.get(request)
      if (cached) {
        return cached
      }
      let session: BiligRequestSession
      if (mode === 'signed-proxy') {
        session = resolveSignedProxySession(request, proxySecret ?? '', now)
      } else {
        const cookieMap = parseCookieHeader(firstHeaderValue(request.headers.cookie))
        const cookieUserId = decodeSessionCookieValue(cookieMap.get(SESSION_COOKIE_NAME), sessionSecret)
        session = cookieUserId
          ? {
              userId: cookieUserId,
              roles: ['editor'],
              authSource: 'cookie',
              isAuthenticated: false,
              setCookie: false,
            }
          : {
              userId: `guest:${createUuid()}`,
              roles: ['editor'],
              authSource: 'guest',
              isAuthenticated: false,
              setCookie: true,
            }
      }
      sessionByRequest.set(request, session)
      return session
    },
    persist(reply, session) {
      if (session.setCookie) {
        reply.header('set-cookie', serializeCookie(session))
      }
    },
    serializeCookie,
  }
}

export function resolveRequestSession(
  request: HeadersRequestLike,
  resolver: RequestSessionResolver = createRequestSessionResolver(),
): BiligRequestSession {
  return resolver.resolve(request)
}

export function resolveSessionIdentity(
  request: HeadersRequestLike,
  reply?: FastifyReply,
  resolver: RequestSessionResolver = createRequestSessionResolver(),
): SessionIdentity {
  const session = resolver.resolve(request)
  if (reply) {
    resolver.persist(reply, session)
  }
  return {
    userID: session.userId,
    roles: session.roles,
  }
}
