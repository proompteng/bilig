import { randomUUID } from 'node:crypto'
import type { FastifyReply, FastifyRequest } from 'fastify'

const SESSION_COOKIE_NAME = 'bilig_user_id'
const SESSION_COOKIE_MAX_AGE = 60 * 60 * 24 * 365

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

function parseCookieHeader(header: string | undefined): ReadonlyMap<string, string> {
  if (!header) {
    return new Map()
  }
  return new Map(
    header
      .split(';')
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0)
      .map((entry) => {
        const separator = entry.indexOf('=')
        if (separator < 0) {
          return [entry, ''] as const
        }
        return [entry.slice(0, separator), decodeURIComponent(entry.slice(separator + 1))] as const
      }),
  )
}

function firstHeaderValue(value: string | string[] | undefined): string | undefined {
  if (Array.isArray(value)) {
    return value.find((entry) => entry.length > 0)
  }
  return typeof value === 'string' && value.length > 0 ? value : undefined
}

function parseRoleHeader(value: string | undefined): string[] {
  if (!value) {
    return []
  }
  return value
    .split(',')
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
}

export function resolveRequestSession(request: FastifyRequest): BiligRequestSession {
  const cookieMap = parseCookieHeader(firstHeaderValue(request.headers.cookie))
  const headerUserId =
    firstHeaderValue(request.headers['x-bilig-user-id']) ??
    firstHeaderValue(request.headers['x-forwarded-user']) ??
    firstHeaderValue(request.headers['x-auth-request-user'])
  const cookieUserId = cookieMap.get(SESSION_COOKIE_NAME)

  if (headerUserId) {
    const roles = parseRoleHeader(
      firstHeaderValue(request.headers['x-bilig-user-roles']) ??
        firstHeaderValue(request.headers['x-forwarded-groups']) ??
        firstHeaderValue(request.headers['x-auth-request-groups']),
    )
    return {
      userId: headerUserId,
      roles: roles.length > 0 ? roles : ['editor'],
      authSource: 'header',
      isAuthenticated: true,
      setCookie: cookieUserId !== headerUserId,
    }
  }

  if (cookieUserId) {
    return {
      userId: cookieUserId,
      roles: ['editor'],
      authSource: 'cookie',
      isAuthenticated: false,
      setCookie: false,
    }
  }

  return {
    userId: `guest:${randomUUID()}`,
    roles: ['editor'],
    authSource: 'guest',
    isAuthenticated: false,
    setCookie: true,
  }
}

function persistRequestSession(reply: FastifyReply, session: BiligRequestSession): void {
  if (!session.setCookie) {
    return
  }
  reply.header(
    'set-cookie',
    `${SESSION_COOKIE_NAME}=${encodeURIComponent(session.userId)}; Path=/; Max-Age=${SESSION_COOKIE_MAX_AGE}; SameSite=Lax`,
  )
}

export function resolveSessionIdentity(request: FastifyRequest, reply?: FastifyReply): SessionIdentity {
  const session = resolveRequestSession(request)
  if (reply) {
    persistRequestSession(reply, session)
  }
  return {
    userID: session.userId,
    roles: session.roles,
  }
}
