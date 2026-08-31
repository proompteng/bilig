import { createHmac } from 'node:crypto'
import { describe, expect, it } from 'vitest'
import { RequestAuthenticationError, createRequestSessionResolver, resolveRequestSession } from './session.js'

function createRequest(headers: Record<string, string>) {
  return {
    headers,
  }
}

describe('resolveRequestSession', () => {
  it('ignores caller-controlled identity headers in explicit demo mode', () => {
    const resolver = createRequestSessionResolver({
      env: {
        NODE_ENV: 'test',
        BILIG_AUTH_MODE: 'demo',
        BILIG_SESSION_SECRET: 'test-session-secret-that-is-at-least-32-bytes',
      },
      randomUUID: () => 'guest-id',
    })

    const session = resolveRequestSession(
      createRequest({
        authorization: 'Bearer admin@example.com',
        'x-bilig-user-id': 'admin@example.com',
        'x-auth-request-user': 'admin@example.com',
      }),
      resolver,
    )

    expect(session).toEqual({
      userId: 'guest:guest-id',
      roles: ['editor'],
      authSource: 'guest',
      isAuthenticated: false,
      setCookie: true,
    })
  })

  it('restores only a valid signed demo cookie', () => {
    const resolver = createRequestSessionResolver({
      env: {
        NODE_ENV: 'test',
        BILIG_AUTH_MODE: 'demo',
        BILIG_SESSION_SECRET: 'test-session-secret-that-is-at-least-32-bytes',
      },
      randomUUID: () => 'new-guest-id',
    })
    const first = resolveRequestSession(createRequest({}), resolver)
    const cookie = resolver.serializeCookie(first)

    const restored = resolveRequestSession(createRequest({ cookie }), resolver)

    expect(restored).toEqual({
      userId: 'guest:new-guest-id',
      roles: ['editor'],
      authSource: 'cookie',
      isAuthenticated: false,
      setCookie: false,
    })
  })

  it('replaces forged and malformed demo cookies without throwing', () => {
    const resolver = createRequestSessionResolver({
      env: {
        NODE_ENV: 'test',
        BILIG_AUTH_MODE: 'demo',
        BILIG_SESSION_SECRET: 'test-session-secret-that-is-at-least-32-bytes',
      },
      randomUUID: () => 'replacement-id',
    })

    for (const cookie of ['bilig_session=v1.Z3Vlc3Q6Zm9yZ2Vk.invalid', 'bilig_session=%E0%A4%A']) {
      expect(resolveRequestSession(createRequest({ cookie }), resolver)).toMatchObject({
        userId: 'guest:replacement-id',
        authSource: 'guest',
        setCookie: true,
      })
    }
  })

  it('accepts only fresh signed proxy assertions in authenticated mode', () => {
    const proxySecret = 'test-proxy-secret-that-is-at-least-32-bytes'
    const timestamp = '1735689600'
    const userId = 'alice@example.com'
    const roleHeader = 'editor,finance'
    const signature = createHmac('sha256', proxySecret).update(`${timestamp}\n${userId}\n${roleHeader}`).digest('base64url')
    const resolver = createRequestSessionResolver({
      env: {
        BILIG_AUTH_MODE: 'signed-proxy',
        BILIG_AUTH_PROXY_SECRET: proxySecret,
        BILIG_SESSION_SECRET: 'test-session-secret-that-is-at-least-32-bytes',
      },
      now: () => 1_735_689_600_000,
    })

    expect(
      resolveRequestSession(
        createRequest({
          'x-bilig-auth-user': userId,
          'x-bilig-auth-roles': roleHeader,
          'x-bilig-auth-timestamp': timestamp,
          'x-bilig-auth-signature': signature,
          'x-bilig-user-id': 'mallory@example.com',
        }),
        resolver,
      ),
    ).toEqual({
      userId,
      roles: ['editor', 'finance'],
      authSource: 'header',
      isAuthenticated: true,
      setCookie: false,
    })
  })

  it('rejects unsigned, forged, and stale proxy assertions', () => {
    const resolver = createRequestSessionResolver({
      env: {
        BILIG_AUTH_MODE: 'signed-proxy',
        BILIG_AUTH_PROXY_SECRET: 'test-proxy-secret-that-is-at-least-32-bytes',
        BILIG_SESSION_SECRET: 'test-session-secret-that-is-at-least-32-bytes',
      },
      now: () => 1_735_689_600_000,
    })

    expect(() => resolveRequestSession(createRequest({ 'x-bilig-user-id': 'alice@example.com' }), resolver)).toThrow(
      RequestAuthenticationError,
    )
    expect(() =>
      resolveRequestSession(
        createRequest({
          'x-bilig-auth-user': 'alice@example.com',
          'x-bilig-auth-roles': 'editor',
          'x-bilig-auth-timestamp': '1735689600',
          'x-bilig-auth-signature': 'forged',
        }),
        resolver,
      ),
    ).toThrow('Invalid signed proxy assertion')
  })

  it('requires an explicit authentication mode and strong secrets in production', () => {
    expect(() => createRequestSessionResolver({ env: { NODE_ENV: 'production' } })).toThrow(
      'BILIG_AUTH_MODE must be explicitly configured in production',
    )
    expect(() =>
      createRequestSessionResolver({
        env: {
          NODE_ENV: 'production',
          BILIG_AUTH_MODE: 'demo',
          BILIG_SESSION_SECRET: 'production-demo-secret-that-is-at-least-32-bytes',
        },
      }),
    ).toThrow('BILIG_AUTH_MODE=demo is not allowed in production; use signed-proxy')
    expect(() =>
      createRequestSessionResolver({
        env: {
          NODE_ENV: 'production',
          BILIG_AUTH_MODE: 'signed-proxy',
          BILIG_AUTH_PROXY_SECRET: 'production-proxy-secret-that-is-at-least-32-bytes',
          BILIG_SESSION_SECRET: 'short',
        },
      }),
    ).toThrow('BILIG_SESSION_SECRET must contain at least 32 bytes')
  })

  it('allows explicit demo mode for local and E2E environments', () => {
    for (const nodeEnv of ['development', 'test']) {
      const resolver = createRequestSessionResolver({
        env: {
          NODE_ENV: nodeEnv,
          BILIG_AUTH_MODE: 'demo',
          BILIG_SESSION_SECRET: 'local-demo-secret-that-is-at-least-32-bytes',
        },
      })

      expect(resolver.mode).toBe('demo')
    }
  })

  it('rejects demo mode outside explicitly local environments', () => {
    for (const nodeEnv of [undefined, '', ' ', 'staging', ' STAGING ', 'development ']) {
      expect(() =>
        createRequestSessionResolver({
          env: {
            NODE_ENV: nodeEnv,
            BILIG_AUTH_MODE: 'demo',
            BILIG_SESSION_SECRET: 'local-demo-secret-that-is-at-least-32-bytes',
          },
        }),
      ).toThrow('BILIG_AUTH_MODE=demo is only allowed when NODE_ENV is explicitly "development" or "test"')
    }
  })

  it('rejects an implicit demo mode without an explicitly local environment', () => {
    expect(() =>
      createRequestSessionResolver({
        env: {
          BILIG_SESSION_SECRET: 'local-demo-secret-that-is-at-least-32-bytes',
        },
      }),
    ).toThrow('BILIG_AUTH_MODE must be explicitly configured outside development and test')
  })

  it('rejects empty and malformed authentication mode values', () => {
    for (const authMode of ['', ' ', ' DEMO ', 'unknown']) {
      expect(() =>
        createRequestSessionResolver({
          env: {
            NODE_ENV: 'test',
            BILIG_AUTH_MODE: authMode,
            BILIG_SESSION_SECRET: 'local-demo-secret-that-is-at-least-32-bytes',
          },
        }),
      ).toThrow(`BILIG_AUTH_MODE must be "demo" or "signed-proxy", got ${authMode}`)
    }
  })
})
