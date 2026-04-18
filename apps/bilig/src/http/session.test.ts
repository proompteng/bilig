import { describe, expect, it } from 'vitest'
import { resolveRequestSession } from './session.js'

function createRequest(headers: Record<string, string>) {
  return {
    headers,
  }
}

describe('resolveRequestSession', () => {
  it('prefers the bearer auth token over a stale bilig session cookie', () => {
    const session = resolveRequestSession(
      createRequest({
        authorization: 'Bearer guest:zero-user',
        cookie: 'bilig_user_id=guest:stale-cookie-user',
      }),
    )

    expect(session).toMatchObject({
      userId: 'guest:zero-user',
      authSource: 'header',
      isAuthenticated: false,
      setCookie: true,
    })
  })

  it('still prefers explicit forwarded user headers over the bearer auth token', () => {
    const session = resolveRequestSession(
      createRequest({
        authorization: 'Bearer guest:zero-user',
        'x-bilig-user-id': 'alice@example.com',
      }),
    )

    expect(session).toMatchObject({
      userId: 'alice@example.com',
      authSource: 'header',
      isAuthenticated: true,
      setCookie: true,
    })
  })
})
