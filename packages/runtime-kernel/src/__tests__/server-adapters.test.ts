import { describe, expect, it } from 'vitest'

import {
  createErrorEnvelope,
  createGuestRuntimeSession,
  normalizeWebSocket,
  resolveRequestBaseUrl,
  resolveServerRuntimeConfig,
  toMessageBytes,
} from '../index.js'

describe('runtime-kernel server adapters', () => {
  it('reads runtime config from env once and trims values', () => {
    expect(
      resolveServerRuntimeConfig({
        BILIG_WEB_APP_BASE_URL: ' https://app.example.com/ ',
        BILIG_CORS_ORIGIN: ' https://origin.example.com ',
      }),
    ).toEqual({
      browserAppBaseUrl: 'https://app.example.com/',
      corsOrigin: 'https://origin.example.com',
    })
  })

  it('creates guest runtime sessions', () => {
    expect(createGuestRuntimeSession('guest:test')).toEqual({
      authToken: 'guest:test',
      userId: 'guest:test',
      roles: ['editor'],
      isAuthenticated: false,
      authSource: 'guest',
    })
  })

  it('normalizes request base urls', () => {
    expect(
      resolveRequestBaseUrl(
        {
          protocol: 'https',
          headers: {
            host: ['sheet.example.com', 'ignored.example.com'],
          },
        },
        '127.0.0.1:4321',
      ),
    ).toBe('https://sheet.example.com')
  })

  it('converts websocket payloads into Uint8Array', () => {
    const bytes = toMessageBytes(Buffer.from([1, 2, 3]))
    expect([...bytes]).toEqual([1, 2, 3])
  })

  it('normalizes event-target websocket wrappers', () => {
    const listeners = new Map<string, (payload: unknown) => void>()
    const sent: Uint8Array[] = []
    const socket = normalizeWebSocket({
      addEventListener(event: string, listener: (payload: unknown) => void) {
        listeners.set(event, listener)
      },
      send(data: Uint8Array) {
        sent.push(data)
      },
    })

    const messages: unknown[] = []
    socket.on('message', (payload) => {
      messages.push(payload)
    })
    listeners.get('message')?.({
      data: new Uint8Array([9, 8, 7]),
    })
    socket.send(new Uint8Array([1, 2]))

    expect(messages).toEqual([new Uint8Array([9, 8, 7])])
    expect(sent).toEqual([new Uint8Array([1, 2])])
  })

  it('creates error envelopes', () => {
    expect(createErrorEnvelope('BROKEN', 'broken', true)).toEqual({
      error: 'BROKEN',
      message: 'broken',
      retryable: true,
    })
  })
})
