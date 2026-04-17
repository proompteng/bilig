import { Buffer } from 'node:buffer'
import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { normalizeBaseUrl, normalizeWebSocket, resolveRequestBaseUrl, resolveServerRuntimeConfig, toMessageBytes } from '../index.js'

describe('runtime-kernel server adapter fuzz', () => {
  it('should normalize env-driven runtime config and request base urls consistently', async () => {
    await runProperty({
      suite: 'runtime-kernel/server-adapters/url-normalization',
      arbitrary: fc.record({
        protocol: fc.constantFrom('http', 'https'),
        host: fc.constantFrom('sheet.example.com', '127.0.0.1:4321', 'api.internal'),
        includeTrailingSlash: fc.boolean(),
        includeCorsOrigin: fc.boolean(),
      }),
      predicate: async ({ protocol, host, includeTrailingSlash, includeCorsOrigin }) => {
        const rawBaseUrl = `${protocol}://${host}${includeTrailingSlash ? '/' : ''}`
        expect(normalizeBaseUrl(rawBaseUrl)).toBe(`${protocol}://${host}`)
        expect(
          resolveRequestBaseUrl(
            {
              protocol,
              headers: {
                host: [host, 'ignored.example.com'],
              },
            },
            'fallback.example.com',
          ),
        ).toBe(`${protocol}://${host}`)

        expect(
          resolveServerRuntimeConfig({
            BILIG_WEB_APP_BASE_URL: ` ${rawBaseUrl} `,
            BILIG_CORS_ORIGIN: includeCorsOrigin ? ` ${protocol}://origin.example.com ` : ' ',
          }),
        ).toEqual({
          browserAppBaseUrl: rawBaseUrl,
          ...(includeCorsOrigin ? { corsOrigin: `${protocol}://origin.example.com` } : {}),
        })
      },
    })
  })

  it('should preserve websocket payload bytes across supported transport wrappers', async () => {
    await runProperty({
      suite: 'runtime-kernel/server-adapters/message-bytes',
      arbitrary: fc.record({
        bytes: fc.uint8Array({ maxLength: 32 }),
        start: fc.integer({ min: 0, max: 4 }),
      }),
      predicate: async ({ bytes, start }) => {
        const offset = Math.min(start, bytes.length)
        const view = bytes.subarray(offset)
        const arrayBuffer = view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength)
        const dataView = new DataView(arrayBuffer)

        expect(toMessageBytes(Buffer.from(view))).toEqual(view)
        expect(toMessageBytes(arrayBuffer)).toEqual(view)
        expect(toMessageBytes(dataView)).toEqual(view)

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

        const received: unknown[] = []
        socket.on('message', (payload) => {
          received.push(payload)
        })
        listeners.get('message')?.({ data: view })
        socket.send(view)

        expect(received).toEqual([view])
        expect(sent).toEqual([view])
      },
    })
  })
})
