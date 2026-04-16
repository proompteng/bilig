import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import { decodeFrame, encodeFrame, type ProtocolFrame } from '@bilig/binary-protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { runProperty } from '@bilig/test-fuzz'
import { createHttpSyncRelay } from '../sync-relay.js'

type RelayAction = { kind: 'send'; batchId: string; counter: number } | { kind: 'disconnect' }

describe('sync relay fuzz', () => {
  it('preserves hello segmentation, server cursor carry-forward, and append order across reconnects', async () => {
    await runProperty({
      suite: 'bilig/sync-relay/reconnect-ordering',
      arbitrary: fc.array(relayActionArbitrary, { minLength: 4, maxLength: 20 }),
      predicate: async (actions) => {
        const seenFrames: ProtocolFrame[] = []
        let serverCursor = 0
        const relay = createHttpSyncRelay({
          documentId: 'doc-1',
          baseUrl: 'https://bilig.proompteng.ai',
          fetchImpl: async (_url, init) => {
            if (!(init?.body instanceof Uint8Array)) {
              throw new TypeError('expected binary request body')
            }
            const frame = decodeFrame(init.body)
            seenFrames.push(frame)
            if (frame.kind === 'hello') {
              return new Response(
                Buffer.from(
                  encodeFrame({
                    kind: 'cursorWatermark',
                    documentId: 'doc-1',
                    cursor: serverCursor,
                    compactedCursor: serverCursor,
                  }),
                ),
              )
            }
            serverCursor += 1
            return new Response(
              Buffer.from(
                encodeFrame({
                  kind: 'ack',
                  documentId: 'doc-1',
                  batchId: frame.kind === 'appendBatch' ? frame.batch.id : 'missing',
                  cursor: serverCursor,
                  acceptedAtUnixMs: serverCursor,
                }),
              ),
            )
          },
        })

        const expectedHelloCursors: number[] = []
        const expectedBatchIds: string[] = []
        let connected = false
        let expectedCursor = 0

        await actions.reduce<Promise<void>>(async (previous, action) => {
          await previous
          if (action.kind === 'disconnect') {
            await relay.disconnect()
            connected = false
            return
          }
          if (!connected) {
            expectedHelloCursors.push(expectedCursor)
            connected = true
          }
          expectedBatchIds.push(action.batchId)
          await relay.send({
            id: action.batchId,
            replicaId: 'worksheet-host:doc-1',
            clock: { counter: action.counter },
            ops: [],
          } satisfies EngineOpBatch)
          expectedCursor += 1
        }, Promise.resolve())

        const helloFrames = seenFrames.filter((frame) => frame.kind === 'hello')
        const appendFrames = seenFrames.filter((frame) => frame.kind === 'appendBatch')
        expect(helloFrames.map((frame) => frame.lastServerCursor)).toEqual(expectedHelloCursors)
        expect(appendFrames.map((frame) => frame.batch.id)).toEqual(expectedBatchIds)
      },
    })
  })
})

// Helpers

const relayActionArbitrary = fc.oneof<RelayAction>(
  fc
    .record({
      batchId: fc.uuid(),
      counter: fc.integer({ min: 1, max: 1000 }),
    })
    .map((action) => Object.assign({ kind: 'send' as const }, action)),
  fc.constant({ kind: 'disconnect' }),
)
