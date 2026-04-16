import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import { decodeFrame, encodeFrame, type ProtocolFrame } from '@bilig/binary-protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { runProperty } from '@bilig/test-fuzz'
import { createHttpSyncRelay } from '../sync-relay.js'

type ReconnectReplayAction = { kind: 'send'; batchId: string; counter: number } | { kind: 'disconnect' }

describe('reconnect replay fuzz', () => {
  it('reconnects and replays local batches without duplicate server side effects', async () => {
    await runProperty({
      suite: 'bilig/sync-relay/reconnect-replay-convergence',
      arbitrary: fc.array(reconnectReplayActionArbitrary, { minLength: 4, maxLength: 20 }),
      predicate: async (actions) => {
        const seenFrames: ProtocolFrame[] = []
        const deliveredBatchIds: string[] = []
        const deliveredBatchIdSet = new Set<string>()
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
              if (frame.lastServerCursor !== serverCursor) {
                throw new Error(`Expected reconnect cursor ${serverCursor}, received ${frame.lastServerCursor}`)
              }
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
            if (frame.cursor !== serverCursor) {
              throw new Error(`Expected append cursor ${serverCursor}, received ${frame.cursor}`)
            }
            if (deliveredBatchIdSet.has(frame.batch.id)) {
              throw new Error(`Duplicate server side effect for batch ${frame.batch.id}`)
            }
            deliveredBatchIdSet.add(frame.batch.id)
            deliveredBatchIds.push(frame.batch.id)
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

        const expectedBatchIds: string[] = []
        await actions.reduce<Promise<void>>(async (previous, action) => {
          await previous
          if (action.kind === 'disconnect') {
            await relay.disconnect()
            return
          }
          expectedBatchIds.push(action.batchId)
          await relay.send({
            id: action.batchId,
            replicaId: 'worksheet-host:doc-1',
            clock: { counter: action.counter },
            ops: [],
          } satisfies EngineOpBatch)
        }, Promise.resolve())

        expect(deliveredBatchIds).toEqual(expectedBatchIds)
      },
    })
  })
})

const reconnectReplayActionArbitrary = fc.oneof<ReconnectReplayAction>(
  fc
    .record({
      batchId: fc.uuid(),
      counter: fc.integer({ min: 1, max: 1000 }),
    })
    .map((action) => Object.assign({ kind: 'send' as const }, action)),
  fc.constant({ kind: 'disconnect' as const }),
)
