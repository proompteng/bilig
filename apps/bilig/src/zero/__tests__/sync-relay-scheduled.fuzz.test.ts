import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import { decodeFrame, encodeFrame, type ProtocolFrame } from '@bilig/binary-protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { runScheduledProperty } from '@bilig/test-fuzz'
import { createHttpSyncRelay } from '../sync-relay.js'

type RelayAction = { kind: 'send'; batchId: string; counter: number } | { kind: 'disconnect' }

describe('sync relay scheduled fuzz', () => {
  it('should serialize send and disconnect operations into a coherent reconnect sequence under scheduled fetch timing', async () => {
    await runScheduledProperty({
      suite: 'bilig/sync-relay/scheduled-reconnect-ordering',
      arbitrary: fc.array(relayActionArbitrary, { minLength: 4, maxLength: 16 }),
      predicate: async ({ scheduler, value: actions }) => {
        const seenFrames: ProtocolFrame[] = []
        const executedActions: RelayAction[] = []
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
            const nextFrame =
              frame.kind === 'hello'
                ? {
                    kind: 'cursorWatermark' as const,
                    documentId: 'doc-1',
                    cursor: serverCursor,
                    compactedCursor: serverCursor,
                  }
                : {
                    kind: 'ack' as const,
                    documentId: 'doc-1',
                    batchId: frame.kind === 'appendBatch' ? frame.batch.id : 'missing',
                    cursor: (serverCursor += 1),
                    acceptedAtUnixMs: serverCursor,
                  }
            return new Response(Buffer.from(encodeFrame(nextFrame)))
          },
        })
        const actionPromises = actions.map((action, index) =>
          scheduler.schedule(
            Promise.resolve().then(async () => {
              executedActions.push(action)
              if (action.kind === 'disconnect') {
                await relay.disconnect()
                return
              }
              await relay.send({
                id: action.batchId,
                replicaId: 'worksheet-host:doc-1',
                clock: { counter: action.counter },
                ops: [],
              } satisfies EngineOpBatch)
              return undefined
            }),
            `relay-action-${index}`,
          ),
        )

        await scheduler.waitFor(Promise.all(actionPromises))

        const { expectedBatchIds, expectedHelloCursors } = replayRelayActions(executedActions)
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
      counter: fc.integer({ min: 1, max: 1_000 }),
    })
    .map((action) => ({
      kind: 'send' as const,
      batchId: action.batchId,
      counter: action.counter,
    })),
  fc.constant({ kind: 'disconnect' as const }),
)

function replayRelayActions(actions: readonly RelayAction[]): {
  expectedHelloCursors: number[]
  expectedBatchIds: string[]
} {
  const expectedHelloCursors: number[] = []
  const expectedBatchIds: string[] = []
  let connected = false
  let expectedCursor = 0

  actions.forEach((action) => {
    if (action.kind === 'disconnect') {
      connected = false
      return
    }
    if (!connected) {
      expectedHelloCursors.push(expectedCursor)
      connected = true
    }
    expectedBatchIds.push(action.batchId)
    expectedCursor += 1
  })

  return { expectedHelloCursors, expectedBatchIds }
}
