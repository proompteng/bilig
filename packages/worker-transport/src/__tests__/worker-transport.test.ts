import { MessageChannel } from 'node:worker_threads'

import { describe, expect, it } from 'vitest'

import type { EngineEvent } from '@bilig/protocol'

import {
  createWorkerEngineClient,
  createWorkerEngineHost,
  decodeRenderTileDeltaBatch,
  decodeViewportPatch,
  decodeWorkbookDeltaBatchV3,
  encodeRenderTileDeltaBatch,
  encodeViewportPatch,
  encodeWorkbookDeltaBatchV3,
} from '../index.js'

async function waitFor(predicate: () => boolean, attempts = 20): Promise<void> {
  const poll = async (remainingAttempts: number): Promise<void> => {
    if (predicate()) {
      return
    }
    if (remainingAttempts <= 0) {
      throw new Error('Timed out waiting for worker transport condition')
    }
    await new Promise((resolve) => setTimeout(resolve, 0))
    await poll(remainingAttempts - 1)
  }

  await poll(attempts)
}

describe('worker transport', () => {
  it('invokes engine methods across a message channel', async () => {
    const channel = new MessageChannel()
    const host = createWorkerEngineHost(
      {
        async ready() {
          return
        },
        add(left: number, right: number) {
          return left + right
        },
      },
      channel.port1,
    )

    const client = createWorkerEngineClient({ port: channel.port2 })

    await expect(client.ready()).resolves.toBeUndefined()
    await expect(client.invoke('add', 2, 5)).resolves.toBe(7)

    client.dispose()
    host.dispose()
  })

  it('preserves engine method context for class instances', async () => {
    class CounterEngine {
      private value = 4

      readValue(): number {
        return this.value
      }
    }

    const channel = new MessageChannel()
    const host = createWorkerEngineHost(new CounterEngine(), channel.port1)
    const client = createWorkerEngineClient({ port: channel.port2 })

    await expect(client.invoke('readValue')).resolves.toBe(4)

    client.dispose()
    host.dispose()
  })

  it('relays subscriptions back to the client', async () => {
    const channel = new MessageChannel()
    const eventListeners = new Set<(event: EngineEvent) => void>()
    const host = createWorkerEngineHost(
      {
        subscribe(listener: (event: EngineEvent) => void) {
          eventListeners.add(listener)
          return () => eventListeners.delete(listener)
        },
      },
      channel.port1,
    )

    const client = createWorkerEngineClient({ port: channel.port2 })
    const received: EngineEvent[] = []
    const unsubscribe = client.subscribe((event) => {
      received.push(event)
    })

    await waitFor(() => eventListeners.size === 1)

    eventListeners.forEach((listener) => {
      listener({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: Uint32Array.from([1, 2]),
        changedCells: [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: {
          batchId: 1,
          changedInputCount: 1,
          dirtyFormulaCount: 0,
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
          rangeNodeVisits: 0,
          recalcMs: 0,
          compileMs: 0,
        },
      })
    })

    await waitFor(() => received.length === 1)

    expect(received).toHaveLength(1)
    unsubscribe()
    client.dispose()
    host.dispose()
  })

  it('relays viewport patch subscriptions with subscription args', async () => {
    const channel = new MessageChannel()
    const host = createWorkerEngineHost(
      {
        subscribeViewportPatches(subscription, listener) {
          listener(
            encodeViewportPatch({
              version: 1,
              full: true,
              viewport: subscription,
              metrics: {
                batchId: 3,
                changedInputCount: 1,
                dirtyFormulaCount: 0,
                wasmFormulaCount: 0,
                jsFormulaCount: 0,
                rangeNodeVisits: 0,
                recalcMs: 0,
                compileMs: 0,
              },
              styles: [],
              cells: [],
              columns: [],
              rows: [],
            }),
          )
          return () => undefined
        },
      },
      channel.port1,
    )

    const client = createWorkerEngineClient({ port: channel.port2 })
    const received: Uint8Array[] = []
    const unsubscribe = client.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
      (patch) => {
        received.push(patch)
      },
    )

    await waitFor(() => received.length === 1)

    const firstPatch = received[0]
    expect(firstPatch).toBeDefined()
    if (!firstPatch) {
      throw new Error('Expected a viewport patch')
    }
    expect(decodeViewportPatch(firstPatch)).toMatchObject({
      full: true,
      viewport: {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    })

    unsubscribe()
    client.dispose()
    host.dispose()
  })

  it('relays render tile delta subscriptions with subscription args', async () => {
    const channel = new MessageChannel()
    const host = createWorkerEngineHost(
      {
        subscribeRenderTileDeltas(subscription, listener) {
          listener(
            encodeRenderTileDeltaBatch({
              magic: 'bilig.render.tile.delta',
              version: 1,
              sheetId: subscription.sheetId,
              batchId: 4,
              cameraSeq: 8,
              mutations: [
                {
                  kind: 'invalidate',
                  tileId: subscription.sheetId * 100,
                  reason: 'subscription-relay-check',
                },
              ],
            }),
          )
          return () => undefined
        },
      },
      channel.port1,
    )

    const client = createWorkerEngineClient({ port: channel.port2 })
    const received: Uint8Array[] = []
    const unsubscribe = client.subscribeRenderTileDeltas(
      {
        sheetId: 2,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      },
      (delta) => {
        received.push(delta)
      },
    )

    await waitFor(() => received.length === 1)

    const firstDelta = received[0]
    expect(firstDelta).toBeDefined()
    if (!firstDelta) {
      throw new Error('Expected a render tile delta')
    }
    expect(decodeRenderTileDeltaBatch(firstDelta)).toMatchObject({
      sheetId: 2,
      batchId: 4,
      cameraSeq: 8,
      mutations: [
        {
          kind: 'invalidate',
          tileId: 200,
          reason: 'subscription-relay-check',
        },
      ],
    })

    unsubscribe()
    client.dispose()
    host.dispose()
  })

  it('relays workbook delta subscriptions', async () => {
    const channel = new MessageChannel()
    const host = createWorkerEngineHost(
      {
        subscribeWorkbookDeltas(listener) {
          listener(
            encodeWorkbookDeltaBatchV3({
              magic: 'bilig.workbook.delta.v3',
              version: 1,
              seq: 9,
              source: 'workerAuthoritative',
              sheetId: 7,
              sheetOrdinal: 7,
              valueSeq: 11,
              styleSeq: 11,
              axisSeqX: 1,
              axisSeqY: 1,
              freezeSeq: 1,
              calcSeq: 11,
              dirty: {
                cellRanges: Uint32Array.from([1, 1, 1, 1, 7]),
                axisX: new Uint32Array(),
                axisY: new Uint32Array(),
              },
            }),
          )
          return () => undefined
        },
      },
      channel.port1,
    )

    const client = createWorkerEngineClient({ port: channel.port2 })
    const received: Uint8Array[] = []
    const unsubscribe = client.subscribeWorkbookDeltas((delta) => {
      received.push(delta)
    })

    await waitFor(() => received.length === 1)

    const firstDelta = received[0]
    expect(firstDelta).toBeDefined()
    if (!firstDelta) {
      throw new Error('Expected a workbook delta')
    }
    expect(decodeWorkbookDeltaBatchV3(firstDelta)).toMatchObject({
      seq: 9,
      sheetId: 7,
      sheetOrdinal: 7,
      dirty: {
        cellRanges: Uint32Array.from([1, 1, 1, 1, 7]),
      },
    })

    unsubscribe()
    client.dispose()
    host.dispose()
  })
})
