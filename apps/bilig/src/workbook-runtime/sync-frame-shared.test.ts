import { describe, expect, it } from 'vitest'
import {
  createAckFrame,
  createAppendBatchFrame,
  createCursorWatermarkFrame,
  createHeartbeatFrame,
  createHelloReplayFrames,
} from './sync-frame-shared.js'

describe('sync-frame-shared', () => {
  it('creates standard sync frames', () => {
    expect(
      createAppendBatchFrame('doc-1', 2, {
        id: 'batch-1',
        replicaId: 'replica-1',
        clock: { counter: 1 },
        ops: [],
      }),
    ).toEqual({
      kind: 'appendBatch',
      documentId: 'doc-1',
      cursor: 2,
      batch: {
        id: 'batch-1',
        replicaId: 'replica-1',
        clock: { counter: 1 },
        ops: [],
      },
    })
    expect(createAckFrame('doc-1', 'batch-1', 2, 99)).toEqual({
      kind: 'ack',
      documentId: 'doc-1',
      batchId: 'batch-1',
      cursor: 2,
      acceptedAtUnixMs: 99,
    })
    expect(createHeartbeatFrame('doc-1', 3, 77)).toEqual({
      kind: 'heartbeat',
      documentId: 'doc-1',
      cursor: 3,
      sentAtUnixMs: 77,
    })
    expect(createCursorWatermarkFrame('doc-1', 4, 1)).toEqual({
      kind: 'cursorWatermark',
      documentId: 'doc-1',
      cursor: 4,
      compactedCursor: 1,
    })
  })

  it('builds hello replay frames in the expected order', () => {
    const snapshotFrame = {
      kind: 'snapshotChunk' as const,
      documentId: 'doc-1',
      snapshotId: 'snap-1',
      cursor: 5,
      contentType: 'application/vnd.bilig.workbook+json',
      chunkIndex: 0,
      chunkCount: 1,
      bytes: new Uint8Array([1]),
    }
    const appendFrame = createAppendBatchFrame('doc-1', 6, {
      id: 'batch-1',
      replicaId: 'replica-1',
      clock: { counter: 1 },
      ops: [],
    })
    const watermark = createCursorWatermarkFrame('doc-1', 6, 5)

    expect(createHelloReplayFrames([snapshotFrame], [appendFrame], watermark)).toEqual([snapshotFrame, appendFrame, watermark])
  })
})
