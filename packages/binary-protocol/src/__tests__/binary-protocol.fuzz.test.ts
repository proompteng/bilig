import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { runProperty } from '@bilig/test-fuzz'
import { createSnapshotChunkFrames, decodeFrame, encodeFrame, type ProtocolFrame, WORKBOOK_SNAPSHOT_CONTENT_TYPE } from '../index.js'

describe('binary protocol fuzz', () => {
  it('should roundtrip generated protocol frames through encode/decode', async () => {
    await runProperty({
      suite: 'binary-protocol/frame-roundtrip',
      arbitrary: protocolFrameArbitrary,
      predicate: async (frame) => {
        expect(normalizeFrame(decodeFrame(encodeFrame(frame)))).toEqual(normalizeFrame(frame))
      },
    })
  })

  it('should preserve snapshot bytes across generated chunk boundaries', async () => {
    await runProperty({
      suite: 'binary-protocol/snapshot-chunk-roundtrip',
      arbitrary: fc.record({
        documentId: fc.constantFrom('doc-a', 'doc-b'),
        snapshotId: fc.uuid(),
        cursor: fc.integer({ min: 0, max: 1_000 }),
        bytes: fc.uint8Array({ maxLength: 64 }),
        chunkSize: fc.integer({ min: 1, max: 16 }),
      }),
      predicate: async ({ documentId, snapshotId, cursor, bytes, chunkSize }) => {
        const frames = createSnapshotChunkFrames({
          documentId,
          snapshotId,
          cursor,
          contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          bytes,
          chunkSize,
        })
        expect(frames.map((frame) => frame.chunkIndex)).toEqual(frames.map((_, index) => index))
        expect(new Set(frames.map((frame) => frame.chunkCount)).size).toBeLessThanOrEqual(1)
        expect(
          new Uint8Array(
            frames.reduce<number[]>((all, frame) => {
              all.push(...frame.bytes)
              return all
            }, []),
          ),
        ).toEqual(bytes)
      },
    })
  })
})

// Helpers

const engineOpArbitrary = fc.oneof(
  fc.constantFrom('Book', 'Spec', 'Revenue').map((name) => ({ kind: 'upsertWorkbook' as const, name })),
  fc
    .record({
      name: fc.constantFrom('Sheet1', 'Sheet2'),
      order: fc.integer({ min: 0, max: 4 }),
    })
    .map(({ name, order }) => ({ kind: 'upsertSheet' as const, name, order })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      address: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
      value: fc.oneof(fc.integer({ min: -50, max: 50 }), fc.boolean(), fc.constantFrom('north', 'south'), fc.constant(null)),
    })
    .map(({ sheetName, address, value }) => ({ kind: 'setCellValue' as const, sheetName, address, value })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      address: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
      formula: fc.constantFrom('A1+1', 'B2*2', '1+2'),
    })
    .map(({ sheetName, address, formula }) => ({ kind: 'setCellFormula' as const, sheetName, address, formula })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      address: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
      format: fc.constantFrom('0.00', '0%', '@'),
    })
    .map(({ sheetName, address, format }) => ({ kind: 'setCellFormat' as const, sheetName, address, format })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      address: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
    })
    .map(({ sheetName, address }) => ({ kind: 'clearCell' as const, sheetName, address })),
)

const engineBatchArbitrary = fc
  .record({
    id: fc.uuid(),
    replicaId: fc.constantFrom('replica-a', 'replica-b'),
    counter: fc.integer({ min: 1, max: 1_000 }),
    ops: fc.array(engineOpArbitrary, { minLength: 0, maxLength: 8 }),
  })
  .map(
    ({ id, replicaId, counter, ops }) =>
      ({
        id,
        replicaId,
        clock: { counter },
        ops,
      }) satisfies EngineOpBatch,
  )

const protocolFrameArbitrary = fc.oneof<ProtocolFrame>(
  fc
    .record({
      documentId: fc.constantFrom('doc-a', 'doc-b'),
      replicaId: fc.uuid(),
      sessionId: fc.uuid(),
      protocolVersion: fc.constant(1),
      lastServerCursor: fc.integer({ min: 0, max: 1_000 }),
      capabilities: fc.array(fc.constantFrom('sync', 'agent', 'presence'), { minLength: 0, maxLength: 3 }),
    })
    .map((frame) => ({
      kind: 'hello' as const,
      documentId: frame.documentId,
      replicaId: frame.replicaId,
      sessionId: frame.sessionId,
      protocolVersion: frame.protocolVersion,
      lastServerCursor: frame.lastServerCursor,
      capabilities: frame.capabilities,
    })),
  fc
    .record({
      documentId: fc.constantFrom('doc-a', 'doc-b'),
      cursor: fc.integer({ min: 0, max: 1_000 }),
      batch: engineBatchArbitrary,
    })
    .map((frame) => ({
      kind: 'appendBatch' as const,
      documentId: frame.documentId,
      cursor: frame.cursor,
      batch: frame.batch,
    })),
  fc
    .record({
      documentId: fc.constantFrom('doc-a', 'doc-b'),
      batchId: fc.uuid(),
      cursor: fc.integer({ min: 0, max: 1_000 }),
      acceptedAtUnixMs: fc.integer({ min: 1_000, max: 99_999 }),
    })
    .map((frame) => ({
      kind: 'ack' as const,
      documentId: frame.documentId,
      batchId: frame.batchId,
      cursor: frame.cursor,
      acceptedAtUnixMs: frame.acceptedAtUnixMs,
    })),
  fc
    .record({
      documentId: fc.constantFrom('doc-a', 'doc-b'),
      snapshotId: fc.uuid(),
      cursor: fc.integer({ min: 0, max: 1_000 }),
      chunkIndex: fc.integer({ min: 0, max: 3 }),
      chunkCount: fc.integer({ min: 1, max: 4 }),
      contentType: fc.constant(WORKBOOK_SNAPSHOT_CONTENT_TYPE),
      bytes: fc.uint8Array({ maxLength: 32 }),
    })
    .filter((frame) => frame.chunkIndex < frame.chunkCount)
    .map((frame) => ({
      kind: 'snapshotChunk' as const,
      documentId: frame.documentId,
      snapshotId: frame.snapshotId,
      cursor: frame.cursor,
      chunkIndex: frame.chunkIndex,
      chunkCount: frame.chunkCount,
      contentType: frame.contentType,
      bytes: frame.bytes,
    })),
  fc
    .record({
      documentId: fc.constantFrom('doc-a', 'doc-b'),
      cursor: fc.integer({ min: 0, max: 1_000 }),
      compactedCursor: fc.integer({ min: 0, max: 1_000 }),
    })
    .filter((frame) => frame.compactedCursor <= frame.cursor)
    .map((frame) => ({
      kind: 'cursorWatermark' as const,
      documentId: frame.documentId,
      cursor: frame.cursor,
      compactedCursor: frame.compactedCursor,
    })),
  fc
    .record({
      documentId: fc.constantFrom('doc-a', 'doc-b'),
      cursor: fc.integer({ min: 0, max: 1_000 }),
      sentAtUnixMs: fc.integer({ min: 1_000, max: 99_999 }),
    })
    .map((frame) => ({
      kind: 'heartbeat' as const,
      documentId: frame.documentId,
      cursor: frame.cursor,
      sentAtUnixMs: frame.sentAtUnixMs,
    })),
  fc
    .record({
      documentId: fc.constantFrom('doc-a', 'doc-b'),
      code: fc.constantFrom('BROKEN', 'RETRY', 'STALE'),
      message: fc.constantFrom('broken', 'retry later', 'stale client'),
      retryable: fc.boolean(),
    })
    .map((frame) => ({
      kind: 'error' as const,
      documentId: frame.documentId,
      code: frame.code,
      message: frame.message,
      retryable: frame.retryable,
    })),
)

function normalizeFrame(frame: ProtocolFrame): unknown {
  if (frame.kind === 'snapshotChunk') {
    return { ...frame, bytes: [...frame.bytes] }
  }
  return frame
}
