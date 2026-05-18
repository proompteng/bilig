import { describe, expect, it, vi } from 'vitest'
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames, type SnapshotChunkFrame } from '@bilig/binary-protocol'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  SNAPSHOT_ASSEMBLY_MAX_AGE_MS,
  SNAPSHOT_ASSEMBLY_MAX_CHUNKS,
  acceptSnapshotChunk,
  createSnapshotPublication,
  decodeWorkbookSnapshotBytes,
  encodeWorkbookSnapshot,
} from './session-shared.js'

function createSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'doc-1' },
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        cells: [{ address: 'A1', value: 1 }],
      },
    ],
  }
}

describe('session-shared', () => {
  it('reuses encoded bytes for the same snapshot object', () => {
    const snapshot = createSnapshot()

    const first = encodeWorkbookSnapshot(snapshot)
    const second = encodeWorkbookSnapshot(snapshot)
    const publication = createSnapshotPublication('doc-1', 1, snapshot)

    expect(second).toBe(first)
    expect(publication.bytes).toBe(first)
  })

  it('creates distinct snapshot publication ids within the same clock tick', () => {
    const nowSpy = vi.spyOn(Date, 'now').mockReturnValue(123)
    try {
      const snapshot = createSnapshot()
      const first = createSnapshotPublication('doc-1', 1, snapshot)
      const second = createSnapshotPublication('doc-1', 2, snapshot)

      expect(second.snapshotId).not.toBe(first.snapshotId)
      expect(first.frames.every((frame) => frame.snapshotId === first.snapshotId)).toBe(true)
      expect(second.frames.every((frame) => frame.snapshotId === second.snapshotId)).toBe(true)
    } finally {
      nowSpy.mockRestore()
    }
  })

  it('evicts stale snapshot assemblies before accepting a new chunk', () => {
    const registry = new Map()
    registry.set('stale', {
      documentId: 'doc-stale',
      snapshotId: 'stale',
      cursor: 1,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      chunkCount: 2,
      chunks: [new Uint8Array([1]), undefined],
      totalByteLength: 1,
      updatedAtUnixMs: 0,
    })

    const bytes = encodeWorkbookSnapshot(createSnapshot())
    const frames = createSnapshotChunkFrames({
      documentId: 'doc-1',
      snapshotId: 'fresh',
      cursor: 2,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes,
    })

    const result = acceptSnapshotChunk(registry, frames[0]!, {
      nowUnixMs: SNAPSHOT_ASSEMBLY_MAX_AGE_MS + 1,
    })

    expect(result?.snapshotId).toBe('fresh')
    expect(registry.has('stale')).toBe(false)
    expect(registry.has('fresh')).toBe(false)
  })

  it('rejects snapshot chunks that conflict with an existing assembly identity', () => {
    const registry = new Map()
    const frames = createSnapshotChunkFrames({
      documentId: 'doc-1',
      snapshotId: 'snapshot-1',
      cursor: 7,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes: new TextEncoder().encode('abcd'),
      chunkSize: 2,
    })

    expect(acceptSnapshotChunk(registry, frames[0]!)).toBeNull()
    expect(
      acceptSnapshotChunk(registry, {
        ...frames[1]!,
        documentId: 'doc-2',
      }),
    ).toBeNull()
    expect(registry.has('snapshot-1')).toBe(false)
  })

  it('invalidates snapshot assemblies when duplicate chunks carry different bytes', () => {
    const registry = new Map()
    const frames = createSnapshotChunkFrames({
      documentId: 'doc-1',
      snapshotId: 'snapshot-1',
      cursor: 7,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes: new TextEncoder().encode('abcd'),
      chunkSize: 2,
    })

    expect(acceptSnapshotChunk(registry, frames[0]!)).toBeNull()
    expect(
      acceptSnapshotChunk(registry, {
        ...frames[0]!,
        bytes: new TextEncoder().encode('zz'),
      }),
    ).toBeNull()
    expect(registry.has('snapshot-1')).toBe(false)
    expect(acceptSnapshotChunk(registry, frames[1]!)).toBeNull()
  })

  it('rejects malformed snapshot chunk geometry before tracking an assembly', () => {
    const registry = new Map()
    const frame: SnapshotChunkFrame = {
      kind: 'snapshotChunk',
      documentId: 'doc-1',
      snapshotId: 'snapshot-1',
      cursor: 7,
      chunkIndex: 0,
      chunkCount: 0,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes: new TextEncoder().encode('ab'),
    }

    expect(acceptSnapshotChunk(registry, frame)).toBeNull()
    expect(registry.has('snapshot-1')).toBe(false)
  })

  it('rejects snapshot assemblies with unbounded chunk counts before allocating the assembly', () => {
    const registry = new Map()
    const frame: SnapshotChunkFrame = {
      kind: 'snapshotChunk',
      documentId: 'doc-1',
      snapshotId: 'snapshot-many-chunks',
      cursor: 7,
      chunkIndex: 0,
      chunkCount: SNAPSHOT_ASSEMBLY_MAX_CHUNKS + 1,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes: new Uint8Array([1]),
    }

    expect(acceptSnapshotChunk(registry, frame)).toBeNull()
    expect(registry.has('snapshot-many-chunks')).toBe(false)
  })

  it('rejects snapshot assemblies once accumulated bytes exceed the configured cap', () => {
    const registry = new Map()
    const frames = createSnapshotChunkFrames({
      documentId: 'doc-1',
      snapshotId: 'snapshot-too-large',
      cursor: 7,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes: new TextEncoder().encode('abcd'),
      chunkSize: 2,
    })

    expect(acceptSnapshotChunk(registry, frames[0]!, { maxBytes: 3 })).toBeNull()
    expect(registry.has('snapshot-too-large')).toBe(true)
    expect(acceptSnapshotChunk(registry, frames[1]!, { maxBytes: 3 })).toBeNull()
    expect(registry.has('snapshot-too-large')).toBe(false)
  })

  it('decodes assembled workbook snapshots through the shared protocol guard', () => {
    const snapshot = createSnapshot()

    expect(
      decodeWorkbookSnapshotBytes({
        documentId: 'doc-1',
        snapshotId: 'snapshot-1',
        cursor: 1,
        contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
        bytes: encodeWorkbookSnapshot(snapshot),
      }),
    ).toEqual(snapshot)
  })

  it('rejects assembled workbook payloads that are missing sheet cells', () => {
    expect(() =>
      decodeWorkbookSnapshotBytes({
        documentId: 'doc-1',
        snapshotId: 'snapshot-1',
        cursor: 1,
        contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
        bytes: new TextEncoder().encode(
          JSON.stringify({
            version: 1,
            workbook: { name: 'doc-1' },
            sheets: [{ name: 'Sheet1', order: 0 }],
          }),
        ),
      }),
    ).toThrow('Workbook snapshot payload does not match the expected schema')
  })
})
