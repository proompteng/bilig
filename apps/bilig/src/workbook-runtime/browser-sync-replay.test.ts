import { describe, expect, it } from 'vitest'
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE } from '@bilig/binary-protocol'
import { createAppendBatchFrame } from './sync-frame-shared.js'
import { createBrowserHelloReplay } from './browser-sync-replay.js'

describe('browser-sync-replay', () => {
  it('reuses prebuilt snapshot frames for local replay state', async () => {
    const replay = await createBrowserHelloReplay({
      documentId: 'doc-local',
      lastServerCursor: 0,
      latestCursor: 7,
      latestSnapshot: {
        cursor: 5,
        frames: [
          {
            kind: 'snapshotChunk',
            documentId: 'doc-local',
            snapshotId: 'snap-1',
            cursor: 5,
            contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
            chunkIndex: 0,
            chunkCount: 1,
            bytes: new Uint8Array([1, 2, 3]),
          },
        ],
      },
      listMissedFrames: () => [
        createAppendBatchFrame('doc-local', 6, {
          id: 'batch-1',
          replicaId: 'replica-1',
          clock: { counter: 1 },
          ops: [],
        }),
      ],
    })

    expect(replay[0]?.kind).toBe('snapshotChunk')
    expect(replay[1]).toEqual(
      expect.objectContaining({
        kind: 'appendBatch',
        documentId: 'doc-local',
        cursor: 6,
      }),
    )
    expect(replay.at(-1)).toEqual({
      kind: 'cursorWatermark',
      documentId: 'doc-local',
      cursor: 7,
      compactedCursor: 5,
    })
  })

  it('builds snapshot chunk frames from persisted snapshot bytes', async () => {
    const replay = await createBrowserHelloReplay({
      documentId: 'doc-persisted',
      lastServerCursor: 0,
      latestCursor: 4,
      latestSnapshot: {
        documentId: 'doc-persisted',
        snapshotId: 'snap-2',
        cursor: 4,
        contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
        bytes: new TextEncoder().encode(
          JSON.stringify({
            version: 1,
            workbook: { name: 'Test' },
            sheets: [],
          }),
        ),
      },
      listMissedFrames: () => [],
    })

    expect(replay).toHaveLength(2)
    expect(replay[0]).toEqual(
      expect.objectContaining({
        kind: 'snapshotChunk',
        documentId: 'doc-persisted',
        snapshotId: 'snap-2',
      }),
    )
    expect(replay[1]).toEqual({
      kind: 'cursorWatermark',
      documentId: 'doc-persisted',
      cursor: 4,
      compactedCursor: 4,
    })
  })
})
