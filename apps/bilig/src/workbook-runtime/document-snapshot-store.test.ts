import { describe, expect, it } from 'vitest'
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames } from '@bilig/binary-protocol'
import { createInMemoryDocumentPersistence } from '@bilig/storage-server'
import { acceptPersistedSnapshotChunk } from './document-snapshot-store.js'

describe('document-snapshot-store', () => {
  it('does not persist assembled snapshot chunks that fail workbook snapshot validation', async () => {
    const persistence = createInMemoryDocumentPersistence()
    const snapshotAssemblies = new Map()
    const frames = createSnapshotChunkFrames({
      documentId: 'doc-1',
      snapshotId: 'snapshot-1',
      cursor: 7,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes: new TextEncoder().encode(
        JSON.stringify({
          version: 1,
          workbook: { name: 'doc-1' },
          sheets: [{ name: 'Sheet1', order: 0 }],
        }),
      ),
      chunkSize: 16,
    })

    await frames
      .slice(0, -1)
      .reduce<Promise<void>>(
        (previous, frame) => previous.then(() => acceptPersistedSnapshotChunk(persistence, snapshotAssemblies, frame)),
        Promise.resolve(),
      )

    await expect(acceptPersistedSnapshotChunk(persistence, snapshotAssemblies, frames.at(-1)!)).rejects.toThrow(
      'Workbook snapshot payload does not match the expected schema',
    )
    expect(await persistence.snapshots.latest('doc-1')).toBeNull()
  })
})
