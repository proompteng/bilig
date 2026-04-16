import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { createAppendBatchFrame } from './sync-frame-shared.js'
import {
  getLocalCachedSnapshot,
  invalidateLocalSnapshotCache,
  maybeCompactLocalSession,
  publishLocalSnapshot,
  type LocalSnapshotSessionState,
} from './local-session-snapshot-store.js'

function createSession(documentId: string): LocalSnapshotSessionState {
  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `worksheet-host:${documentId}`,
  })
  engine.createSheet('Sheet1')
  return {
    documentId,
    engine,
    batches: [],
    latestSnapshot: null,
    snapshotCache: null,
    snapshotDirty: true,
    cursor: 0,
    replicaSnapshot: null,
    compactScheduled: false,
  }
}

describe('local-session-snapshot-store', () => {
  it('reuses cached snapshots until invalidated', () => {
    const session = createSession('doc-cache')

    const first = getLocalCachedSnapshot(session)
    const second = getLocalCachedSnapshot(session)

    expect(second).toBe(first)

    session.engine.setRangeValues(
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      [[123]],
    )
    invalidateLocalSnapshotCache(session)

    const third = getLocalCachedSnapshot(session)
    expect(third).not.toBe(first)
    expect(third.sheets[0]?.cells).toContainEqual(expect.objectContaining({ address: 'A1', value: 123 }))
  })

  it('publishes and compacts snapshots through the shared helper', () => {
    const session = createSession('doc-compact')
    const broadcastFrames: string[] = []
    let capturedSnapshot: WorkbookSnapshot | null = null
    let capturedBatch: ReturnType<typeof createAppendBatchFrame>['batch'] | null = null

    session.engine.subscribeBatches((batch) => {
      capturedBatch = batch
    })
    session.engine.setRangeValues(
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      [[456]],
    )
    if (!capturedBatch) {
      throw new Error('Expected engine batch')
    }
    session.cursor = 1
    const appendFrame = createAppendBatchFrame(session.documentId, 1, capturedBatch)
    session.batches = [{ cursor: appendFrame.cursor, frame: appendFrame }]

    maybeCompactLocalSession(session, {
      broadcast: (_documentId, protocolFrame) => {
        broadcastFrames.push(protocolFrame.kind)
      },
      getSession: () => session,
      snapshotAssemblies: new Map(),
      maxBatchBacklog: 0,
      schedule: (callback) => callback(),
    })

    expect(session.latestSnapshot).not.toBeNull()
    expect(session.batches).toEqual([])
    expect(broadcastFrames).toContain('snapshotChunk')
    expect(broadcastFrames.at(-1)).toBe('cursorWatermark')

    capturedSnapshot = getLocalCachedSnapshot(session)
    publishLocalSnapshot(session, capturedSnapshot, (_documentId, protocolFrame) => {
      broadcastFrames.push(protocolFrame.kind)
    })

    expect(session.latestSnapshot?.cursor).toBeGreaterThan(1)
    expect(broadcastFrames.at(-1)).toBe('cursorWatermark')
  })
})
