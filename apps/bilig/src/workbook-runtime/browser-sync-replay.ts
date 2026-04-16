import { createSnapshotChunkFrames, type ProtocolFrame, type SnapshotChunkFrame } from '@bilig/binary-protocol'
import { createCursorWatermarkFrame, createHelloReplayFrames } from './sync-frame-shared.js'

type AppendBatchFrame = Extract<ProtocolFrame, { kind: 'appendBatch' }>

export interface SnapshotFrameReplayState {
  cursor: number
  frames: SnapshotChunkFrame[]
}

export interface SnapshotBytesReplayState {
  documentId: string
  snapshotId: string
  cursor: number
  contentType: string
  bytes: Uint8Array
}

export type SnapshotReplayState = SnapshotFrameReplayState | SnapshotBytesReplayState | null

function createSnapshotReplayFrames(
  documentId: string,
  lastServerCursor: number,
  latestSnapshot: SnapshotReplayState,
): SnapshotChunkFrame[] {
  if (!latestSnapshot || lastServerCursor >= latestSnapshot.cursor) {
    return []
  }
  if ('frames' in latestSnapshot) {
    return latestSnapshot.frames
  }
  return createSnapshotChunkFrames({
    documentId,
    snapshotId: latestSnapshot.snapshotId,
    cursor: latestSnapshot.cursor,
    contentType: latestSnapshot.contentType,
    bytes: latestSnapshot.bytes,
  })
}

export async function createBrowserHelloReplay(options: {
  documentId: string
  lastServerCursor: number
  latestCursor: number | Promise<number>
  latestSnapshot: SnapshotReplayState | Promise<SnapshotReplayState>
  listMissedFrames(cursorFloor: number): AppendBatchFrame[] | Promise<AppendBatchFrame[]>
}): Promise<ProtocolFrame[]> {
  const latestSnapshot = await options.latestSnapshot
  const latestCursor = await options.latestCursor
  const cursorFloor = Math.max(options.lastServerCursor, latestSnapshot?.cursor ?? 0)
  const missedFrames = await options.listMissedFrames(cursorFloor)
  return createHelloReplayFrames(
    createSnapshotReplayFrames(options.documentId, options.lastServerCursor, latestSnapshot),
    missedFrames,
    createCursorWatermarkFrame(options.documentId, latestCursor, latestSnapshot?.cursor ?? 0),
  )
}
