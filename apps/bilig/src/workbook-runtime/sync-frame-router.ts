import type { ErrorFrame, ProtocolFrame } from '@bilig/binary-protocol'

type HelloFrame = Extract<ProtocolFrame, { kind: 'hello' }>
type AppendBatchFrame = Extract<ProtocolFrame, { kind: 'appendBatch' }>
type SnapshotChunkFrame = Extract<ProtocolFrame, { kind: 'snapshotChunk' }>
type HeartbeatFrame = Extract<ProtocolFrame, { kind: 'heartbeat' }>
type PassthroughFrame = Extract<ProtocolFrame, { kind: 'cursorWatermark' | 'ack' | 'error' }>

export interface WorkbookSyncFrameRouter<Output> {
  hello(frame: HelloFrame): Output | Promise<Output>
  appendBatch(frame: AppendBatchFrame): Output | Promise<Output>
  snapshotChunk(frame: SnapshotChunkFrame): Output | Promise<Output>
  heartbeat(frame: HeartbeatFrame): Output | Promise<Output>
  passthrough(frame: PassthroughFrame): Output | Promise<Output>
  unsupported(frame: ProtocolFrame): Output | Promise<Output>
}

export function createUnsupportedSyncFrame(documentId: string, code: string, frameKind: string, message: string): ErrorFrame {
  return {
    kind: 'error',
    documentId,
    code,
    message: `${message} ${frameKind}`,
    retryable: false,
  }
}

export async function routeWorkbookSyncFrame<Output>(frame: ProtocolFrame, router: WorkbookSyncFrameRouter<Output>): Promise<Output> {
  switch (frame.kind) {
    case 'hello':
      return router.hello(frame)
    case 'appendBatch':
      return router.appendBatch(frame)
    case 'snapshotChunk':
      return router.snapshotChunk(frame)
    case 'heartbeat':
      return router.heartbeat(frame)
    case 'cursorWatermark':
    case 'ack':
    case 'error':
      return router.passthrough(frame)
    default:
      return router.unsupported(frame)
  }
}
