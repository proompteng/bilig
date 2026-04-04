import type {
  AckFrame,
  CursorWatermarkFrame,
  HeartbeatFrame,
  ProtocolFrame,
} from "@bilig/binary-protocol";

export type AppendBatchFrame = Extract<ProtocolFrame, { kind: "appendBatch" }>;

export function createAppendBatchFrame(
  documentId: string,
  cursor: number,
  batch: AppendBatchFrame["batch"],
): AppendBatchFrame {
  return {
    kind: "appendBatch",
    documentId,
    cursor,
    batch,
  };
}

export function createAckFrame(
  documentId: string,
  batchId: string,
  cursor: number,
  acceptedAtUnixMs = Date.now(),
): AckFrame {
  return {
    kind: "ack",
    documentId,
    batchId,
    cursor,
    acceptedAtUnixMs,
  };
}

export function createHeartbeatFrame(
  documentId: string,
  cursor: number,
  sentAtUnixMs = Date.now(),
): HeartbeatFrame {
  return {
    kind: "heartbeat",
    documentId,
    cursor,
    sentAtUnixMs,
  };
}

export function createCursorWatermarkFrame(
  documentId: string,
  cursor: number,
  compactedCursor: number,
): CursorWatermarkFrame {
  return {
    kind: "cursorWatermark",
    documentId,
    cursor,
    compactedCursor,
  };
}

export function createHelloReplayFrames(
  snapshotFrames: ProtocolFrame[],
  missedFrames: AppendBatchFrame[],
  watermark: CursorWatermarkFrame,
): ProtocolFrame[] {
  return [...snapshotFrames, ...missedFrames, watermark];
}
