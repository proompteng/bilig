import type { HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import { createBrowserHelloReplay, type SnapshotReplayState } from "./browser-sync-replay.js";

type AppendBatchFrame = Extract<ProtocolFrame, { kind: "appendBatch" }>;

export interface OpenWorkbookBrowserSessionOptions {
  register?(frame: HelloFrame): void | Promise<void>;
  latestCursor: number | Promise<number>;
  latestSnapshot: SnapshotReplayState | Promise<SnapshotReplayState>;
  listMissedFrames(cursorFloor: number): AppendBatchFrame[] | Promise<AppendBatchFrame[]>;
}

export async function openWorkbookBrowserSession(
  frame: HelloFrame,
  options: OpenWorkbookBrowserSessionOptions,
): Promise<ProtocolFrame[]> {
  await options.register?.(frame);
  return createBrowserHelloReplay({
    documentId: frame.documentId,
    lastServerCursor: frame.lastServerCursor,
    latestCursor: options.latestCursor,
    latestSnapshot: options.latestSnapshot,
    listMissedFrames: (cursorFloor) => options.listMissedFrames(cursorFloor),
  });
}
