import { describe, expect, it, vi } from "vitest";
import type { HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import { WorkbookBrowserSessionHost } from "./browser-session-host.js";
import { WorkbookSyncSessionHost } from "./workbook-sync-session-host.js";

const helloFrame: HelloFrame = {
  kind: "hello",
  documentId: "doc-1",
  replicaId: "browser-1",
  sessionId: "session-1",
  protocolVersion: 1,
  lastServerCursor: 0,
  capabilities: [],
};

describe("WorkbookSyncSessionHost", () => {
  it("routes sync frames through the shared delegates", async () => {
    const browserSessionHost = new WorkbookBrowserSessionHost({
      latestCursor: async () => 0,
      latestSnapshot: async () => null,
      listMissedFrames: async () => [],
    });
    const hello = vi.fn(
      async () =>
        [
          { kind: "ack", documentId: "doc-1", batchId: "h", cursor: 0, acceptedAtUnixMs: 1 },
        ] as ProtocolFrame[],
    );
    const appendBatch = vi.fn(
      async () =>
        [
          { kind: "ack", documentId: "doc-1", batchId: "a", cursor: 1, acceptedAtUnixMs: 2 },
        ] as ProtocolFrame[],
    );
    const snapshotChunk = vi.fn(
      async () =>
        [
          { kind: "ack", documentId: "doc-1", batchId: "s", cursor: 2, acceptedAtUnixMs: 3 },
        ] as ProtocolFrame[],
    );
    const heartbeat = vi.fn(
      async () =>
        [{ kind: "heartbeat", documentId: "doc-1", cursor: 4, sentAtUnixMs: 4 }] as ProtocolFrame[],
    );

    const host = new WorkbookSyncSessionHost<ProtocolFrame[]>({
      browserSessionHost,
      hello,
      appendBatch,
      snapshotChunk,
      heartbeat,
      passthrough: async (frame) => [frame],
      unsupported: async (frame) => [
        {
          kind: "error",
          documentId: frame.documentId,
          code: "UNSUPPORTED",
          message: `Unsupported sync frame ${frame.kind}`,
          retryable: false,
        },
      ],
    });

    await host.handleSyncFrame(helloFrame);
    await host.handleSyncFrame({
      kind: "appendBatch",
      documentId: "doc-1",
      cursor: 1,
      batch: { id: "batch-1", replicaId: "replica-1", clock: { counter: 1 }, ops: [] },
    });
    await host.handleSyncFrame({
      kind: "snapshotChunk",
      documentId: "doc-1",
      snapshotId: "snap-1",
      cursor: 2,
      contentType: "application/test",
      chunkIndex: 0,
      chunkCount: 1,
      bytes: new Uint8Array([1]),
    });
    await host.handleSyncFrame({
      kind: "heartbeat",
      documentId: "doc-1",
      cursor: 3,
      sentAtUnixMs: 10,
    });

    expect(hello).toHaveBeenCalledWith(helloFrame);
    expect(appendBatch).toHaveBeenCalledOnce();
    expect(snapshotChunk).toHaveBeenCalledOnce();
    expect(heartbeat).toHaveBeenCalledOnce();
  });
});
