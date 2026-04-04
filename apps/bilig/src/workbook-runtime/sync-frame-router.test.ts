import { describe, expect, it, vi } from "vitest";
import { createUnsupportedSyncFrame, routeWorkbookSyncFrame } from "./sync-frame-router.js";

describe("sync-frame-router", () => {
  it("routes supported frame kinds to the matching handler", async () => {
    const hello = vi.fn(async () => "hello");
    const appendBatch = vi.fn(async () => "append");
    const snapshotChunk = vi.fn(async () => "snapshot");
    const heartbeat = vi.fn(async () => "heartbeat");
    const passthrough = vi.fn(async () => "passthrough");
    const unsupported = vi.fn(async () => "unsupported");

    expect(
      await routeWorkbookSyncFrame(
        {
          kind: "appendBatch",
          documentId: "doc-1",
          cursor: 0,
          batch: {
            id: "batch-1",
            replicaId: "replica-1",
            clock: { counter: 1 },
            ops: [],
          },
        },
        {
          hello,
          appendBatch,
          snapshotChunk,
          heartbeat,
          passthrough,
          unsupported,
        },
      ),
    ).toBe("append");

    expect(appendBatch).toHaveBeenCalledTimes(1);
    expect(hello).not.toHaveBeenCalled();
    expect(snapshotChunk).not.toHaveBeenCalled();
    expect(heartbeat).not.toHaveBeenCalled();
    expect(passthrough).not.toHaveBeenCalled();
    expect(unsupported).not.toHaveBeenCalled();
  });

  it("passes ack, error, and cursor watermark frames through the passthrough handler", async () => {
    const passthrough = vi.fn(async (frame) => frame.kind);

    const result = await routeWorkbookSyncFrame(
      {
        kind: "cursorWatermark",
        documentId: "doc-1",
        cursor: 2,
        compactedCursor: 1,
      },
      {
        hello: async () => "hello",
        appendBatch: async () => "append",
        snapshotChunk: async () => "snapshot",
        heartbeat: async () => "heartbeat",
        passthrough,
        unsupported: async () => "unsupported",
      },
    );

    expect(result).toBe("cursorWatermark");
    expect(passthrough).toHaveBeenCalledTimes(1);
  });

  it("builds a consistent unsupported-frame error payload", () => {
    expect(
      createUnsupportedSyncFrame("doc-1", "UNSUPPORTED_SYNC_FRAME", "mystery", "Unsupported frame"),
    ).toEqual({
      kind: "error",
      documentId: "doc-1",
      code: "UNSUPPORTED_SYNC_FRAME",
      message: "Unsupported frame mystery",
      retryable: false,
    });
  });
});
