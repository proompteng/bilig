import { describe, expect, it, vi } from "vitest";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import { createHttpSyncRelay } from "../sync-relay.js";

describe("sync relay", () => {
  it("uses worksheet-host replica ids and forwards batches over the frame endpoint", async () => {
    const seenFrames: unknown[] = [];
    const fetchImpl = vi.fn(async (url: string | URL, init?: RequestInit) => {
      expect(String(url)).toBe("https://bilig.proompteng.ai/v2/documents/doc-1/frames");
      if (!(init?.body instanceof Uint8Array)) {
        throw new TypeError("expected binary request body");
      }
      seenFrames.push(decodeFrame(init.body));
      const responseFrame =
        seenFrames.length === 1
          ? {
              kind: "cursorWatermark" as const,
              documentId: "doc-1",
              cursor: 7,
              compactedCursor: 7,
            }
          : {
              kind: "ack" as const,
              documentId: "doc-1",
              batchId: "batch-1",
              cursor: 8,
              acceptedAtUnixMs: 1,
            };
      return new Response(Buffer.from(encodeFrame(responseFrame)), {
        status: 200,
      });
    });

    const relay = createHttpSyncRelay({
      documentId: "doc-1",
      baseUrl: "https://bilig.proompteng.ai",
      fetchImpl,
    });

    const batch: EngineOpBatch = {
      id: "batch-1",
      replicaId: "worksheet-host:doc-1",
      clock: { counter: 1 },
      ops: [],
    };

    await relay.send(batch);

    expect(seenFrames[0]).toEqual(
      expect.objectContaining({
        kind: "hello",
        documentId: "doc-1",
        replicaId: "worksheet-host:doc-1",
        sessionId: "doc-1:worksheet-host:doc-1",
        capabilities: ["local-relay"],
      }),
    );
    expect(seenFrames[1]).toEqual(
      expect.objectContaining({
        kind: "appendBatch",
        documentId: "doc-1",
        batch: expect.objectContaining({ id: "batch-1" }),
      }),
    );
  });
});
