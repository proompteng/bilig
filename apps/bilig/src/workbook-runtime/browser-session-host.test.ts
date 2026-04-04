import { describe, expect, it, vi } from "vitest";
import type { HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import { WorkbookBrowserSessionHost } from "./browser-session-host.js";

const helloFrame: HelloFrame = {
  kind: "hello",
  documentId: "doc-1",
  replicaId: "browser-1",
  sessionId: "session-1",
  protocolVersion: 1,
  lastServerCursor: 0,
  capabilities: [],
};

describe("WorkbookBrowserSessionHost", () => {
  it("opens browser sessions and tracks subscriber broadcasts", async () => {
    const register = vi.fn(async () => {});
    const host = new WorkbookBrowserSessionHost({
      register,
      latestCursor: async () => 4,
      latestSnapshot: async () => null,
      listMissedFrames: async (_documentId, _cursorFloor) => [
        {
          kind: "appendBatch" as const,
          documentId: "doc-1",
          cursor: 4,
          batch: {
            id: "batch-1",
            replicaId: "replica-1",
            clock: { counter: 1 },
            ops: [],
          },
        },
      ],
    });

    const received: ProtocolFrame[] = [];
    const detach = host.attachBrowser("doc-1", "sub-1", (frame) => {
      received.push(frame);
    });

    expect(host.listSubscriberIds("doc-1")).toEqual(["sub-1"]);

    const frames = await host.openBrowserSession(helloFrame);
    expect(register).toHaveBeenCalledWith(helloFrame);
    expect(frames).toContainEqual(
      expect.objectContaining({
        kind: "appendBatch",
        documentId: "doc-1",
      }),
    );

    const heartbeat: ProtocolFrame = {
      kind: "heartbeat",
      documentId: "doc-1",
      cursor: 7,
      sentAtUnixMs: 123,
    };
    host.broadcast("doc-1", heartbeat);
    expect(received).toEqual([heartbeat]);

    detach();
    expect(host.listSubscriberIds("doc-1")).toEqual([]);
  });
});
