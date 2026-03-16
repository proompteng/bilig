import { describe, expect, it } from "vitest";

import { decodeAgentFrame, decodeStdioMessages, encodeAgentFrame, encodeStdioMessage } from "../index.js";

describe("agent api", () => {
  it("roundtrips request and response envelopes", () => {
    const frame = {
      kind: "request" as const,
      request: {
        kind: "openWorkbookSession" as const,
        id: "req-1",
        documentId: "book-1",
        replicaId: "agent-local"
      }
    };

    expect(decodeAgentFrame(encodeAgentFrame(frame))).toEqual(frame);
  });

  it("decodes multiple stdio messages from one buffer", () => {
    const encoded = new Uint8Array([
      ...encodeStdioMessage({
        kind: "request",
        request: {
          kind: "getMetrics",
          id: "req-1",
          sessionId: "sess-1"
        }
      }),
      ...encodeStdioMessage({
        kind: "response",
        response: {
          kind: "ok",
          id: "req-1",
          value: { ok: true }
        }
      })
    ]);

    const decoded = decodeStdioMessages(encoded);
    expect(decoded.frames).toHaveLength(2);
    expect(decoded.remainder.byteLength).toBe(0);
  });
});
