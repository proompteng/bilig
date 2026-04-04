import { describe, expect, it, vi } from "vitest";

import type { AgentFrame } from "@bilig/agent-api";
import type { ProtocolFrame } from "@bilig/binary-protocol";

import { resolveAgentDocumentId, updateActorFromFrames } from "./document-supervisor-shared.js";

describe("document-supervisor-shared", () => {
  it("updates actor state from cursor and error frames", () => {
    const send = vi.fn();

    updateActorFromFrames({ send }, [
      {
        kind: "appendBatch",
        documentId: "doc-1",
        cursor: 7,
        batch: {
          id: "batch-1",
          replicaId: "replica-1",
          clock: { counter: 1 },
          ops: [],
        },
      },
      {
        kind: "cursorWatermark",
        documentId: "doc-1",
        cursor: 9,
        compactedCursor: 4,
      },
      {
        kind: "error",
        documentId: "doc-1",
        code: "BROKEN",
        message: "broken",
        retryable: false,
      },
    ] satisfies ProtocolFrame[]);

    expect(send).toHaveBeenCalledWith({ type: "cursor.updated", cursor: 7 });
    expect(send).toHaveBeenCalledWith({ type: "cursor.updated", cursor: 9 });
    expect(send).toHaveBeenCalledWith({ type: "snapshot.updated", cursor: 4 });
    expect(send).toHaveBeenCalledWith({ type: "error.raised", message: "broken" });
  });

  it("derives document ids from direct and session-scoped agent requests", () => {
    expect(
      resolveAgentDocumentId({
        kind: "request",
        request: {
          kind: "openWorkbookSession",
          id: "open-1",
          documentId: "doc-open",
          replicaId: "replica-1",
        },
      } satisfies AgentFrame),
    ).toBe("doc-open");

    expect(
      resolveAgentDocumentId({
        kind: "request",
        request: {
          kind: "readRange",
          id: "read-1",
          sessionId: "doc-read:replica-1",
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "A1",
          },
        },
      } satisfies AgentFrame),
    ).toBe("doc-read");
  });
});
