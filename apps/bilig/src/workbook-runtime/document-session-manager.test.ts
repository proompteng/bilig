import { describe, expect, it, vi } from "vitest";
import { createInMemoryDocumentPersistence } from "@bilig/storage-server";
import type { AgentFrame } from "@bilig/agent-api";
import { DocumentSessionManager } from "./document-session-manager.js";

describe("DocumentSessionManager", () => {
  it("delegates open and close session requests through the worksheet executor", async () => {
    const execute = vi.fn(async (frame: AgentFrame) => {
      if (frame.kind !== "request") {
        throw new Error("expected request frame");
      }
      if (frame.request.kind === "openWorkbookSession") {
        return {
          kind: "response",
          response: {
            kind: "ok",
            id: frame.request.id,
            sessionId: `${frame.request.documentId}:${frame.request.replicaId}`,
          },
        } satisfies AgentFrame;
      }
      return {
        kind: "response",
        response: {
          kind: "ok",
          id: frame.request.id,
        },
      } satisfies AgentFrame;
    });

    const manager = new DocumentSessionManager(createInMemoryDocumentPersistence(), "bilig-app", {
      execute,
    });

    const openResponse = await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "openWorkbookSession",
        id: "open-1",
        documentId: "doc-1",
        replicaId: "replica-1",
      },
    } satisfies AgentFrame);

    expect(openResponse).toEqual({
      kind: "response",
      response: {
        kind: "ok",
        id: "open-1",
        sessionId: "doc-1:replica-1",
      },
    });
    expect(execute).toHaveBeenCalledTimes(1);
    expect((await manager.getDocumentState("doc-1")).sessions).toContain("doc-1:replica-1");

    const closeResponse = await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "closeWorkbookSession",
        id: "close-1",
        sessionId: "doc-1:replica-1",
      },
    } satisfies AgentFrame);

    expect(closeResponse).toEqual({
      kind: "response",
      response: {
        kind: "ok",
        id: "close-1",
      },
    });
    expect(execute).toHaveBeenCalledTimes(2);
    expect((await manager.getDocumentState("doc-1")).sessions).toEqual([]);
  });
});
