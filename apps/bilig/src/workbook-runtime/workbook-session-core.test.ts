import { describe, expect, it, vi } from "vitest";
import type { AgentFrame } from "@bilig/agent-api";
import type { HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import { WorkbookBrowserSessionHost } from "./browser-session-host.js";
import { WorkbookSessionCore } from "./workbook-session-core.js";
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

describe("WorkbookSessionCore", () => {
  it("routes browser, sync, and agent traffic through shared handlers", async () => {
    const browserSessionHost = new WorkbookBrowserSessionHost({
      latestCursor: async () => 0,
      latestSnapshot: async () => null,
      listMissedFrames: async () => [],
    });
    const syncSessionHost = new WorkbookSyncSessionHost<ProtocolFrame[]>({
      browserSessionHost,
      hello: async () => [
        { kind: "ack", documentId: "doc-1", batchId: "hello", cursor: 0, acceptedAtUnixMs: 1 },
      ],
      appendBatch: async () => [
        { kind: "ack", documentId: "doc-1", batchId: "append", cursor: 1, acceptedAtUnixMs: 2 },
      ],
      snapshotChunk: async () => [
        { kind: "ack", documentId: "doc-1", batchId: "snapshot", cursor: 2, acceptedAtUnixMs: 3 },
      ],
      heartbeat: async () => [
        { kind: "heartbeat", documentId: "doc-1", cursor: 4, sentAtUnixMs: 4 },
      ],
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
    const handleWorksheetRequest = vi.fn(
      async (_request: unknown) =>
        ({
          kind: "response",
          response: { kind: "ok", id: "worksheet-1" },
        }) satisfies AgentFrame,
    );
    const core = new WorkbookSessionCore({
      syncSessionHost,
      invalidFrameMessage: "Only requests are supported",
      errorCode: "SESSION_FAILURE",
      loadWorkbookFile: async (request) => ({
        kind: "response",
        response: { kind: "ok", id: request.id },
      }),
      openWorkbookSession: async (request) => `${request.documentId}:${request.replicaId}`,
      closeWorkbookSession: async () => undefined,
      getMetrics: async (request) => ({
        kind: "metrics",
        id: request.id,
        value: { service: "test" },
      }),
      handleWorksheetRequest: (_frame, request) => handleWorksheetRequest(request),
    });

    const detach = core.attachBrowser("doc-1", "browser-1", () => {});
    expect(core.listSubscriberIds("doc-1")).toEqual(["browser-1"]);
    detach();

    await expect(core.openBrowserSession(helloFrame)).resolves.toEqual([
      {
        kind: "cursorWatermark",
        documentId: "doc-1",
        cursor: 0,
        compactedCursor: 0,
      },
    ]);

    await expect(
      core.handleSyncFrame({
        kind: "heartbeat",
        documentId: "doc-1",
        cursor: 1,
        sentAtUnixMs: 1,
      }),
    ).resolves.toEqual([
      {
        kind: "heartbeat",
        documentId: "doc-1",
        cursor: 4,
        sentAtUnixMs: 4,
      },
    ]);

    const openResponse = await core.handleAgentFrame({
      kind: "request",
      request: {
        kind: "openWorkbookSession",
        id: "open-1",
        documentId: "doc-1",
        replicaId: "replica-1",
      },
    });
    expect(openResponse).toEqual({
      kind: "response",
      response: {
        kind: "ok",
        id: "open-1",
        sessionId: "doc-1:replica-1",
      },
    });

    await core.handleAgentFrame({
      kind: "request",
      request: {
        kind: "readRange",
        id: "worksheet-1",
        sessionId: "doc-1:replica-1",
        range: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A1",
        },
      },
    });
    expect(handleWorksheetRequest).toHaveBeenCalledOnce();
  });
});
