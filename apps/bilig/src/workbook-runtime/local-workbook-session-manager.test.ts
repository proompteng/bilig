import { describe, expect, it } from "vitest";
import type { ProtocolFrame } from "@bilig/binary-protocol";
import type { AgentFrame } from "@bilig/agent-api";
import { LocalWorkbookSessionManager } from "./local-workbook-session-manager.js";

async function openSession(
  manager: LocalWorkbookSessionManager,
  documentId: string,
): Promise<string> {
  const response = await manager.handleAgentFrame({
    kind: "request",
    request: {
      kind: "openWorkbookSession",
      id: "open-1",
      documentId,
      replicaId: "replica-1",
    },
  } satisfies AgentFrame);

  if (response.kind !== "response" || response.response.kind !== "ok") {
    throw new Error("Expected workbook session to open");
  }
  if (!response.response.sessionId) {
    throw new Error("Expected workbook session id");
  }

  return response.response.sessionId;
}

describe("LocalWorkbookSessionManager", () => {
  it("tracks browser subscribers independently from workbook sessions", async () => {
    const manager = new LocalWorkbookSessionManager();
    const detach = manager.attachBrowser("doc-browser", "browser-1", () => {});

    expect(manager.getDocumentState("doc-browser").browserSessions).toEqual(["browser-1"]);

    detach();

    expect(manager.getDocumentState("doc-browser").browserSessions).toEqual([]);
  });

  it("compacts long batch backlogs into a snapshot", async () => {
    const manager = new LocalWorkbookSessionManager();
    const sessionId = await openSession(manager, "doc-perf");

    await Promise.all(
      Array.from({ length: 270 }, (_entry, index) =>
        manager.handleAgentFrame({
          kind: "request",
          request: {
            kind: "writeRange",
            id: `write-${index}`,
            sessionId,
            range: {
              sheetName: "Sheet1",
              startAddress: "A1",
              endAddress: "A1",
            },
            values: [[index]],
          },
        } satisfies AgentFrame),
      ),
    );

    await new Promise((resolve) => setImmediate(resolve));

    const snapshot = manager.getLatestSnapshot("doc-perf");
    expect(snapshot).not.toBeNull();

    const helloFrames = await manager.handleSyncFrame({
      kind: "hello",
      documentId: "doc-perf",
      replicaId: "browser-1",
      sessionId: "browser-1",
      protocolVersion: 1,
      lastServerCursor: 0,
      capabilities: [],
    } satisfies ProtocolFrame);

    expect(helloFrames.some((frame) => frame.kind === "snapshotChunk")).toBe(true);
    const appendFrames = helloFrames.filter((frame) => frame.kind === "appendBatch");
    expect(appendFrames.length).toBeLessThanOrEqual(256);
  });

  it("emits large range subscription events for style-only invalidations", async () => {
    const manager = new LocalWorkbookSessionManager();
    const sessionId = await openSession(manager, "doc-range");
    const events: AgentFrame[] = [];
    const unsubscribeEvents = manager.subscribeAgentEvents((event) => {
      events.push({ kind: "event", event });
    });

    await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "subscribeRange",
        id: "subscribe-1",
        sessionId,
        subscriptionId: "sub-1",
        range: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "Z20",
        },
      },
    } satisfies AgentFrame);

    await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "setRangeStyle",
        id: "style-1",
        sessionId,
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B2",
        },
        patch: {
          fill: { backgroundColor: "#ff0000" },
        },
      },
    } satisfies AgentFrame);

    await new Promise((resolve) => setImmediate(resolve));

    expect(events).toContainEqual({
      kind: "event",
      event: {
        kind: "rangeChanged",
        subscriptionId: "sub-1",
        range: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "Z20",
        },
        changedAddresses: ["B2"],
      },
    });

    unsubscribeEvents();
  });

  it("reuses cached snapshots for repeated export requests until a write occurs", async () => {
    const manager = new LocalWorkbookSessionManager();
    const sessionId = await openSession(manager, "doc-snapshot-cache");

    const firstResponse = await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "exportSnapshot",
        id: "export-1",
        sessionId,
      },
    } satisfies AgentFrame);
    const secondResponse = await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "exportSnapshot",
        id: "export-2",
        sessionId,
      },
    } satisfies AgentFrame);

    if (
      firstResponse.kind !== "response" ||
      firstResponse.response.kind !== "snapshot" ||
      secondResponse.kind !== "response" ||
      secondResponse.response.kind !== "snapshot"
    ) {
      throw new Error("Expected snapshot responses");
    }

    expect(secondResponse.response.snapshot).toBe(firstResponse.response.snapshot);

    await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "writeRange",
        id: "write-1",
        sessionId,
        range: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A1",
        },
        values: [[123]],
      },
    } satisfies AgentFrame);

    const thirdResponse = await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "exportSnapshot",
        id: "export-3",
        sessionId,
      },
    } satisfies AgentFrame);

    if (thirdResponse.kind !== "response" || thirdResponse.response.kind !== "snapshot") {
      throw new Error("Expected snapshot response after write");
    }

    expect(thirdResponse.response.snapshot).not.toBe(firstResponse.response.snapshot);
    expect(thirdResponse.response.snapshot.sheets[0]?.cells).toContainEqual(
      expect.objectContaining({ address: "A1", value: 123 }),
    );
  });
});
