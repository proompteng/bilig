import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import type { ProtocolFrame } from "@bilig/binary-protocol";
import { ValueTag } from "@bilig/protocol";

import { createLocalServer } from "../server.js";
import { LocalWorkbookSessionManager } from "../local-workbook-session-manager.js";

describe("local-server", () => {
  const { app } = createLocalServer();

  beforeAll(async () => {
    await app.listen({ port: 0, host: "127.0.0.1" });
  });

  afterAll(async () => {
    await app.close();
  });

  it("opens a live workbook session and applies range writes through the agent API", async () => {
    const openResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "openWorkbookSession",
            id: "open-1",
            documentId: "book-1",
            replicaId: "agent-a",
          },
        }),
      ),
    });

    const openFrame = decodeAgentFrame(openResponse.rawPayload);
    expect(openFrame.kind).toBe("response");
    if (openFrame.kind !== "response" || openFrame.response.kind !== "ok") {
      throw new Error("Expected ok agent response");
    }
    expect(openFrame.response.sessionId).toBe("book-1:agent-a");

    const writeResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "writeRange",
            id: "write-1",
            sessionId: "book-1:agent-a",
            range: {
              sheetName: "Sheet1",
              startAddress: "A1",
              endAddress: "B2",
            },
            values: [
              [1, 2],
              [3, 4],
            ],
          },
        }),
      ),
    });

    const writeFrame = decodeAgentFrame(writeResponse.rawPayload);
    expect(writeFrame.kind).toBe("response");
    if (writeFrame.kind !== "response" || writeFrame.response.kind !== "ok") {
      throw new Error("Expected ok write response");
    }

    const readResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "readRange",
            id: "read-1",
            sessionId: "book-1:agent-a",
            range: {
              sheetName: "Sheet1",
              startAddress: "A1",
              endAddress: "B2",
            },
          },
        }),
      ),
    });

    const readFrame = decodeAgentFrame(readResponse.rawPayload);
    expect(readFrame.kind).toBe("response");
    if (readFrame.kind !== "response" || readFrame.response.kind !== "rangeValues") {
      throw new Error("Expected rangeValues response");
    }
    expect(
      readFrame.response.values.map((row) =>
        row.map((cell) => (cell.tag === ValueTag.Number ? cell.value : null)),
      ),
    ).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  it("rejects streaming subscriptions over the non-streaming HTTP agent endpoint", async () => {
    const openResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "openWorkbookSession",
            id: "open-http-stream",
            documentId: "http-stream-doc",
            replicaId: "agent-http",
          },
        }),
      ),
    });

    const openFrame = decodeAgentFrame(openResponse.rawPayload);
    expect(openFrame).toMatchObject({
      kind: "response",
      response: {
        kind: "ok",
        sessionId: "http-stream-doc:agent-http",
      },
    });

    const subscribeResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "subscribeRange",
            id: "subscribe-http-stream",
            sessionId: "http-stream-doc:agent-http",
            subscriptionId: "sub-http",
            range: {
              sheetName: "Sheet1",
              startAddress: "A1",
              endAddress: "A1",
            },
          },
        }),
      ),
    });

    const subscribeFrame = decodeAgentFrame(subscribeResponse.rawPayload);
    expect(subscribeFrame).toEqual({
      kind: "response",
      response: {
        kind: "error",
        id: "subscribe-http-stream",
        code: "AGENT_STREAM_REQUIRES_STREAMING_TRANSPORT",
        message: "subscribeRange requires a streaming agent transport such as stdio",
        retryable: false,
      },
    });
  });

  it("acknowledges committed browser batches and broadcasts them through the local session manager", async () => {
    const manager = new LocalWorkbookSessionManager();
    const broadcasts: ProtocolFrame[] = [];
    const detach = manager.attachBrowser("book-2", "browser-sub", (frame) => {
      broadcasts.push(frame);
    });

    const helloFrames = await manager.handleSyncFrame({
      kind: "hello",
      documentId: "book-2",
      replicaId: "browser-a",
      sessionId: "browser-a",
      protocolVersion: 1,
      lastServerCursor: 0,
      capabilities: ["local-session"],
    });

    expect(helloFrames).toEqual([
      {
        kind: "cursorWatermark",
        documentId: "book-2",
        cursor: 0,
        compactedCursor: 0,
      },
    ]);

    const responses = await manager.handleSyncFrame({
      kind: "appendBatch",
      documentId: "book-2",
      cursor: 0,
      batch: {
        id: "browser-a:1",
        replicaId: "browser-a",
        clock: { counter: 1 },
        ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 42 }],
      },
    });

    expect(broadcasts).toHaveLength(1);
    const committed = broadcasts[0];
    expect(committed.kind).toBe("appendBatch");
    if (committed.kind !== "appendBatch") throw new Error("Expected appendBatch broadcast");
    expect(committed.cursor).toBe(1);

    expect(responses).toHaveLength(1);
    const ack = responses[0];
    expect(ack.kind).toBe("ack");
    if (ack.kind !== "ack") throw new Error("Expected ack frame");
    expect(ack.cursor).toBe(1);
    detach();
  });

  it("relays both agent and browser batches upstream when a sync relay is configured", async () => {
    const relayedBatchIds: string[] = [];
    const manager = new LocalWorkbookSessionManager({
      createSyncRelay: () => ({
        async send(batch) {
          relayedBatchIds.push(batch.id);
        },
        async disconnect() {},
      }),
    });

    await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "openWorkbookSession",
        id: "open-relay",
        documentId: "relay-doc",
        replicaId: "agent-relay",
      },
    });

    await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "writeRange",
        id: "write-relay",
        sessionId: "relay-doc:agent-relay",
        range: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A1",
        },
        values: [[7]],
      },
    });
    await Promise.resolve();

    await manager.handleSyncFrame({
      kind: "appendBatch",
      documentId: "relay-doc",
      cursor: 0,
      batch: {
        id: "browser-relay:1",
        replicaId: "browser-relay",
        clock: { counter: 1 },
        ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "B1", value: 9 }],
      },
    });
    await Promise.resolve();

    expect(relayedBatchIds).toContain("local-server:relay-doc:1");
    expect(relayedBatchIds).toContain("browser-relay:1");
  });
});
