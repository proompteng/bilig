import { describe, expect, it } from "vitest";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";

import { createLocalServer } from "../../../local-server/src/server.js";
import { createSyncServer } from "../server.js";
import { createHttpWorksheetExecutor } from "../worksheet-executor.js";

describe("sync-server", () => {
  it("accepts hello frames and returns cursor watermarks", async () => {
    const { app } = createSyncServer();
    const response = await app.inject({
      method: "POST",
      url: "/v1/frames",
      headers: {
        "content-type": "application/octet-stream",
      },
      payload: Buffer.from(
        encodeFrame({
          kind: "hello",
          documentId: "book-1",
          replicaId: "browser-a",
          sessionId: "sess-1",
          protocolVersion: 1,
          lastServerCursor: 0,
          capabilities: ["sync"],
        }),
      ),
    });

    expect(response.statusCode).toBe(200);
    const frame = decodeFrame(response.rawPayload);
    expect(frame).toEqual({
      kind: "cursorWatermark",
      documentId: "book-1",
      cursor: 0,
      compactedCursor: 0,
    });

    await app.close();
  });

  it("persists appendBatch frames and acknowledges them", async () => {
    const { app, sessionManager } = createSyncServer();

    const response = await app.inject({
      method: "POST",
      url: "/v1/frames",
      headers: {
        "content-type": "application/octet-stream",
      },
      payload: Buffer.from(
        encodeFrame({
          kind: "appendBatch",
          documentId: "book-1",
          cursor: 0,
          batch: {
            id: "replica:1",
            replicaId: "replica",
            clock: { counter: 1 },
            ops: [
              {
                kind: "upsertWorkbook",
                name: "spec",
              },
            ],
          },
        }),
      ),
    });

    const frame = decodeFrame(response.rawPayload);
    expect(frame.kind).toBe("ack");
    expect(await sessionManager.getDocumentState("book-1")).toMatchObject({ cursor: 1 });

    await app.close();
  });

  it("delegates worksheet mutations to the live local worksheet host when configured", async () => {
    const { app: localApp } = createLocalServer({ logger: false });
    await localApp.listen({ port: 0, host: "127.0.0.1" });
    const localAddress = new URL(
      localApp.server.address() && typeof localApp.server.address() === "object"
        ? `http://127.0.0.1:${localApp.server.address().port}`
        : "http://127.0.0.1:4381",
    );
    const { app } = createSyncServer({
      logger: false,
      worksheetExecutor: createHttpWorksheetExecutor({ baseUrl: localAddress.toString() }),
    });

    const openResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "openWorkbookSession",
            id: "open-remote",
            documentId: "remote-doc",
            replicaId: "remote-agent",
          },
        }),
      ),
    });
    const openFrame = decodeAgentFrame(openResponse.rawPayload);
    expect(openFrame).toMatchObject({
      kind: "response",
      response: {
        kind: "ok",
        sessionId: "remote-doc:remote-agent",
      },
    });

    const writeResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "writeRange",
            id: "write-remote",
            sessionId: "remote-doc:remote-agent",
            range: {
              sheetName: "Sheet1",
              startAddress: "A1",
              endAddress: "A1",
            },
            values: [[99]],
          },
        }),
      ),
    });
    const writeFrame = decodeAgentFrame(writeResponse.rawPayload);
    expect(writeFrame).toMatchObject({
      kind: "response",
      response: {
        kind: "ok",
      },
    });

    const readResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "readRange",
            id: "read-remote",
            sessionId: "remote-doc:remote-agent",
            range: {
              sheetName: "Sheet1",
              startAddress: "A1",
              endAddress: "A1",
            },
          },
        }),
      ),
    });
    const readFrame = decodeAgentFrame(readResponse.rawPayload);
    expect(readFrame).toMatchObject({
      kind: "response",
      response: {
        kind: "rangeValues",
      },
    });

    await app.close();
    await localApp.close();
  });

  it("rejects streaming subscriptions when delegated over the HTTP worksheet executor", async () => {
    const { app: localApp } = createLocalServer({ logger: false });
    await localApp.listen({ port: 0, host: "127.0.0.1" });
    const localAddress = new URL(
      localApp.server.address() && typeof localApp.server.address() === "object"
        ? `http://127.0.0.1:${localApp.server.address().port}`
        : "http://127.0.0.1:4381",
    );
    const { app } = createSyncServer({
      logger: false,
      worksheetExecutor: createHttpWorksheetExecutor({ baseUrl: localAddress.toString() }),
    });

    const openResponse = await app.inject({
      method: "POST",
      url: "/v1/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "openWorkbookSession",
            id: "open-remote-stream",
            documentId: "remote-stream-doc",
            replicaId: "remote-stream-agent",
          },
        }),
      ),
    });
    const openFrame = decodeAgentFrame(openResponse.rawPayload);
    expect(openFrame).toMatchObject({
      kind: "response",
      response: {
        kind: "ok",
        sessionId: "remote-stream-doc:remote-stream-agent",
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
            id: "subscribe-remote-stream",
            sessionId: "remote-stream-doc:remote-stream-agent",
            subscriptionId: "remote-sub-1",
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
        id: "subscribe-remote-stream",
        code: "AGENT_STREAM_REQUIRES_STREAMING_TRANSPORT",
        message: "subscribeRange requires a streaming agent transport such as stdio",
        retryable: false,
      },
    });

    await app.close();
    await localApp.close();
  });
});
