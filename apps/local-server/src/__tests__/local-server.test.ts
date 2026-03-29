import { afterAll, beforeAll, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { XLSX_CONTENT_TYPE, decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import type { ProtocolFrame } from "@bilig/binary-protocol";
import { ValueTag } from "@bilig/protocol";

import { createLocalServer } from "../server.js";
import { LocalWorkbookSessionManager } from "../local-workbook-session-manager.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function parseSnapshotPayload(value: unknown): {
  workbook: { name: string };
  sheets: Array<{ cells: Array<{ address: string; value?: number; formula?: string }> }>;
} {
  if (
    !isRecord(value) ||
    !isRecord(value["workbook"]) ||
    typeof value["workbook"]["name"] !== "string" ||
    !Array.isArray(value["sheets"])
  ) {
    throw new Error("Invalid snapshot payload");
  }
  return {
    workbook: { name: value["workbook"]["name"] },
    sheets: value["sheets"].map((sheet) => {
      if (!isRecord(sheet) || !Array.isArray(sheet["cells"])) {
        throw new Error("Invalid snapshot payload");
      }
      return {
        cells: sheet["cells"].map((cell) => {
          if (!isRecord(cell) || typeof cell["address"] !== "string") {
            throw new Error("Invalid snapshot payload");
          }
          return {
            address: cell["address"],
            ...(typeof cell["value"] === "number" ? { value: cell["value"] } : {}),
            ...(typeof cell["formula"] === "string" ? { formula: cell["formula"] } : {}),
          };
        }),
      };
    }),
  };
}

function buildWorkbookUploadBase64(): string {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    [11, 5],
    [null, null],
  ]);
  sheet["C1"] = { t: "n", f: "A1+B1" };
  sheet["!ref"] = "A1:C2";
  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  return Buffer.from(XLSX.write(workbook, { bookType: "xlsx", type: "buffer" })).toString("base64");
}

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
      url: "/v2/agent/frames",
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
      url: "/v2/agent/frames",
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
      url: "/v2/agent/frames",
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

  it("adds CORS headers for browser access to document snapshots", async () => {
    const response = await app.inject({
      method: "GET",
      url: "/v2/documents/cors-doc/snapshot/latest",
      headers: {
        origin: "http://localhost:3000",
      },
    });

    expect(response.headers["access-control-allow-origin"]).toBe("http://localhost:3000");
    expect(response.headers["access-control-allow-credentials"]).toBe("true");
    expect(response.headers["access-control-expose-headers"]).toBe("x-bilig-snapshot-cursor");
    expect(response.headers["vary"]).toBe("origin");
  });

  it("rejects streaming subscriptions over the non-streaming HTTP agent endpoint", async () => {
    const openResponse = await app.inject({
      method: "POST",
      url: "/v2/agent/frames",
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
      url: "/v2/agent/frames",
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
        kind: "appendBatch",
        documentId: "book-2",
        cursor: 1,
        batch: {
          id: "local-server:book-2:1",
          replicaId: "local-server:book-2",
          clock: { counter: 1 },
          ops: [{ kind: "upsertSheet", name: "Sheet1", order: 0 }],
        },
      },
      {
        kind: "cursorWatermark",
        documentId: "book-2",
        cursor: 1,
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
    expect(committed.cursor).toBe(2);

    expect(responses).toHaveLength(1);
    const ack = responses[0];
    expect(ack.kind).toBe("ack");
    if (ack.kind !== "ack") throw new Error("Expected ack frame");
    expect(ack.cursor).toBe(2);
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

  it("imports xlsx uploads through the agent API and exposes the imported snapshot", async () => {
    const uploadResponse = await app.inject({
      method: "POST",
      url: "/v2/agent/frames",
      headers: { "content-type": "application/octet-stream" },
      payload: Buffer.from(
        encodeAgentFrame({
          kind: "request",
          request: {
            kind: "loadWorkbookFile",
            id: "upload-1",
            replicaId: "agent-upload",
            openMode: "create",
            fileName: "report.xlsx",
            contentType: XLSX_CONTENT_TYPE,
            bytesBase64: buildWorkbookUploadBase64(),
          },
        }),
      ),
    });

    const uploadFrame = decodeAgentFrame(uploadResponse.rawPayload);
    expect(uploadFrame.kind).toBe("response");
    if (uploadFrame.kind !== "response" || uploadFrame.response.kind !== "workbookLoaded") {
      throw new Error("Expected workbookLoaded response");
    }
    expect(uploadFrame.response.workbookName).toBe("report");
    expect(uploadFrame.response.sheetNames).toEqual(["Sheet1"]);

    const snapshotResponse = await app.inject({
      method: "GET",
      url: `/v2/documents/${encodeURIComponent(uploadFrame.response.documentId)}/snapshot/latest`,
    });
    expect(snapshotResponse.statusCode).toBe(200);
    const snapshot = parseSnapshotPayload(JSON.parse(snapshotResponse.body) as unknown);
    expect(snapshot.workbook.name).toBe("report");
    expect(snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: "A1", value: 11 })]),
    );
    expect(snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: "C1", formula: "A1+B1" })]),
    );
  });

  it("broadcasts snapshot chunks to connected browsers when replacing a workbook from xlsx", async () => {
    const manager = new LocalWorkbookSessionManager();
    const broadcasts: ProtocolFrame[] = [];
    const detach = manager.attachBrowser("replace-doc", "browser-sub", (frame) => {
      broadcasts.push(frame);
    });

    const response = await manager.handleAgentFrame({
      kind: "request",
      request: {
        kind: "loadWorkbookFile",
        id: "replace-1",
        replicaId: "agent-replace",
        openMode: "replace",
        documentId: "replace-doc",
        fileName: "replace.xlsx",
        contentType: XLSX_CONTENT_TYPE,
        bytesBase64: buildWorkbookUploadBase64(),
      },
    });

    expect(response).toMatchObject({
      kind: "response",
      response: {
        kind: "workbookLoaded",
        documentId: "replace-doc",
      },
    });
    expect(broadcasts.some((frame) => frame.kind === "snapshotChunk")).toBe(true);
    expect(broadcasts.at(-1)).toMatchObject({
      kind: "cursorWatermark",
      documentId: "replace-doc",
    });

    detach();
  });
});
