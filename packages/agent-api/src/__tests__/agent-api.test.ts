import { describe, expect, it } from "vitest";

import {
  decodeAgentFrame,
  decodeStdioMessages,
  encodeAgentFrame,
  encodeStdioMessage,
  XLSX_CONTENT_TYPE,
} from "../index.js";

describe("agent api", () => {
  it("roundtrips request and response envelopes", () => {
    const frame = {
      kind: "request" as const,
      request: {
        kind: "openWorkbookSession" as const,
        id: "req-1",
        documentId: "book-1",
        replicaId: "agent-local",
      },
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
          sessionId: "sess-1",
        },
      }),
      ...encodeStdioMessage({
        kind: "response",
        response: {
          kind: "ok",
          id: "req-1",
          value: { ok: true },
        },
      }),
    ]);

    const decoded = decodeStdioMessages(encoded);
    expect(decoded.frames).toHaveLength(2);
    expect(decoded.remainder.byteLength).toBe(0);
  });

  it("roundtrips workbook file load requests and responses", () => {
    const requestFrame = {
      kind: "request" as const,
      request: {
        kind: "loadWorkbookFile" as const,
        id: "upload-1",
        replicaId: "agent-local",
        openMode: "create" as const,
        fileName: "report.xlsx",
        contentType: XLSX_CONTENT_TYPE,
        bytesBase64: "QUJD",
      },
    };
    expect(decodeAgentFrame(encodeAgentFrame(requestFrame))).toEqual(requestFrame);

    const responseFrame = {
      kind: "response" as const,
      response: {
        kind: "workbookLoaded" as const,
        id: "upload-1",
        documentId: "xlsx:abc123",
        sessionId: "xlsx:abc123:agent-local",
        workbookName: "report.xlsx",
        sheetNames: ["Sheet1"],
        serverUrl: "http://127.0.0.1:4381",
        browserUrl: "http://127.0.0.1:4173/?document=xlsx%3Aabc123",
        warnings: [],
      },
    };
    expect(decodeAgentFrame(encodeAgentFrame(responseFrame))).toEqual(responseFrame);
  });
});
