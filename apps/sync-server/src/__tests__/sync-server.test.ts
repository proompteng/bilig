import { describe, expect, it } from "vitest";

import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";

import { createSyncServer } from "../server.js";

describe("sync-server", () => {
  it("accepts hello frames and returns cursor watermarks", async () => {
    const { app } = createSyncServer();
    const response = await app.inject({
      method: "POST",
      url: "/v1/frames",
      headers: {
        "content-type": "application/octet-stream"
      },
      payload: Buffer.from(encodeFrame({
        kind: "hello",
        documentId: "book-1",
        replicaId: "browser-a",
        sessionId: "sess-1",
        protocolVersion: 1,
        lastServerCursor: 0,
        capabilities: ["sync"]
      }))
    });

    expect(response.statusCode).toBe(200);
    const frame = decodeFrame(response.rawPayload);
    expect(frame).toEqual({
      kind: "cursorWatermark",
      documentId: "book-1",
      cursor: 0,
      compactedCursor: 0
    });

    await app.close();
  });

  it("persists appendBatch frames and acknowledges them", async () => {
    const { app, sessionManager } = createSyncServer();

    const response = await app.inject({
      method: "POST",
      url: "/v1/frames",
      headers: {
        "content-type": "application/octet-stream"
      },
      payload: Buffer.from(encodeFrame({
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
              name: "spec"
            }
          ]
        }
      }))
    });

    const frame = decodeFrame(response.rawPayload);
    expect(frame.kind).toBe("ack");
    expect(await sessionManager.getDocumentState("book-1")).toMatchObject({ cursor: 1 });

    await app.close();
  });
});
