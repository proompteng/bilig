import { describe, expect, it } from "vitest";

import { SpreadsheetEngine } from "@bilig/core";
import {
  WORKBOOK_SNAPSHOT_CONTENT_TYPE,
  createSnapshotChunkFrames,
  decodeFrame,
  encodeFrame,
} from "@bilig/binary-protocol";
import { ValueTag } from "@bilig/protocol";
import type { WorkbookSnapshot } from "@bilig/protocol";

import { createWebSocketSyncClient, type BrowserWebSocketLike } from "../websocket-sync-client.js";

class FakeSocket extends EventTarget implements BrowserWebSocketLike {
  binaryType = "arraybuffer";
  readyState = 0;
  readonly sent: Uint8Array[] = [];

  send(data: ArrayBufferLike | ArrayBufferView): void {
    if (data instanceof Uint8Array) {
      this.sent.push(data);
      return;
    }
    if (ArrayBuffer.isView(data)) {
      this.sent.push(new Uint8Array(data.buffer, data.byteOffset, data.byteLength));
      return;
    }
    this.sent.push(new Uint8Array(data));
  }

  close(): void {
    this.readyState = 3;
    this.dispatchEvent(new Event("close"));
  }

  open(): void {
    this.readyState = 1;
    this.dispatchEvent(new Event("open"));
  }

  push(bytes: Uint8Array): void {
    this.dispatchEvent(new MessageEvent<Uint8Array>("message", { data: bytes }));
  }
}

describe("package websocket sync client", () => {
  it("applies remote snapshot chunks during sync", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "remote-doc", replicaId: "browser:test" });
    await engine.ready();
    const socket = new FakeSocket();
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: "Imported" },
      sheets: [
        {
          name: "Sheet1",
          order: 0,
          cells: [{ address: "A1", value: 42 }],
        },
      ],
    };

    const connectPromise = engine.connectSyncClient(
      createWebSocketSyncClient({
        documentId: "remote-doc",
        replicaId: "browser:test",
        baseUrl: "http://127.0.0.1:4321",
        createSocket: () => socket,
      }),
    );

    await Promise.resolve();
    socket.open();
    await Promise.resolve();
    const firstMessage = socket.sent[0];
    if (!firstMessage) {
      throw new Error("Expected hello frame from sync client");
    }
    const helloFrame = decodeFrame(firstMessage);
    expect(helloFrame.kind).toBe("hello");

    const bytes = new TextEncoder().encode(JSON.stringify(snapshot));
    const frames = createSnapshotChunkFrames({
      documentId: "remote-doc",
      snapshotId: "snap-1",
      cursor: 1,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes,
      chunkSize: 4,
    });
    frames.forEach((frame) => socket.push(encodeFrame(frame)));
    await connectPromise;

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({
      tag: ValueTag.Number,
      value: 42,
    });
    expect(engine.getSyncState()).toBe("live");
  });
});
