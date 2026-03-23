import { describe, expect, it } from "vitest";

import { SpreadsheetEngine } from "@bilig/core";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";
import { ValueTag } from "@bilig/protocol";

import { createWebSocketSyncClient, type BrowserWebSocketLike } from "../createWebSocketSyncClient.js";

class FakeSocket implements BrowserWebSocketLike {
  binaryType = "arraybuffer";
  readyState = 0;
  readonly sent: Uint8Array[] = [];
  private listeners: {
    open: ((event: Event) => void) | null;
    message: ((event: MessageEvent<unknown>) => void) | null;
    error: ((event: Event) => void) | null;
    close: ((event: Event) => void) | null;
  } = {
    open: null,
    message: null,
    error: null,
    close: null
  };

  addEventListener<K extends "open" | "message" | "error" | "close">(
    type: K,
    listener: typeof this.listeners[K]
  ): void {
    this.listeners[type] = listener;
  }

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
    this.listeners.close?.(new Event("close"));
  }

  open(): void {
    this.readyState = 1;
    this.listeners.open?.(new Event("open"));
  }

  push(frame: Parameters<typeof encodeFrame>[0]): void {
    this.listeners.message?.(new MessageEvent<Uint8Array>("message", { data: encodeFrame(frame) }));
  }
}

describe("websocket sync client", () => {
  it("handshakes, applies remote batches, and forwards local batches", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "remote-doc", replicaId: "browser:test" });
    await engine.ready();
    const socket = new FakeSocket();

    const connectPromise = engine.connectSyncClient(createWebSocketSyncClient({
      documentId: "remote-doc",
      replicaId: "browser:test",
      baseUrl: "http://127.0.0.1:4381",
      createSocket: () => socket
    }));

    await Promise.resolve();
    socket.open();
    await Promise.resolve();
    const helloBytes = socket.sent.find((entry) => entry.byteLength > 0);
    expect(helloBytes).toBeDefined();
    const helloFrame = decodeFrame(helloBytes!);
    expect(helloFrame.kind).toBe("hello");

    socket.push({
      kind: "cursorWatermark",
      documentId: "remote-doc",
      cursor: 0,
      compactedCursor: 0
    });
    await connectPromise;
    expect(engine.getSyncState()).toBe("live");

    socket.push({
      kind: "appendBatch",
      documentId: "remote-doc",
      cursor: 1,
      batch: {
        id: "remote:1",
        replicaId: "remote",
        clock: { counter: 1 },
        ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "B1", value: 9 }]
      }
    });
    await Promise.resolve();
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({
      tag: ValueTag.Number,
      value: 9
    });

    engine.setCellValue("Sheet1", "A1", 42);
    const appendFrame = decodeFrame(socket.sent.at(-1)!);
    expect(appendFrame.kind).toBe("appendBatch");
    if (appendFrame.kind !== "appendBatch") {
      throw new Error("Expected appendBatch frame");
    }
    expect(appendFrame.batch.ops).toEqual([
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 42 }
    ]);
  });
});
