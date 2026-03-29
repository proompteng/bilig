import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { SpreadsheetEngine } from "@bilig/core";
import {
  WORKBOOK_SNAPSHOT_CONTENT_TYPE,
  createSnapshotChunkFrames,
  decodeFrame,
  encodeFrame,
} from "@bilig/binary-protocol";
import { ValueTag, type WorkbookSnapshot } from "@bilig/protocol";
import { createWebSocketSyncClient, type BrowserWebSocketLike } from "../websocket-sync-client.js";
import { runProperty } from "@bilig/test-fuzz";

class FakeSocket extends EventTarget implements BrowserWebSocketLike {
  binaryType: BinaryType = "arraybuffer";
  readyState = 0;
  readonly sent: Uint8Array[] = [];

  send(data: BufferSource): void {
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

function buildSnapshot(
  cells: ReadonlyArray<{ address: string; value: number | string | boolean }>,
): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: "fuzz-transport-book" },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: cells.map((cell) => ({
          address: cell.address,
          value: cell.value,
        })),
      },
    ],
  };
}

async function openSyncConnection(
  snapshotBytes: Uint8Array,
  contentType: string,
  chunkSize: number,
): Promise<{ engine: SpreadsheetEngine; connectPromise: Promise<void>; socket: FakeSocket }> {
  const engine = new SpreadsheetEngine({
    workbookName: "transport-doc",
    replicaId: "browser:fuzz",
  });
  await engine.ready();
  const socket = new FakeSocket();
  const connectPromise = engine
    .connectSyncClient(
      createWebSocketSyncClient({
        documentId: "transport-doc",
        replicaId: "browser:fuzz",
        baseUrl: "http://127.0.0.1:4321",
        createSocket: () => socket,
      }),
    )
    .then(() => undefined);

  await Promise.resolve();
  socket.open();
  await Promise.resolve();

  const helloFrame = socket.sent[0];
  if (!helloFrame) {
    throw new Error("Expected hello frame from websocket sync client");
  }
  expect(decodeFrame(helloFrame).kind).toBe("hello");

  const frames = createSnapshotChunkFrames({
    documentId: "transport-doc",
    snapshotId: "fuzz-snapshot",
    cursor: 1,
    contentType,
    bytes: snapshotBytes,
    chunkSize,
  });
  frames.forEach((frame) => socket.push(encodeFrame(frame)));

  return { engine, connectPromise, socket };
}

describe("websocket sync client fuzz", () => {
  it("reconstructs remote workbook snapshots across random chunk sizes", async () => {
    await runProperty({
      suite: "worker-transport/websocket-sync-client/snapshot-chunks",
      arbitrary: fc.record({
        chunkSize: fc.integer({ min: 1, max: 32 }),
        cells: fc.uniqueArray(
          fc.record({
            address: fc.constantFrom("A1", "A2", "B1", "B2", "C3", "D4"),
            value: fc.oneof(
              fc.integer({ min: 1, max: 999 }),
              fc.string({ minLength: 1, maxLength: 4 }),
              fc.boolean(),
            ),
          }),
          {
            minLength: 1,
            maxLength: 4,
            selector: (entry) => entry.address,
          },
        ),
      }),
      predicate: async ({ chunkSize, cells }) => {
        const snapshot = buildSnapshot(cells);
        const { engine, connectPromise } = await openSyncConnection(
          new TextEncoder().encode(JSON.stringify(snapshot)),
          WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          chunkSize,
        );
        await connectPromise;
        cells.forEach((cell) => {
          const value = engine.getCellValue("Sheet1", cell.address);
          const expectedValue =
            typeof cell.value === "number"
              ? { tag: ValueTag.Number, value: cell.value }
              : typeof cell.value === "boolean"
                ? { tag: ValueTag.Boolean, value: cell.value }
                : { tag: ValueTag.String, value: cell.value };
          const comparableValue =
            typeof cell.value === "string" && value !== null && typeof value === "object"
              ? { tag: value.tag, value: value.value }
              : value;
          expect(comparableValue).toEqual(expectedValue);
        });
        expect(engine.getSyncState()).toBe("live");
      },
    });
  });

  it("rejects malformed snapshot payloads deterministically", async () => {
    await runProperty({
      suite: "worker-transport/websocket-sync-client/invalid-payloads",
      arbitrary: fc.record({
        chunkSize: fc.integer({ min: 1, max: 16 }),
        failureMode: fc.constantFrom("invalid-json", "wrong-content-type"),
      }),
      predicate: async ({ chunkSize, failureMode }) => {
        const { engine, connectPromise } = await openSyncConnection(
          failureMode === "invalid-json"
            ? new TextEncoder().encode('{"broken":')
            : new TextEncoder().encode(
                JSON.stringify(buildSnapshot([{ address: "A1", value: 7 }])),
              ),
          failureMode === "invalid-json"
            ? WORKBOOK_SNAPSHOT_CONTENT_TYPE
            : "application/octet-stream",
          chunkSize,
        );

        await expect(connectPromise).rejects.toThrow(/.+/);
        expect(engine.getSyncState()).toBe("local-only");
      },
    });
  });
});
