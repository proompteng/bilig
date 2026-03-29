import type { EngineSyncClient } from "@bilig/core";
import type { WorkbookSnapshot } from "@bilig/protocol";
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, decodeFrame, encodeFrame } from "@bilig/binary-protocol";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  return (
    isRecord(value) &&
    value["version"] === 1 &&
    isRecord(value["workbook"]) &&
    typeof value["workbook"]["name"] === "string" &&
    Array.isArray(value["sheets"]) &&
    value["sheets"].every((sheet) => {
      return (
        isRecord(sheet) &&
        typeof sheet["name"] === "string" &&
        typeof sheet["order"] === "number" &&
        Array.isArray(sheet["cells"]) &&
        sheet["cells"].every((cell) => {
          return (
            isRecord(cell) &&
            typeof cell["address"] === "string" &&
            (cell["value"] === undefined ||
              cell["value"] === null ||
              typeof cell["value"] === "string" ||
              typeof cell["value"] === "number" ||
              typeof cell["value"] === "boolean") &&
            (cell["formula"] === undefined || typeof cell["formula"] === "string") &&
            (cell["format"] === undefined || typeof cell["format"] === "string")
          );
        })
      );
    })
  );
}

function parseWorkbookSnapshot(bytes: Uint8Array): WorkbookSnapshot {
  const parsed: unknown = JSON.parse(new TextDecoder().decode(bytes));
  if (!isWorkbookSnapshot(parsed)) {
    throw new Error("Invalid workbook snapshot payload");
  }
  return parsed;
}

function toBufferSource(bytes: Uint8Array): ArrayBuffer {
  return bytes.slice().buffer;
}

export interface BrowserWebSocketLike {
  binaryType: BinaryType;
  readyState: number;
  addEventListener(type: "open" | "message" | "error" | "close", listener: EventListener): void;
  send(data: BufferSource): void;
  close(): void;
}

export interface WebSocketSyncClientOptions {
  documentId: string;
  replicaId: string;
  baseUrl: string;
  initialServerCursor?: number;
  createSocket?: (url: string) => BrowserWebSocketLike;
}

function createBrowserWebSocket(url: string): BrowserWebSocketLike {
  const socket = new WebSocket(url);
  return {
    get binaryType() {
      return socket.binaryType;
    },
    set binaryType(value: BinaryType) {
      socket.binaryType = value;
    },
    get readyState() {
      return socket.readyState;
    },
    addEventListener(type, listener) {
      socket.addEventListener(type, listener);
    },
    send(data) {
      socket.send(data);
    },
    close() {
      socket.close();
    },
  };
}

function toWebSocketUrl(baseUrl: string, documentId: string): string {
  const url = new URL(`/v2/documents/${encodeURIComponent(documentId)}/ws`, baseUrl);
  if (url.protocol === "http:") {
    url.protocol = "ws:";
  } else if (url.protocol === "https:") {
    url.protocol = "wss:";
  }
  return url.toString();
}

async function toBytes(data: unknown): Promise<Uint8Array> {
  if (data instanceof Uint8Array) {
    return data;
  }
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }
  if (ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  if (typeof Blob !== "undefined" && data instanceof Blob) {
    return new Uint8Array(await data.arrayBuffer());
  }
  throw new Error("Unsupported websocket message payload");
}

export function createWebSocketSyncClient(options: WebSocketSyncClientOptions): EngineSyncClient {
  return {
    async connect(handlers) {
      const createSocket: (url: string) => BrowserWebSocketLike =
        options.createSocket ??
        ((url: string) => {
          if (typeof WebSocket === "undefined") {
            throw new Error("WebSocket is unavailable in this runtime");
          }
          return createBrowserWebSocket(url);
        });
      const socket = createSocket(toWebSocketUrl(options.baseUrl, options.documentId));
      let lastServerCursor = options.initialServerCursor ?? 0;
      let settled = false;
      const pendingSnapshots = new Map<
        string,
        {
          cursor: number;
          contentType: string;
          chunks: Array<Uint8Array | undefined>;
        }
      >();

      socket.binaryType = "arraybuffer";

      const connection = {
        send(batch: Parameters<Parameters<EngineSyncClient["connect"]>[0]["applyRemoteBatch"]>[0]) {
          if (socket.readyState !== 1) {
            throw new Error("WebSocket sync client is not connected");
          }
          socket.send(
            toBufferSource(
              encodeFrame({
                kind: "appendBatch",
                documentId: options.documentId,
                cursor: lastServerCursor,
                batch,
              }),
            ),
          );
        },
        disconnect() {
          socket.close();
          handlers.setState("local-only");
        },
      };

      const complete = <T>(callback: () => T, resolve: (value: T) => void) => {
        if (settled) {
          return;
        }
        settled = true;
        resolve(callback());
      };

      return new Promise<typeof connection>((resolve, reject) => {
        const fail = (message: string) => {
          handlers.setState("local-only");
          if (settled) {
            return;
          }
          settled = true;
          reject(new Error(message));
        };

        const handleFrame = async (payload: unknown) => {
          const frame = decodeFrame(await toBytes(payload));
          switch (frame.kind) {
            case "hello":
              break;
            case "appendBatch":
              lastServerCursor = Math.max(lastServerCursor, frame.cursor);
              handlers.applyRemoteBatch(frame.batch);
              handlers.setState("live");
              complete(() => connection, resolve);
              break;
            case "cursorWatermark":
              lastServerCursor = Math.max(lastServerCursor, frame.cursor);
              handlers.setState("live");
              complete(() => connection, resolve);
              break;
            case "ack":
              lastServerCursor = Math.max(lastServerCursor, frame.cursor);
              handlers.setState("live");
              complete(() => connection, resolve);
              break;
            case "heartbeat":
              lastServerCursor = Math.max(lastServerCursor, frame.cursor);
              handlers.setState("live");
              complete(() => connection, resolve);
              break;
            case "error":
              fail(frame.message);
              break;
            case "snapshotChunk":
              lastServerCursor = Math.max(lastServerCursor, frame.cursor);
              const assembly = pendingSnapshots.get(frame.snapshotId) ?? {
                cursor: frame.cursor,
                contentType: frame.contentType,
                chunks: Array.from<Uint8Array | undefined>({ length: frame.chunkCount }),
              };
              assembly.chunks[frame.chunkIndex] = frame.bytes;
              pendingSnapshots.set(frame.snapshotId, assembly);
              if (assembly.chunks.every((chunk): chunk is Uint8Array => chunk !== undefined)) {
                pendingSnapshots.delete(frame.snapshotId);
                if (assembly.contentType !== WORKBOOK_SNAPSHOT_CONTENT_TYPE) {
                  fail(`Unsupported snapshot content type ${assembly.contentType}`);
                  return;
                }
                const totalLength = assembly.chunks.reduce(
                  (sum, chunk) => sum + chunk.byteLength,
                  0,
                );
                const bytes = new Uint8Array(totalLength);
                let offset = 0;
                assembly.chunks.forEach((chunk) => {
                  bytes.set(chunk, offset);
                  offset += chunk.byteLength;
                });
                handlers.applyRemoteSnapshot?.(parseWorkbookSnapshot(bytes));
                handlers.setState("live");
                complete(() => connection, resolve);
              }
              break;
          }
        };

        socket.addEventListener("open", () => {
          handlers.setState("syncing");
          socket.send(
            toBufferSource(
              encodeFrame({
                kind: "hello",
                documentId: options.documentId,
                replicaId: options.replicaId,
                sessionId: `${options.documentId}:${options.replicaId}`,
                protocolVersion: 1,
                lastServerCursor,
                capabilities: ["browser-sync"],
              }),
            ),
          );
        });
        socket.addEventListener("message", (event) => {
          if (!(event instanceof MessageEvent)) {
            fail("WebSocket message event did not provide a MessageEvent payload");
            return;
          }
          void handleFrame(event.data).catch((error: unknown) => {
            fail(error instanceof Error ? error.message : String(error));
          });
        });
        socket.addEventListener("error", () => {
          fail("WebSocket sync connection failed");
        });
        socket.addEventListener("close", () => {
          handlers.setState("local-only");
          if (!settled) {
            fail("WebSocket sync connection closed before handshake");
          }
        });
      });
    },
  };
}
