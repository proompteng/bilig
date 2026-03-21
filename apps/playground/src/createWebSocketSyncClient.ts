import type { EngineSyncClient } from "@bilig/core";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";

export interface BrowserWebSocketLike {
  binaryType: string;
  readyState: number;
  onopen: (() => void) | null;
  onmessage: ((event: { data: unknown }) => void) | null;
  onerror: (() => void) | null;
  onclose: (() => void) | null;
  send(data: ArrayBufferLike | ArrayBufferView): void;
  close(): void;
}

export interface WebSocketSyncClientOptions {
  documentId: string;
  replicaId: string;
  baseUrl: string;
  createSocket?: (url: string) => BrowserWebSocketLike;
}

function toWebSocketUrl(baseUrl: string, documentId: string): string {
  const url = new URL(`/v1/documents/${encodeURIComponent(documentId)}/ws`, baseUrl);
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
      const createSocket = options.createSocket ?? ((url: string) => {
        if (typeof WebSocket === "undefined") {
          throw new Error("WebSocket is unavailable in this runtime");
        }
        return new WebSocket(url);
      });
      const socket = createSocket(toWebSocketUrl(options.baseUrl, options.documentId));
      let lastServerCursor = 0;
      let settled = false;

      socket.binaryType = "arraybuffer";

      const connection = {
        send(batch: Parameters<Parameters<EngineSyncClient["connect"]>[0]["applyRemoteBatch"]>[0]) {
          if (socket.readyState !== 1) {
            throw new Error("WebSocket sync client is not connected");
          }
          socket.send(encodeFrame({
            kind: "appendBatch",
            documentId: options.documentId,
            cursor: lastServerCursor,
            batch
          }));
        },
        disconnect() {
          socket.close();
          handlers.setState("local-only");
        }
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
              break;
          }
        };

        socket.onopen = () => {
          handlers.setState("syncing");
          socket.send(encodeFrame({
            kind: "hello",
            documentId: options.documentId,
            replicaId: options.replicaId,
            sessionId: `${options.documentId}:${options.replicaId}`,
            protocolVersion: 1,
            lastServerCursor,
            capabilities: ["browser-sync"]
          }));
        };
        socket.onmessage = (event: { data: unknown }) => {
          void handleFrame(event.data).catch((error: unknown) => {
            fail(error instanceof Error ? error.message : String(error));
          });
        };
        socket.onerror = () => {
          fail("WebSocket sync connection failed");
        };
        socket.onclose = () => {
          handlers.setState("local-only");
          if (!settled) {
            fail("WebSocket sync connection closed before handshake");
          }
        };
      });
    }
  };
}
