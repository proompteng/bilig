import type { ProtocolFrame } from "@bilig/binary-protocol";
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames } from "@bilig/binary-protocol";
import type { WorkbookSnapshot } from "@bilig/protocol";

const snapshotEncoder = new TextEncoder();

export interface BrowserSubscriber {
  id: string;
  send(frame: ProtocolFrame): void;
}

export type BrowserSubscriberRegistry = Map<string, Map<string, BrowserSubscriber>>;

export interface SnapshotPublication {
  snapshotId: string;
  contentType: typeof WORKBOOK_SNAPSHOT_CONTENT_TYPE;
  bytes: Uint8Array;
  frames: ReturnType<typeof createSnapshotChunkFrames>;
}

export function attachBrowserSubscriber(
  registry: BrowserSubscriberRegistry,
  documentId: string,
  subscriberId: string,
  send: (frame: ProtocolFrame) => void,
): () => void {
  const subscribers = registry.get(documentId) ?? new Map<string, BrowserSubscriber>();
  subscribers.set(subscriberId, { id: subscriberId, send });
  registry.set(documentId, subscribers);
  return () => {
    const next = registry.get(documentId);
    next?.delete(subscriberId);
    if (next && next.size === 0) {
      registry.delete(documentId);
    }
  };
}

export function broadcastToBrowsers(
  registry: BrowserSubscriberRegistry,
  documentId: string,
  frame: ProtocolFrame,
): void {
  registry.get(documentId)?.forEach((subscriber) => subscriber.send(frame));
}

export function listBrowserSubscriberIds(
  registry: BrowserSubscriberRegistry,
  documentId: string,
): string[] {
  return [...(registry.get(documentId)?.keys() ?? [])];
}

export function normalizeBaseUrl(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

export function buildBrowserUrl(
  browserAppBaseUrl: string | undefined,
  serverUrl: string,
  documentId: string,
): string | undefined {
  if (!browserAppBaseUrl) {
    return undefined;
  }
  const url = new URL(normalizeBaseUrl(browserAppBaseUrl));
  url.searchParams.set("document", documentId);
  url.searchParams.set("server", serverUrl);
  return url.toString();
}

export function decodeWorkbookBase64(bytesBase64: string): Uint8Array {
  const normalized = bytesBase64.trim();
  if (normalized.length === 0 || normalized.length % 4 !== 0) {
    throw new Error("Workbook upload bytesBase64 must be a non-empty base64 string");
  }
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(normalized)) {
    throw new Error("Workbook upload bytesBase64 contains invalid base64 characters");
  }
  return new Uint8Array(Buffer.from(normalized, "base64"));
}

export function createImportedDocumentId(): string {
  const random =
    typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : Math.random().toString(36).slice(2);
  return `xlsx:${random}`;
}

export function createSnapshotPublication(
  documentId: string,
  cursor: number,
  snapshot: WorkbookSnapshot,
): SnapshotPublication {
  const snapshotId = `${documentId}:snapshot:${Date.now()}`;
  const bytes = snapshotEncoder.encode(JSON.stringify(snapshot));
  return {
    snapshotId,
    contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
    bytes,
    frames: createSnapshotChunkFrames({
      documentId,
      snapshotId,
      cursor,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes,
    }),
  };
}
