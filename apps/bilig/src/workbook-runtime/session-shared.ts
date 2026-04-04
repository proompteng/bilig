import type { ProtocolFrame, SnapshotChunkFrame } from "@bilig/binary-protocol";
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames } from "@bilig/binary-protocol";
import type { WorkbookSnapshot } from "@bilig/protocol";

const snapshotEncoder = new TextEncoder();
const snapshotDecoder = new TextDecoder();

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

interface SnapshotAssembly {
  documentId: string;
  snapshotId: string;
  cursor: number;
  contentType: string;
  chunkCount: number;
  chunks: Array<Uint8Array | undefined>;
}

export interface CompletedSnapshotAssembly {
  documentId: string;
  snapshotId: string;
  cursor: number;
  contentType: string;
  bytes: Uint8Array;
}

export type SnapshotAssemblyRegistry = Map<string, SnapshotAssembly>;

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
  return createSnapshotPublicationFromBytes({
    documentId,
    snapshotId,
    cursor,
    contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
    bytes,
  });
}

export function createSnapshotPublicationFromBytes(
  snapshot: CompletedSnapshotAssembly,
): SnapshotPublication {
  return {
    snapshotId: snapshot.snapshotId,
    contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
    bytes: snapshot.bytes,
    frames: createSnapshotChunkFrames({
      documentId: snapshot.documentId,
      snapshotId: snapshot.snapshotId,
      cursor: snapshot.cursor,
      contentType: snapshot.contentType,
      bytes: snapshot.bytes,
    }),
  };
}

export function acceptSnapshotChunk(
  registry: SnapshotAssemblyRegistry,
  frame: SnapshotChunkFrame,
): CompletedSnapshotAssembly | null {
  const assembly = registry.get(frame.snapshotId) ?? {
    documentId: frame.documentId,
    snapshotId: frame.snapshotId,
    cursor: frame.cursor,
    contentType: frame.contentType,
    chunkCount: frame.chunkCount,
    chunks: Array.from<Uint8Array | undefined>({ length: frame.chunkCount }),
  };
  assembly.chunks[frame.chunkIndex] = frame.bytes;
  registry.set(frame.snapshotId, assembly);

  if (!assembly.chunks.every((chunk): chunk is Uint8Array => chunk !== undefined)) {
    return null;
  }

  const totalLength = assembly.chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const bytes = new Uint8Array(totalLength);
  let offset = 0;
  assembly.chunks.forEach((chunk) => {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  });
  registry.delete(frame.snapshotId);

  return {
    documentId: assembly.documentId,
    snapshotId: assembly.snapshotId,
    cursor: assembly.cursor,
    contentType: assembly.contentType,
    bytes,
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  if (!isRecord(value) || value["version"] !== 1) {
    return false;
  }
  const workbook = value["workbook"];
  if (!isRecord(workbook) || typeof workbook["name"] !== "string") {
    return false;
  }
  const sheets = value["sheets"];
  if (!Array.isArray(sheets)) {
    return false;
  }
  return sheets.every((sheet) => {
    if (
      !isRecord(sheet) ||
      typeof sheet["name"] !== "string" ||
      typeof sheet["order"] !== "number"
    ) {
      return false;
    }
    const cells = sheet["cells"];
    if (!Array.isArray(cells)) {
      return false;
    }
    return cells.every((cell) => isRecord(cell) && typeof cell["address"] === "string");
  });
}

export function decodeWorkbookSnapshotBytes(snapshot: CompletedSnapshotAssembly): WorkbookSnapshot {
  if (snapshot.contentType !== WORKBOOK_SNAPSHOT_CONTENT_TYPE) {
    throw new Error(`Unsupported snapshot content type: ${snapshot.contentType}`);
  }
  const decoded: unknown = JSON.parse(snapshotDecoder.decode(snapshot.bytes));
  if (!isWorkbookSnapshot(decoded)) {
    throw new Error("Workbook snapshot payload does not match the expected schema");
  }
  return decoded;
}
