import type {
  AckFrame,
  CursorWatermarkFrame,
  ErrorFrame,
  HeartbeatFrame,
  HelloFrame,
  ProtocolFrame,
} from "@bilig/binary-protocol";
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames } from "@bilig/binary-protocol";
import {
  XLSX_CONTENT_TYPE,
  type AgentFrame,
  type AgentResponse,
  type LoadWorkbookFileRequest,
} from "@bilig/agent-api";
import { importXlsx } from "@bilig/excel-import";
import {
  type InMemoryDocumentPersistence,
  createInMemoryDocumentPersistence,
} from "@bilig/storage-server";
import type { WorkbookSnapshot } from "@bilig/protocol";
import type { WorksheetExecutor } from "./worksheet-executor.js";

interface SnapshotAssembly {
  documentId: string;
  snapshotId: string;
  cursor: number;
  contentType: string;
  chunkCount: number;
  chunks: Array<Uint8Array | undefined>;
}

interface BrowserSubscriber {
  id: string;
  send(frame: ProtocolFrame): void;
}

export interface AgentFrameContext {
  serverUrl?: string;
  browserAppBaseUrl?: string;
}

export interface DocumentStateSummary {
  documentId: string;
  cursor: number;
  owner: string | null;
  sessions: string[];
  latestSnapshotCursor: number | null;
}

export interface DocumentSessionManagerOptions {
  browserAppBaseUrl?: string;
  publicServerUrl?: string;
  maxImportBytes?: number;
}

const DEFAULT_IMPORT_MAX_BYTES = 10 * 1024 * 1024;
const snapshotEncoder = new TextEncoder();

function normalizeBaseUrl(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

function buildBrowserUrl(
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

function decodeBase64(bytesBase64: string): Uint8Array {
  const normalized = bytesBase64.trim();
  if (normalized.length === 0 || normalized.length % 4 !== 0) {
    throw new Error("Workbook upload bytesBase64 must be a non-empty base64 string");
  }
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(normalized)) {
    throw new Error("Workbook upload bytesBase64 contains invalid base64 characters");
  }
  return new Uint8Array(Buffer.from(normalized, "base64"));
}

function createImportedDocumentId(): string {
  const random =
    typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : Math.random().toString(36).slice(2);
  return `xlsx:${random}`;
}

function normalizeSessionId(documentId: string, replicaId: string): string {
  return `${documentId}:${replicaId}`;
}

export class DocumentSessionManager {
  private readonly snapshotAssemblies = new Map<string, SnapshotAssembly>();
  private readonly browserSubscribers = new Map<string, Map<string, BrowserSubscriber>>();

  constructor(
    readonly persistence: InMemoryDocumentPersistence = createInMemoryDocumentPersistence(),
    private readonly ownerId = "bilig-app",
    private readonly worksheetExecutor: WorksheetExecutor | null = null,
    private readonly options: DocumentSessionManagerOptions = {},
  ) {}

  async handleSyncFrame(frame: ProtocolFrame): Promise<ProtocolFrame> {
    switch (frame.kind) {
      case "hello":
        await this.persistence.presence.join(frame.documentId, frame.sessionId);
        await this.persistence.ownership.claim(frame.documentId, this.ownerId, Date.now() + 60_000);
        return {
          kind: "cursorWatermark",
          documentId: frame.documentId,
          cursor: await this.persistence.batches.latestCursor(frame.documentId),
          compactedCursor: (await this.persistence.snapshots.latest(frame.documentId))?.cursor ?? 0,
        } satisfies CursorWatermarkFrame;

      case "appendBatch": {
        const stored = await this.persistence.batches.append(frame.documentId, frame.batch);
        this.broadcast(frame.documentId, {
          kind: "appendBatch",
          documentId: frame.documentId,
          cursor: stored.cursor,
          batch: frame.batch,
        });
        return {
          kind: "ack",
          documentId: frame.documentId,
          batchId: frame.batch.id,
          cursor: stored.cursor,
          acceptedAtUnixMs: stored.receivedAtUnixMs,
        } satisfies AckFrame;
      }

      case "snapshotChunk":
        await this.acceptSnapshotChunk(frame);
        return {
          kind: "ack",
          documentId: frame.documentId,
          batchId: frame.snapshotId,
          cursor: frame.cursor,
          acceptedAtUnixMs: Date.now(),
        } satisfies AckFrame;

      case "heartbeat":
        return {
          kind: "heartbeat",
          documentId: frame.documentId,
          cursor: await this.persistence.batches.latestCursor(frame.documentId),
          sentAtUnixMs: Date.now(),
        } satisfies HeartbeatFrame;

      case "cursorWatermark":
      case "ack":
        return frame;

      case "error":
        return frame;

      default:
        return {
          kind: "error",
          documentId: "unknown",
          code: "UNSUPPORTED_FRAME",
          message: `Unsupported sync frame ${(frame as ProtocolFrame).kind}`,
          retryable: false,
        } satisfies ErrorFrame;
    }
  }

  async handleAgentFrame(frame: AgentFrame, context: AgentFrameContext = {}): Promise<AgentFrame> {
    if (frame.kind !== "request") {
      return {
        kind: "response",
        response: {
          kind: "error",
          id: "unknown",
          code: "INVALID_AGENT_FRAME",
          message: "Sync server accepts only agent requests on the remote API ingress",
          retryable: false,
        },
      };
    }

    const request = frame.request;
    let response: AgentResponse;
    try {
      if (request.kind === "loadWorkbookFile") {
        response = await this.loadWorkbookFile(request, context);
        return {
          kind: "response",
          response,
        };
      }

      if (this.worksheetExecutor) {
        switch (request.kind) {
          case "openWorkbookSession":
            await this.persistence.presence.join(
              request.documentId,
              `${request.documentId}:${request.replicaId}`,
            );
            return this.worksheetExecutor.execute(frame);
          case "closeWorkbookSession":
            await this.persistence.presence.leave(
              request.sessionId.split(":")[0] ?? request.sessionId,
              request.sessionId,
            );
            return this.worksheetExecutor.execute(frame);
          case "readRange":
          case "writeRange":
          case "setRangeFormulas":
          case "setRangeStyle":
          case "clearRangeStyle":
          case "setRangeNumberFormat":
          case "clearRangeNumberFormat":
          case "clearRange":
          case "fillRange":
          case "copyRange":
          case "pasteRange":
          case "getDependents":
          case "getPrecedents":
          case "subscribeRange":
          case "unsubscribe":
          case "exportSnapshot":
          case "importSnapshot":
          case "createPivotTable":
            return this.worksheetExecutor.execute(frame);
          case "getMetrics":
            break;
        }
      }

      switch (request.kind) {
        case "openWorkbookSession":
          await this.persistence.presence.join(request.documentId, request.id);
          response = {
            kind: "ok",
            id: request.id,
            sessionId: `${request.documentId}:${request.replicaId}`,
          };
          break;
        case "closeWorkbookSession":
          await this.persistence.presence.leave(
            request.sessionId.split(":")[0] ?? request.sessionId,
            request.sessionId,
          );
          response = {
            kind: "ok",
            id: request.id,
          };
          break;
        case "getMetrics":
          response = {
            kind: "metrics",
            id: request.id,
            value: {
              service: "bilig-app",
              documentSessions: (
                await this.persistence.presence.sessions(
                  request.sessionId.split(":")[0] ?? request.sessionId,
                )
              ).length,
            },
          };
          break;
        case "readRange":
        case "writeRange":
        case "setRangeFormulas":
        case "setRangeStyle":
        case "clearRangeStyle":
        case "setRangeNumberFormat":
        case "clearRangeNumberFormat":
        case "clearRange":
        case "fillRange":
        case "copyRange":
        case "pasteRange":
        case "getDependents":
        case "getPrecedents":
        case "subscribeRange":
        case "unsubscribe":
        case "exportSnapshot":
        case "importSnapshot":
        case "createPivotTable":
          response = {
            kind: "error",
            id: request.id,
            code: "NOT_IMPLEMENTED",
            message: `${request.kind} is reserved in the remote API contract but not wired to a live worksheet host yet`,
            retryable: false,
          };
          break;
      }
    } catch (error) {
      response = {
        kind: "error",
        id: request.id,
        code: "SYNC_SERVER_FAILURE",
        message: error instanceof Error ? error.message : String(error),
        retryable: false,
      };
    }

    return {
      kind: "response",
      response,
    };
  }

  attachBrowser(
    documentId: string,
    subscriberId: string,
    send: (frame: ProtocolFrame) => void,
  ): () => void {
    const subscribers =
      this.browserSubscribers.get(documentId) ?? new Map<string, BrowserSubscriber>();
    subscribers.set(subscriberId, { id: subscriberId, send });
    this.browserSubscribers.set(documentId, subscribers);
    return () => {
      const next = this.browserSubscribers.get(documentId);
      next?.delete(subscriberId);
      if (next && next.size === 0) {
        this.browserSubscribers.delete(documentId);
      }
    };
  }

  async openBrowserSession(frame: HelloFrame): Promise<ProtocolFrame[]> {
    await this.persistence.presence.join(frame.documentId, frame.sessionId);
    await this.persistence.ownership.claim(frame.documentId, this.ownerId, Date.now() + 60_000);
    const latestSnapshot = await this.persistence.snapshots.latest(frame.documentId);
    const snapshotFrames =
      latestSnapshot && frame.lastServerCursor < latestSnapshot.cursor
        ? createSnapshotChunkFrames({
            documentId: frame.documentId,
            snapshotId: latestSnapshot.snapshotId,
            cursor: latestSnapshot.cursor,
            contentType: latestSnapshot.contentType,
            bytes: latestSnapshot.bytes,
          })
        : [];
    const cursorFloor = Math.max(frame.lastServerCursor, latestSnapshot?.cursor ?? 0);
    const missed = await this.persistence.batches.listAfter(frame.documentId, cursorFloor);
    return [
      ...snapshotFrames,
      ...missed.map((entry) => ({
        kind: "appendBatch" as const,
        documentId: entry.documentId,
        cursor: entry.cursor,
        batch: entry.batch,
      })),
      {
        kind: "cursorWatermark",
        documentId: frame.documentId,
        cursor: await this.persistence.batches.latestCursor(frame.documentId),
        compactedCursor: latestSnapshot?.cursor ?? 0,
      } satisfies CursorWatermarkFrame,
    ];
  }

  async getDocumentState(documentId: string): Promise<DocumentStateSummary> {
    const latestSnapshot = await this.persistence.snapshots.latest(documentId);
    return {
      documentId,
      cursor: await this.persistence.batches.latestCursor(documentId),
      owner: await this.persistence.ownership.owner(documentId),
      sessions: await this.persistence.presence.sessions(documentId),
      latestSnapshotCursor: latestSnapshot?.cursor ?? null,
    };
  }

  private async loadWorkbookFile(
    request: LoadWorkbookFileRequest,
    context: AgentFrameContext,
  ): Promise<AgentResponse> {
    if (request.contentType !== XLSX_CONTENT_TYPE) {
      throw new Error("Unsupported workbook upload content type");
    }
    if (request.openMode === "replace" && !request.documentId) {
      throw new Error("Workbook replace uploads require documentId");
    }

    const bytes = decodeBase64(request.bytesBase64);
    const maxImportBytes = this.options.maxImportBytes ?? DEFAULT_IMPORT_MAX_BYTES;
    if (bytes.byteLength > maxImportBytes) {
      throw new Error(`Workbook upload exceeds ${maxImportBytes} bytes`);
    }

    const imported = importXlsx(bytes, request.fileName);
    const documentId = request.documentId ?? createImportedDocumentId();
    const sessionId = normalizeSessionId(documentId, request.replicaId);

    await this.persistence.presence.join(documentId, sessionId);
    await this.publishImportedSnapshot(documentId, imported.snapshot);

    const serverUrl = normalizeBaseUrl(
      context.serverUrl ?? this.options.publicServerUrl ?? "http://127.0.0.1:4321",
    );
    return {
      kind: "workbookLoaded",
      id: request.id,
      documentId,
      sessionId,
      workbookName: imported.workbookName,
      sheetNames: imported.sheetNames,
      serverUrl,
      ...(() => {
        const browserUrl = buildBrowserUrl(
          context.browserAppBaseUrl ?? this.options.browserAppBaseUrl,
          serverUrl,
          documentId,
        );
        return browserUrl ? { browserUrl } : {};
      })(),
      warnings: imported.warnings,
    };
  }

  private async publishImportedSnapshot(
    documentId: string,
    snapshot: WorkbookSnapshot,
  ): Promise<void> {
    const cursor = (await this.persistence.batches.latestCursor(documentId)) + 1;
    const snapshotId = `${documentId}:snapshot:${Date.now()}`;
    const bytes = snapshotEncoder.encode(JSON.stringify(snapshot));

    await this.persistence.batches.reset(documentId, cursor);
    await this.persistence.snapshots.put({
      documentId,
      snapshotId,
      cursor,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes,
      createdAtUnixMs: Date.now(),
    });

    const frames = createSnapshotChunkFrames({
      documentId,
      snapshotId,
      cursor,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes,
    });
    frames.forEach((frame) => this.broadcast(documentId, frame));
    this.broadcast(documentId, {
      kind: "cursorWatermark",
      documentId,
      cursor,
      compactedCursor: cursor,
    });
  }

  private async acceptSnapshotChunk(
    frame: Extract<ProtocolFrame, { kind: "snapshotChunk" }>,
  ): Promise<void> {
    const assembly = this.snapshotAssemblies.get(frame.snapshotId) ?? {
      documentId: frame.documentId,
      snapshotId: frame.snapshotId,
      cursor: frame.cursor,
      contentType: frame.contentType,
      chunkCount: frame.chunkCount,
      chunks: Array.from<Uint8Array | undefined>({ length: frame.chunkCount }),
    };
    assembly.chunks[frame.chunkIndex] = frame.bytes;
    this.snapshotAssemblies.set(frame.snapshotId, assembly);

    if (assembly.chunks.every((chunk): chunk is Uint8Array => chunk !== undefined)) {
      const totalLength = assembly.chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
      const bytes = new Uint8Array(totalLength);
      let offset = 0;
      assembly.chunks.forEach((chunk) => {
        bytes.set(chunk, offset);
        offset += chunk.byteLength;
      });
      await this.persistence.snapshots.put({
        documentId: frame.documentId,
        snapshotId: frame.snapshotId,
        cursor: frame.cursor,
        contentType: frame.contentType,
        bytes,
        createdAtUnixMs: Date.now(),
      });
      this.snapshotAssemblies.delete(frame.snapshotId);
    }
  }

  private broadcast(documentId: string, frame: ProtocolFrame): void {
    this.browserSubscribers.get(documentId)?.forEach((subscriber) => subscriber.send(frame));
  }
}
