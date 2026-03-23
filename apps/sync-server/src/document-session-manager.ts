import type {
  AckFrame,
  CursorWatermarkFrame,
  ErrorFrame,
  HeartbeatFrame,
  ProtocolFrame
} from "@bilig/binary-protocol";
import type { AgentFrame, AgentResponse } from "@bilig/agent-api";
import {
  type InMemoryDocumentPersistence,
  createInMemoryDocumentPersistence
} from "@bilig/storage-server";
import type { WorksheetExecutor } from "./worksheet-executor.js";

interface SnapshotAssembly {
  documentId: string;
  snapshotId: string;
  cursor: number;
  contentType: string;
  chunkCount: number;
  chunks: Array<Uint8Array | undefined>;
}

export interface DocumentStateSummary {
  documentId: string;
  cursor: number;
  owner: string | null;
  sessions: string[];
  latestSnapshotCursor: number | null;
}

export class DocumentSessionManager {
  private readonly snapshotAssemblies = new Map<string, SnapshotAssembly>();

  constructor(
    readonly persistence: InMemoryDocumentPersistence = createInMemoryDocumentPersistence(),
    private readonly ownerId = "bilig-sync-server",
    private readonly worksheetExecutor: WorksheetExecutor | null = null
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
          compactedCursor: (await this.persistence.snapshots.latest(frame.documentId))?.cursor ?? 0
        } satisfies CursorWatermarkFrame;

      case "appendBatch": {
        const stored = await this.persistence.batches.append(frame.documentId, frame.batch);
        return {
          kind: "ack",
          documentId: frame.documentId,
          batchId: frame.batch.id,
          cursor: stored.cursor,
          acceptedAtUnixMs: stored.receivedAtUnixMs
        } satisfies AckFrame;
      }

      case "snapshotChunk":
        await this.acceptSnapshotChunk(frame);
        return {
          kind: "ack",
          documentId: frame.documentId,
          batchId: frame.snapshotId,
          cursor: frame.cursor,
          acceptedAtUnixMs: Date.now()
        } satisfies AckFrame;

      case "heartbeat":
        return {
          kind: "heartbeat",
          documentId: frame.documentId,
          cursor: await this.persistence.batches.latestCursor(frame.documentId),
          sentAtUnixMs: Date.now()
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
          retryable: false
        } satisfies ErrorFrame;
    }
  }

  async handleAgentFrame(frame: AgentFrame): Promise<AgentFrame> {
    if (frame.kind !== "request") {
      return {
        kind: "response",
        response: {
          kind: "error",
          id: "unknown",
          code: "INVALID_AGENT_FRAME",
          message: "Sync server accepts only agent requests on the remote API ingress",
          retryable: false
        }
      };
    }

    const request = frame.request;
    let response: AgentResponse;
    if (this.worksheetExecutor) {
      switch (request.kind) {
        case "openWorkbookSession":
          await this.persistence.presence.join(request.documentId, `${request.documentId}:${request.replicaId}`);
          return this.worksheetExecutor.execute(frame);
        case "closeWorkbookSession":
          await this.persistence.presence.leave(request.sessionId.split(":")[0] ?? request.sessionId, request.sessionId);
          return this.worksheetExecutor.execute(frame);
        case "readRange":
        case "writeRange":
        case "setRangeFormulas":
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
          sessionId: `${request.documentId}:${request.replicaId}`
        };
        break;
      case "closeWorkbookSession":
        await this.persistence.presence.leave(request.sessionId.split(":")[0] ?? request.sessionId, request.sessionId);
        response = {
          kind: "ok",
          id: request.id
        };
        break;
      case "getMetrics":
        response = {
          kind: "metrics",
          id: request.id,
          value: {
            service: "sync-server",
            documentSessions: (await this.persistence.presence.sessions(request.sessionId.split(":")[0] ?? request.sessionId)).length
          }
        };
        break;
      case "readRange":
      case "writeRange":
      case "setRangeFormulas":
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
          retryable: false
        };
        break;
    }

    return {
      kind: "response",
      response
    };
  }

  async getDocumentState(documentId: string): Promise<DocumentStateSummary> {
    const latestSnapshot = await this.persistence.snapshots.latest(documentId);
    return {
      documentId,
      cursor: await this.persistence.batches.latestCursor(documentId),
      owner: await this.persistence.ownership.owner(documentId),
      sessions: await this.persistence.presence.sessions(documentId),
      latestSnapshotCursor: latestSnapshot?.cursor ?? null
    };
  }

  private async acceptSnapshotChunk(
    frame: Extract<ProtocolFrame, { kind: "snapshotChunk" }>
  ): Promise<void> {
    const assembly = this.snapshotAssemblies.get(frame.snapshotId) ?? {
      documentId: frame.documentId,
      snapshotId: frame.snapshotId,
      cursor: frame.cursor,
      contentType: frame.contentType,
      chunkCount: frame.chunkCount,
      chunks: Array.from<Uint8Array | undefined>({ length: frame.chunkCount })
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
        createdAtUnixMs: Date.now()
      });
      this.snapshotAssemblies.delete(frame.snapshotId);
    }
  }
}
