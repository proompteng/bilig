import type {
  AckFrame,
  CursorWatermarkFrame,
  ErrorFrame,
  HeartbeatFrame,
  HelloFrame,
  ProtocolFrame,
} from "@bilig/binary-protocol";
import { createSnapshotChunkFrames } from "@bilig/binary-protocol";
import type { AgentFrame, AgentResponse, LoadWorkbookFileRequest } from "@bilig/agent-api";
import {
  type InMemoryDocumentPersistence,
  createInMemoryDocumentPersistence,
} from "@bilig/storage-server";
import type { WorkbookSnapshot } from "@bilig/protocol";
import type { WorksheetExecutor } from "./worksheet-executor.js";
import {
  type AgentFrameContext,
  createWorkbookLoadedResponse,
  normalizeSessionId,
  prepareWorkbookLoad,
  routeAgentFrame,
  worksheetHostUnavailableResponse,
} from "./agent-routing.js";
import {
  acceptSnapshotChunk,
  attachBrowserSubscriber,
  broadcastToBrowsers,
  type BrowserSubscriberRegistry,
  createSnapshotPublication,
  type SnapshotAssemblyRegistry,
} from "./session-shared.js";

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

export class DocumentSessionManager {
  private readonly snapshotAssemblies: SnapshotAssemblyRegistry = new Map();
  private readonly browserSubscribers: BrowserSubscriberRegistry = new Map();

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
    return routeAgentFrame(frame, context, {
      invalidFrameMessage: "Sync server accepts only agent requests on the remote API ingress",
      errorCode: "SYNC_SERVER_FAILURE",
      loadWorkbookFile: (request, requestContext) => this.loadWorkbookFile(request, requestContext),
      openWorkbookSession: async (request) => {
        const sessionId = normalizeSessionId(request.documentId, request.replicaId);
        await this.persistence.presence.join(request.documentId, sessionId);
        if (this.worksheetExecutor) {
          return this.worksheetExecutor.execute({
            kind: "request",
            request,
          });
        }
        return {
          kind: "ok",
          id: request.id,
          sessionId,
        };
      },
      closeWorkbookSession: async (request) => {
        await this.persistence.presence.leave(
          request.sessionId.split(":")[0] ?? request.sessionId,
          request.sessionId,
        );
        if (this.worksheetExecutor) {
          return this.worksheetExecutor.execute({
            kind: "request",
            request,
          });
        }
        return {
          kind: "ok",
          id: request.id,
        };
      },
      getMetrics: async (request) => ({
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
      }),
      handleWorksheetRequest: async (requestFrame, request) => {
        if (!this.worksheetExecutor) {
          return worksheetHostUnavailableResponse(request);
        }
        return this.worksheetExecutor.execute(requestFrame);
      },
    });
  }

  attachBrowser(
    documentId: string,
    subscriberId: string,
    send: (frame: ProtocolFrame) => void,
  ): () => void {
    return attachBrowserSubscriber(this.browserSubscribers, documentId, subscriberId, send);
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
    const prepared = prepareWorkbookLoad(request, context, {
      ...(this.options.maxImportBytes !== undefined
        ? { maxImportBytes: this.options.maxImportBytes }
        : {}),
      ...(this.options.publicServerUrl ? { publicServerUrl: this.options.publicServerUrl } : {}),
      ...(this.options.browserAppBaseUrl
        ? { browserAppBaseUrl: this.options.browserAppBaseUrl }
        : {}),
    });
    await this.persistence.presence.join(prepared.documentId, prepared.sessionId);
    await this.publishImportedSnapshot(prepared.documentId, prepared.imported.snapshot);
    return createWorkbookLoadedResponse(request.id, prepared);
  }

  private async publishImportedSnapshot(
    documentId: string,
    snapshot: WorkbookSnapshot,
  ): Promise<void> {
    const cursor = (await this.persistence.batches.latestCursor(documentId)) + 1;
    const publication = createSnapshotPublication(documentId, cursor, snapshot);

    await this.persistence.batches.reset(documentId, cursor);
    await this.persistence.snapshots.put({
      documentId,
      snapshotId: publication.snapshotId,
      cursor,
      contentType: publication.contentType,
      bytes: publication.bytes,
      createdAtUnixMs: Date.now(),
    });
    publication.frames.forEach((frame) => this.broadcast(documentId, frame));
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
    const snapshot = acceptSnapshotChunk(this.snapshotAssemblies, frame);
    if (!snapshot) {
      return;
    }
    await this.persistence.snapshots.put({
      documentId: snapshot.documentId,
      snapshotId: snapshot.snapshotId,
      cursor: snapshot.cursor,
      contentType: snapshot.contentType,
      bytes: snapshot.bytes,
      createdAtUnixMs: Date.now(),
    });
  }

  private broadcast(documentId: string, frame: ProtocolFrame): void {
    broadcastToBrowsers(this.browserSubscribers, documentId, frame);
  }
}
