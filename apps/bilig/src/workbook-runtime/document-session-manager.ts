import type { ErrorFrame, HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import type { AgentFrame, AgentResponse, LoadWorkbookFileRequest } from "@bilig/agent-api";
import {
  type InMemoryDocumentPersistence,
  createInMemoryDocumentPersistence,
} from "@bilig/storage-server";
import type { WorksheetExecutor } from "./worksheet-executor.js";
import {
  type AgentFrameContext,
  routeAgentFrame,
  worksheetHostUnavailableResponse,
} from "./agent-routing.js";
import { createBrowserHelloReplay } from "./browser-sync-replay.js";
import {
  closePresenceBackedWorkbookSession,
  countPresenceBackedWorkbookSessions,
  joinOwnedBrowserSession,
  openPresenceBackedWorkbookSession,
} from "./document-presence-session-store.js";
import {
  acceptPersistedSnapshotChunk,
  publishPersistedSnapshot,
} from "./document-snapshot-store.js";
import {
  attachBrowserSubscriber,
  broadcastToBrowsers,
  type BrowserSubscriberRegistry,
  type SnapshotAssemblyRegistry,
} from "./session-shared.js";
import {
  createAckFrame,
  createAppendBatchFrame,
  createCursorWatermarkFrame,
  createHeartbeatFrame,
} from "./sync-frame-shared.js";
import {
  createCloseWorkbookSessionResponse,
  createOpenWorkbookSessionResponse,
  loadWorkbookIntoRuntime,
} from "./workbook-session-shared.js";

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
        await joinOwnedBrowserSession(
          this.persistence,
          this.ownerId,
          frame.documentId,
          frame.sessionId,
        );
        return createCursorWatermarkFrame(
          frame.documentId,
          await this.persistence.batches.latestCursor(frame.documentId),
          (await this.persistence.snapshots.latest(frame.documentId))?.cursor ?? 0,
        );

      case "appendBatch": {
        const stored = await this.persistence.batches.append(frame.documentId, frame.batch);
        this.broadcast(
          frame.documentId,
          createAppendBatchFrame(frame.documentId, stored.cursor, frame.batch),
        );
        return createAckFrame(
          frame.documentId,
          frame.batch.id,
          stored.cursor,
          stored.receivedAtUnixMs,
        );
      }

      case "snapshotChunk":
        await this.acceptSnapshotChunk(frame);
        return createAckFrame(frame.documentId, frame.snapshotId, frame.cursor);

      case "heartbeat":
        return createHeartbeatFrame(
          frame.documentId,
          await this.persistence.batches.latestCursor(frame.documentId),
        );

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
        const sessionId = await openPresenceBackedWorkbookSession(
          this.persistence,
          request.documentId,
          request.replicaId,
        );
        if (this.worksheetExecutor) {
          return this.worksheetExecutor.execute({
            kind: "request",
            request,
          });
        }
        return createOpenWorkbookSessionResponse(request.id, sessionId);
      },
      closeWorkbookSession: async (request) => {
        await closePresenceBackedWorkbookSession(this.persistence, request.sessionId);
        if (this.worksheetExecutor) {
          return this.worksheetExecutor.execute({
            kind: "request",
            request,
          });
        }
        return createCloseWorkbookSessionResponse(request.id);
      },
      getMetrics: async (request) => ({
        kind: "metrics",
        id: request.id,
        value: {
          service: "bilig-app",
          documentSessions: await countPresenceBackedWorkbookSessions(
            this.persistence,
            request.sessionId,
          ),
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
    await joinOwnedBrowserSession(
      this.persistence,
      this.ownerId,
      frame.documentId,
      frame.sessionId,
    );
    return createBrowserHelloReplay({
      documentId: frame.documentId,
      lastServerCursor: frame.lastServerCursor,
      latestCursor: this.persistence.batches.latestCursor(frame.documentId),
      latestSnapshot: this.persistence.snapshots.latest(frame.documentId),
      listMissedFrames: async (cursorFloor) => {
        const missed = await this.persistence.batches.listAfter(frame.documentId, cursorFloor);
        return missed.map((entry) =>
          createAppendBatchFrame(entry.documentId, entry.cursor, entry.batch),
        );
      },
    });
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
    return loadWorkbookIntoRuntime(request, context, {
      ...(this.options.maxImportBytes !== undefined
        ? { maxImportBytes: this.options.maxImportBytes }
        : {}),
      ...(this.options.publicServerUrl ? { publicServerUrl: this.options.publicServerUrl } : {}),
      ...(this.options.browserAppBaseUrl
        ? { browserAppBaseUrl: this.options.browserAppBaseUrl }
        : {}),
      registerPreparedSession: async (prepared) => {
        await joinOwnedBrowserSession(
          this.persistence,
          this.ownerId,
          prepared.documentId,
          prepared.sessionId,
        );
      },
      publishImportedSnapshot: async (documentId, snapshot) => {
        await publishPersistedSnapshot(
          this.persistence,
          documentId,
          snapshot,
          this.broadcast.bind(this),
        );
      },
    });
  }

  private async acceptSnapshotChunk(
    frame: Extract<ProtocolFrame, { kind: "snapshotChunk" }>,
  ): Promise<void> {
    await acceptPersistedSnapshotChunk(this.persistence, this.snapshotAssemblies, frame);
  }

  private broadcast(documentId: string, frame: ProtocolFrame): void {
    broadcastToBrowsers(this.browserSubscribers, documentId, frame);
  }
}
