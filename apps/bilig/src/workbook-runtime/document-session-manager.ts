import type { HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import type { AgentFrame, AgentResponse, LoadWorkbookFileRequest } from "@bilig/agent-api";
import {
  type InMemoryDocumentPersistence,
  createInMemoryDocumentPersistence,
} from "@bilig/storage-server";
import type { WorksheetExecutor } from "./worksheet-executor.js";
import type { AgentFrameContext } from "./agent-routing.js";
import { openWorkbookBrowserSession } from "./browser-session-shared.js";
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
import { createUnsupportedSyncFrame, routeWorkbookSyncFrame } from "./sync-frame-router.js";
import {
  createWorkbookLoadOptions,
  handleWorkbookAgentFrame,
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
    return routeWorkbookSyncFrame<ProtocolFrame>(frame, {
      hello: async (helloFrame) => {
        await joinOwnedBrowserSession(
          this.persistence,
          this.ownerId,
          helloFrame.documentId,
          helloFrame.sessionId,
        );
        return createCursorWatermarkFrame(
          helloFrame.documentId,
          await this.persistence.batches.latestCursor(helloFrame.documentId),
          (await this.persistence.snapshots.latest(helloFrame.documentId))?.cursor ?? 0,
        );
      },
      appendBatch: async (appendFrame) => {
        const stored = await this.persistence.batches.append(
          appendFrame.documentId,
          appendFrame.batch,
        );
        this.broadcast(
          appendFrame.documentId,
          createAppendBatchFrame(appendFrame.documentId, stored.cursor, appendFrame.batch),
        );
        return createAckFrame(
          appendFrame.documentId,
          appendFrame.batch.id,
          stored.cursor,
          stored.receivedAtUnixMs,
        );
      },
      snapshotChunk: async (snapshotFrame) => {
        await this.acceptSnapshotChunk(snapshotFrame);
        return createAckFrame(
          snapshotFrame.documentId,
          snapshotFrame.snapshotId,
          snapshotFrame.cursor,
        );
      },
      heartbeat: async (heartbeatFrame) =>
        createHeartbeatFrame(
          heartbeatFrame.documentId,
          await this.persistence.batches.latestCursor(heartbeatFrame.documentId),
        ),
      passthrough: (passthroughFrame) => passthroughFrame,
      unsupported: (unsupportedFrame) =>
        createUnsupportedSyncFrame(
          "unknown",
          "UNSUPPORTED_FRAME",
          unsupportedFrame.kind,
          "Unsupported sync frame",
        ),
    });
  }

  async handleAgentFrame(frame: AgentFrame, context: AgentFrameContext = {}): Promise<AgentFrame> {
    return handleWorkbookAgentFrame(frame, context, {
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
        return sessionId;
      },
      closeWorkbookSession: async (request) => {
        await closePresenceBackedWorkbookSession(this.persistence, request.sessionId);
        if (this.worksheetExecutor) {
          return this.worksheetExecutor.execute({
            kind: "request",
            request,
          });
        }
        return undefined;
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
      ...(this.worksheetExecutor
        ? {
            handleWorksheetRequest: (requestFrame: Extract<AgentFrame, { kind: "request" }>) =>
              this.worksheetExecutor!.execute(requestFrame),
          }
        : {}),
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
    return openWorkbookBrowserSession(frame, {
      register: async (helloFrame) => {
        await joinOwnedBrowserSession(
          this.persistence,
          this.ownerId,
          helloFrame.documentId,
          helloFrame.sessionId,
        );
      },
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
    return loadWorkbookIntoRuntime(
      request,
      context,
      createWorkbookLoadOptions(this.options, {
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
      }),
    );
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
