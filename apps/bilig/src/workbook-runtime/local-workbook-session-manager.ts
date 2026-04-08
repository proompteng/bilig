import type { HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import type {
  AgentEvent,
  AgentFrame,
  AgentResponse,
  LoadWorkbookFileRequest,
} from "@bilig/agent-api";
import { SpreadsheetEngine } from "@bilig/core";
import type { UpstreamSyncRelay } from "../zero/sync-relay.js";
import type { AgentFrameContext } from "./agent-routing.js";
import { WorkbookBrowserSessionHost } from "./browser-session-host.js";
import {
  flushQueuedLocalAgentEvents,
  queueLocalAgentEvent,
  removeQueuedSubscriptionEvents,
} from "./local-agent-event-queue.js";
import {
  acceptLocalSnapshotChunk,
  getLocalCachedSnapshot,
  invalidateLocalSnapshotCache,
  maybeCompactLocalSession,
  publishLocalSnapshot,
  storeLocalCachedSnapshot,
  type LocalSnapshotSessionState,
  type StoredBatch,
  type StoredSnapshotPublication,
} from "./local-session-snapshot-store.js";
import {
  handleLocalWorksheetAgentRequest,
  type LocalWorksheetSessionState,
} from "./local-worksheet-agent-handler.js";
import {
  closeLocalAgentSession,
  getLocalSessionByAgentSessionId,
  type LocalAgentRangeSubscriptionState,
  type LocalAgentSessionState,
  openLocalAgentSession,
  removeLocalAgentSubscription,
} from "./local-agent-session-store.js";
import {
  createAckFrame,
  createAppendBatchFrame,
  createHeartbeatFrame,
} from "./sync-frame-shared.js";
import { createUnsupportedSyncFrame } from "./sync-frame-router.js";
import { createWorkbookLoadOptions, loadWorkbookIntoRuntime } from "./workbook-session-shared.js";
import { WorkbookSessionCore } from "./workbook-session-core.js";
import { WorkbookSyncSessionHost } from "./workbook-sync-session-host.js";

interface LocalWorkbookSession extends LocalWorksheetSessionState, LocalSnapshotSessionState {
  documentId: string;
  engine: SpreadsheetEngine;
  agentSessions: Map<string, LocalAgentSessionState>;
  agentSubscriptions: Map<string, LocalAgentRangeSubscriptionState>;
  eventBacklog: AgentEvent[];
  eventFlushScheduled: boolean;
  batches: StoredBatch[];
  latestSnapshot: StoredSnapshotPublication | null;
  upstreamRelay: UpstreamSyncRelay | null;
  unsubscribeBatches: () => void;
  unsubscribeEvents: () => void;
}

export interface LocalWorkbookSessionManagerOptions {
  createSyncRelay?: (documentId: string) => UpstreamSyncRelay | null;
  browserAppBaseUrl?: string;
  publicServerUrl?: string;
  maxImportBytes?: number;
}

export interface LocalDocumentStateSummary {
  documentId: string;
  cursor: number;
  browserSessions: string[];
  agentSessions: string[];
  lastBatchId: string | null;
}

export interface LocalSnapshotSummary {
  cursor: number;
  contentType: string;
  bytes: Uint8Array;
}

const MAX_BATCH_BACKLOG = 256;
const LARGE_RANGE_SUBSCRIPTION_THRESHOLD = 256;

export class LocalWorkbookSessionManager {
  private readonly sessions = new Map<string, LocalWorkbookSession>();
  private readonly agentEventListeners = new Set<(event: AgentEvent) => void>();
  private readonly agentSubscriptionOwners = new Map<string, string>();
  private readonly sessionCore: WorkbookSessionCore<ProtocolFrame[]>;

  constructor(private readonly options: LocalWorkbookSessionManagerOptions = {}) {
    const browserSessionHost = new WorkbookBrowserSessionHost({
      latestCursor: (documentId) => this.ensureSession(documentId).cursor,
      latestSnapshot: (documentId) => this.ensureSession(documentId).latestSnapshot,
      listMissedFrames: (documentId, cursorFloor) =>
        this.ensureSession(documentId)
          .batches.filter((entry) => entry.cursor > cursorFloor)
          .map((entry) => entry.frame),
    });
    const syncSessionHost = new WorkbookSyncSessionHost<ProtocolFrame[]>({
      browserSessionHost,
      hello: (helloFrame) => this.openBrowserSession(helloFrame),
      appendBatch: (appendFrame) => {
        const session = this.ensureSession(appendFrame.documentId);
        if (!session.engine.applyRemoteBatch(appendFrame.batch)) {
          return [createAckFrame(appendFrame.documentId, appendFrame.batch.id, session.cursor)];
        }
        session.cursor += 1;
        const committedFrame = createAppendBatchFrame(
          appendFrame.documentId,
          session.cursor,
          appendFrame.batch,
        );
        session.batches.push({ cursor: session.cursor, frame: committedFrame });
        session.replicaSnapshot = session.engine.exportReplicaSnapshot();
        this.broadcast(appendFrame.documentId, committedFrame);
        maybeCompactLocalSession(session, this.snapshotStoreContext());
        void this.relayUpstream(session, appendFrame.batch);
        return [createAckFrame(appendFrame.documentId, appendFrame.batch.id, session.cursor)];
      },
      snapshotChunk: (snapshotFrame) => {
        const session = this.ensureSession(snapshotFrame.documentId);
        acceptLocalSnapshotChunk(session, snapshotFrame, this.snapshotStoreContext());
        return [
          createAckFrame(snapshotFrame.documentId, snapshotFrame.snapshotId, snapshotFrame.cursor),
        ];
      },
      heartbeat: (heartbeatFrame) => {
        const session = this.ensureSession(heartbeatFrame.documentId);
        return [createHeartbeatFrame(heartbeatFrame.documentId, session.cursor)];
      },
      passthrough: (passthroughFrame) => [passthroughFrame],
      unsupported: (unsupportedFrame) => [
        createUnsupportedSyncFrame(
          unsupportedFrame.documentId,
          "UNSUPPORTED_SYNC_FRAME",
          unsupportedFrame.kind,
          "Unsupported frame",
        ),
      ],
    });
    this.sessionCore = new WorkbookSessionCore({
      syncSessionHost,
      invalidFrameMessage: "Local server accepts only agent requests",
      errorCode: "LOCAL_SERVER_FAILURE",
      loadWorkbookFile: (request, requestContext) => this.loadWorkbookFile(request, requestContext),
      openWorkbookSession: (request) => {
        const session = this.ensureSession(request.documentId);
        return openLocalAgentSession(session, request.replicaId);
      },
      closeWorkbookSession: (request) => {
        const session = this.getSessionByAgentSessionId(request.sessionId);
        closeLocalAgentSession(
          session,
          request.sessionId,
          (ownedSession, sessionId, subscriptionId) => {
            this.removeAgentSubscription(ownedSession, sessionId, subscriptionId);
          },
        );
      },
      getMetrics: (request) => {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        return { kind: "metrics", id: request.id, value: engine.getLastMetrics() };
      },
      handleWorksheetRequest: (_frame, request) =>
        handleLocalWorksheetAgentRequest(
          {
            largeRangeSubscriptionThreshold: LARGE_RANGE_SUBSCRIPTION_THRESHOLD,
            agentSubscriptionOwners: this.agentSubscriptionOwners,
            getSessionByAgentSessionId: (sessionId) => this.getSessionByAgentSessionId(sessionId),
            getCachedSnapshot: (session) => getLocalCachedSnapshot(session),
            importSnapshot: (session, snapshot) => {
              session.engine.importSnapshot(snapshot);
              session.replicaSnapshot = session.engine.exportReplicaSnapshot();
              storeLocalCachedSnapshot(session, snapshot);
            },
            removeAgentSubscription: (session, sessionId, subscriptionId) => {
              this.removeAgentSubscription(session, sessionId, subscriptionId);
            },
            queueAgentEvent: (documentId, event) => {
              this.queueAgentEvent(documentId, event);
            },
          },
          request,
        ),
    });
  }

  private ensureSession(documentId: string): LocalWorkbookSession {
    const existing = this.sessions.get(documentId);
    if (existing) {
      return existing;
    }

    const engine = new SpreadsheetEngine({
      workbookName: documentId,
      replicaId: `worksheet-host:${documentId}`,
    });

    const session: LocalWorkbookSession = {
      documentId,
      engine,
      agentSessions: new Map(),
      agentSubscriptions: new Map(),
      eventBacklog: [],
      eventFlushScheduled: false,
      batches: [],
      latestSnapshot: null,
      snapshotCache: null,
      snapshotDirty: true,
      cursor: 0,
      replicaSnapshot: null,
      upstreamRelay: this.options.createSyncRelay?.(documentId) ?? null,
      unsubscribeBatches: () => {},
      unsubscribeEvents: () => {},
      compactScheduled: false,
    };

    session.unsubscribeEvents = engine.subscribe(() => {
      invalidateLocalSnapshotCache(session);
    });

    session.unsubscribeBatches = engine.subscribeBatches((batch) => {
      session.cursor += 1;
      const frame = createAppendBatchFrame(documentId, session.cursor, batch);
      session.batches.push({ cursor: session.cursor, frame });
      session.replicaSnapshot = engine.exportReplicaSnapshot();
      this.broadcast(documentId, frame);
      maybeCompactLocalSession(session, this.snapshotStoreContext());
      void this.relayUpstream(session, batch);
    });

    if (engine.workbook.sheetsByName.size === 0) {
      engine.createSheet("Sheet1");
    }

    this.sessions.set(documentId, session);
    return session;
  }

  attachBrowser(
    documentId: string,
    subscriberId: string,
    send: (frame: ProtocolFrame) => void,
  ): () => void {
    this.ensureSession(documentId);
    return this.sessionCore.attachBrowser(documentId, subscriberId, send);
  }

  subscribeAgentEvents(listener: (event: AgentEvent) => void): () => void {
    this.agentEventListeners.add(listener);
    if (this.agentEventListeners.size === 1) {
      this.sessions.forEach((session) => this.flushQueuedAgentEvents(session.documentId));
    }
    return () => {
      this.agentEventListeners.delete(listener);
    };
  }

  async openBrowserSession(frame: HelloFrame): Promise<ProtocolFrame[]> {
    this.ensureSession(frame.documentId);
    return this.sessionCore.openBrowserSession(frame);
  }

  async handleSyncFrame(frame: ProtocolFrame): Promise<ProtocolFrame[]> {
    return this.sessionCore.handleSyncFrame(frame);
  }

  async handleAgentFrame(frame: AgentFrame, context: AgentFrameContext = {}): Promise<AgentFrame> {
    return this.sessionCore.handleAgentFrame(frame, context);
  }

  getDocumentState(documentId: string): LocalDocumentStateSummary {
    const session = this.ensureSession(documentId);
    return {
      documentId,
      cursor: session.cursor,
      browserSessions: this.sessionCore.listSubscriberIds(documentId),
      agentSessions: [...session.agentSessions.keys()],
      lastBatchId: session.batches.at(-1)?.frame.batch.id ?? null,
    };
  }

  getLatestSnapshot(documentId: string): LocalSnapshotSummary | null {
    const session = this.sessions.get(documentId);
    if (!session?.latestSnapshot) {
      return null;
    }
    return {
      cursor: session.latestSnapshot.cursor,
      contentType: session.latestSnapshot.contentType,
      bytes: session.latestSnapshot.bytes,
    };
  }

  private getSessionByAgentSessionId(sessionId: string): LocalWorkbookSession {
    return getLocalSessionByAgentSessionId(this.sessions, sessionId);
  }

  private async loadWorkbookFile(
    request: LoadWorkbookFileRequest,
    context: AgentFrameContext,
  ): Promise<AgentResponse> {
    return loadWorkbookIntoRuntime(
      request,
      context,
      createWorkbookLoadOptions(this.options, {
        registerPreparedSession: (prepared, loadRequest) => {
          const session = this.ensureSession(prepared.documentId);
          openLocalAgentSession(session, loadRequest.replicaId);
        },
        publishImportedSnapshot: (documentId, snapshot) => {
          const session = this.ensureSession(documentId);
          session.engine.importSnapshot(snapshot);
          session.replicaSnapshot = session.engine.exportReplicaSnapshot();
          publishLocalSnapshot(session, snapshot, this.broadcast.bind(this));
        },
      }),
    );
  }

  private removeAgentSubscription(
    session: LocalWorkbookSession,
    sessionId: string,
    subscriptionId: string,
  ): void {
    removeLocalAgentSubscription(
      session,
      sessionId,
      subscriptionId,
      this.agentSubscriptionOwners,
      (removedSubscriptionId) => {
        removeQueuedSubscriptionEvents(session, removedSubscriptionId);
      },
    );
  }

  private queueAgentEvent(documentId: string, event: AgentEvent): void {
    queueLocalAgentEvent(this.agentEventQueueContext(), documentId, event);
  }

  private flushQueuedAgentEvents(documentId: string): void {
    flushQueuedLocalAgentEvents(this.agentEventQueueContext(), documentId);
  }

  private broadcast(documentId: string, frame: ProtocolFrame): void {
    this.sessionCore.broadcast(documentId, frame);
  }

  private snapshotStoreContext() {
    return {
      broadcast: this.broadcast.bind(this),
      getSession: (documentId: string) => this.sessions.get(documentId),
      snapshotAssemblies: this.sessionCore.snapshotAssemblies,
      maxBatchBacklog: MAX_BATCH_BACKLOG,
    };
  }

  private agentEventQueueContext() {
    return {
      getSession: (documentId: string) => this.sessions.get(documentId),
      listeners: this.agentEventListeners,
    };
  }

  private async relayUpstream(
    session: LocalWorkbookSession,
    batch: Extract<ProtocolFrame, { kind: "appendBatch" }>["batch"],
  ): Promise<void> {
    if (!session.upstreamRelay) {
      return;
    }
    try {
      await session.upstreamRelay.send(batch);
    } catch (error) {
      console.error(`Failed to relay batch ${batch.id} for document ${session.documentId}:`, error);
    }
  }
}
