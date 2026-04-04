import type { ProtocolFrame } from "@bilig/binary-protocol";
import type {
  AgentEvent,
  AgentFrame,
  AgentResponse,
  LoadWorkbookFileRequest,
} from "@bilig/agent-api";
import { shouldApplyBatch } from "@bilig/crdt";
import { SpreadsheetEngine } from "@bilig/core";
import type { UpstreamSyncRelay } from "../zero/sync-relay.js";
import { type AgentFrameContext, routeAgentFrame } from "./agent-routing.js";
import { createBrowserHelloReplay } from "./browser-sync-replay.js";
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
  attachBrowserSubscriber,
  broadcastToBrowsers,
  type BrowserSubscriberRegistry,
  listBrowserSubscriberIds,
  type SnapshotAssemblyRegistry,
} from "./session-shared.js";
import {
  createAckFrame,
  createAppendBatchFrame,
  createHeartbeatFrame,
} from "./sync-frame-shared.js";
import { createUnsupportedSyncFrame, routeWorkbookSyncFrame } from "./sync-frame-router.js";
import {
  createCloseWorkbookSessionResponse,
  createOpenWorkbookSessionResponse,
  loadWorkbookIntoRuntime,
} from "./workbook-session-shared.js";

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
  private readonly snapshotAssemblies: SnapshotAssemblyRegistry = new Map();
  private readonly browserSubscribers: BrowserSubscriberRegistry = new Map();
  private readonly agentEventListeners = new Set<(event: AgentEvent) => void>();
  private readonly agentSubscriptionOwners = new Map<string, string>();

  constructor(private readonly options: LocalWorkbookSessionManagerOptions = {}) {}

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
    return attachBrowserSubscriber(this.browserSubscribers, documentId, subscriberId, send);
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

  async handleSyncFrame(frame: ProtocolFrame): Promise<ProtocolFrame[]> {
    const session = this.ensureSession(frame.documentId);
    return routeWorkbookSyncFrame<ProtocolFrame[]>(frame, {
      hello: (helloFrame) =>
        createBrowserHelloReplay({
          documentId: helloFrame.documentId,
          lastServerCursor: helloFrame.lastServerCursor,
          latestCursor: session.cursor,
          latestSnapshot: session.latestSnapshot,
          listMissedFrames: (cursorFloor) =>
            session.batches
              .filter((entry) => entry.cursor > cursorFloor)
              .map((entry) => entry.frame),
        }),
      appendBatch: (appendFrame) => {
        if (!shouldApplyBatch(session.engine.replica, appendFrame.batch)) {
          return [createAckFrame(appendFrame.documentId, appendFrame.batch.id, session.cursor)];
        }
        session.engine.applyRemoteBatch(appendFrame.batch);
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
      heartbeat: (heartbeatFrame) => [
        createHeartbeatFrame(heartbeatFrame.documentId, session.cursor),
      ],
      snapshotChunk: (snapshotFrame) => {
        acceptLocalSnapshotChunk(session, snapshotFrame, this.snapshotStoreContext());
        return [
          createAckFrame(snapshotFrame.documentId, snapshotFrame.snapshotId, snapshotFrame.cursor),
        ];
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
  }

  async handleAgentFrame(frame: AgentFrame, context: AgentFrameContext = {}): Promise<AgentFrame> {
    return routeAgentFrame(frame, context, {
      invalidFrameMessage: "Local server accepts only agent requests",
      errorCode: "LOCAL_SERVER_FAILURE",
      loadWorkbookFile: (request, requestContext) => this.loadWorkbookFile(request, requestContext),
      openWorkbookSession: (request) => {
        const session = this.ensureSession(request.documentId);
        const sessionId = openLocalAgentSession(session, request.replicaId);
        return createOpenWorkbookSessionResponse(request.id, sessionId);
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
        return createCloseWorkbookSessionResponse(request.id);
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

  getDocumentState(documentId: string): LocalDocumentStateSummary {
    const session = this.ensureSession(documentId);
    return {
      documentId,
      cursor: session.cursor,
      browserSessions: listBrowserSubscriberIds(this.browserSubscribers, documentId),
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
    return loadWorkbookIntoRuntime(request, context, {
      ...(this.options.maxImportBytes !== undefined
        ? { maxImportBytes: this.options.maxImportBytes }
        : {}),
      ...(this.options.publicServerUrl ? { publicServerUrl: this.options.publicServerUrl } : {}),
      ...(this.options.browserAppBaseUrl
        ? { browserAppBaseUrl: this.options.browserAppBaseUrl }
        : {}),
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
    });
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
        session.eventBacklog = session.eventBacklog.filter((event) => {
          return event.kind !== "rangeChanged" || event.subscriptionId !== removedSubscriptionId;
        });
      },
    );
  }

  private queueAgentEvent(documentId: string, event: AgentEvent): void {
    const session = this.sessions.get(documentId);
    if (!session) {
      return;
    }
    session.eventBacklog.push(event);
    if (session.eventFlushScheduled) {
      return;
    }
    session.eventFlushScheduled = true;
    setImmediate(() => {
      session.eventFlushScheduled = false;
      this.flushQueuedAgentEvents(documentId);
    });
  }

  private flushQueuedAgentEvents(documentId: string): void {
    const session = this.sessions.get(documentId);
    if (!session || session.eventBacklog.length === 0 || this.agentEventListeners.size === 0) {
      return;
    }
    const pending = session.eventBacklog.splice(0);
    pending.forEach((event) => {
      this.agentEventListeners.forEach((listener) => listener(event));
    });
  }

  private broadcast(documentId: string, frame: ProtocolFrame): void {
    broadcastToBrowsers(this.browserSubscribers, documentId, frame);
  }

  private snapshotStoreContext() {
    return {
      broadcast: this.broadcast.bind(this),
      getSession: (documentId: string) => this.sessions.get(documentId),
      snapshotAssemblies: this.snapshotAssemblies,
      maxBatchBacklog: MAX_BATCH_BACKLOG,
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
