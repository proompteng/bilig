import {
  SpreadsheetEngine,
  type EngineReplicaSnapshot
} from "@bilig/core";
import type { CellRangeRef, CellValue } from "@bilig/protocol";
import {
  type AckFrame,
  type CursorWatermarkFrame,
  type ErrorFrame,
  type HeartbeatFrame,
  type ProtocolFrame
} from "@bilig/binary-protocol";
import type { AgentEvent, AgentFrame, AgentResponse } from "@bilig/agent-api";
import { shouldApplyBatch } from "@bilig/crdt";
import type { UpstreamSyncRelay } from "./sync-relay.js";

interface BrowserSubscriber {
  id: string;
  send(frame: ProtocolFrame): void;
}

interface AgentSession {
  sessionId: string;
  documentId: string;
  replicaId: string;
  subscriptionIds: Set<string>;
}

interface AgentRangeSubscription {
  subscriptionId: string;
  sessionId: string;
  range: CellRangeRef;
  changedAddresses: string[];
  unsubscribe: () => void;
}

interface StoredBatch {
  cursor: number;
  frame: Extract<ProtocolFrame, { kind: "appendBatch" }>;
}

interface LocalWorkbookSession {
  documentId: string;
  engine: SpreadsheetEngine;
  browserSubscribers: Map<string, BrowserSubscriber>;
  agentSessions: Map<string, AgentSession>;
  agentSubscriptions: Map<string, AgentRangeSubscription>;
  eventBacklog: AgentEvent[];
  eventFlushScheduled: boolean;
  batches: StoredBatch[];
  cursor: number;
  replicaSnapshot: EngineReplicaSnapshot | null;
  upstreamRelay: UpstreamSyncRelay | null;
  unsubscribeBatches: () => void;
}

export interface LocalWorkbookSessionManagerOptions {
  createSyncRelay?: (documentId: string) => UpstreamSyncRelay | null;
}

export interface LocalDocumentStateSummary {
  documentId: string;
  cursor: number;
  browserSessions: string[];
  agentSessions: string[];
  lastBatchId: string | null;
}

function normalizeSessionId(documentId: string, replicaId: string): string {
  return `${documentId}:${replicaId}`;
}

function iterateRange(range: CellRangeRef): string[] {
  const [startColPart, startRowPart] = splitAddress(range.startAddress);
  const [endColPart, endRowPart] = splitAddress(range.endAddress);
  const startCol = decodeColumn(startColPart);
  const endCol = decodeColumn(endColPart);
  const startRow = Number.parseInt(startRowPart, 10);
  const endRow = Number.parseInt(endRowPart, 10);
  const addresses: string[] = [];
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      addresses.push(`${encodeColumn(col)}${row}`);
    }
  }
  return addresses;
}

function splitAddress(address: string): [string, string] {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }
  return [match[1]!.toUpperCase(), match[2]!];
}

function decodeColumn(column: string): number {
  let value = 0;
  for (let index = 0; index < column.length; index += 1) {
    value = value * 26 + (column.charCodeAt(index) - 64);
  }
  return value;
}

function encodeColumn(value: number): string {
  let next = value;
  let output = "";
  while (next > 0) {
    const remainder = (next - 1) % 26;
    output = String.fromCharCode(65 + remainder) + output;
    next = Math.floor((next - 1) / 26);
  }
  return output;
}

export class LocalWorkbookSessionManager {
  private readonly sessions = new Map<string, LocalWorkbookSession>();
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
      replicaId: `local-server:${documentId}`
    });

    const session: LocalWorkbookSession = {
      documentId,
      engine,
      browserSubscribers: new Map(),
      agentSessions: new Map(),
      agentSubscriptions: new Map(),
      eventBacklog: [],
      eventFlushScheduled: false,
      batches: [],
      cursor: 0,
      replicaSnapshot: null,
      upstreamRelay: this.options.createSyncRelay?.(documentId) ?? null,
      unsubscribeBatches: () => {}
    };

    session.unsubscribeBatches = engine.subscribeBatches((batch) => {
      session.cursor += 1;
      const frame = {
        kind: "appendBatch",
        documentId,
        cursor: session.cursor,
        batch
      } satisfies Extract<ProtocolFrame, { kind: "appendBatch" }>;
      session.batches.push({ cursor: session.cursor, frame });
      session.replicaSnapshot = engine.exportReplicaSnapshot();
      this.broadcast(documentId, frame);
      void this.relayUpstream(session, batch);
    });

    this.sessions.set(documentId, session);
    return session;
  }

  attachBrowser(documentId: string, subscriberId: string, send: (frame: ProtocolFrame) => void): () => void {
    const session = this.ensureSession(documentId);
    session.browserSubscribers.set(subscriberId, { id: subscriberId, send });
    return () => {
      session.browserSubscribers.delete(subscriberId);
    };
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
    const documentId = frame.documentId;
    switch (frame.kind) {
      case "hello": {
        const missed = session.batches
          .filter((entry) => entry.cursor > frame.lastServerCursor)
          .map((entry) => entry.frame);
        return [
          ...missed,
          {
            kind: "cursorWatermark",
            documentId: frame.documentId,
            cursor: session.cursor,
            compactedCursor: Math.max(0, session.cursor - session.batches.length)
          } satisfies CursorWatermarkFrame
        ];
      }

      case "appendBatch": {
        if (!shouldApplyBatch(session.engine.replica, frame.batch)) {
          return [this.ack(frame.documentId, frame.batch.id, session.cursor)];
        }
        session.engine.applyRemoteBatch(frame.batch);
        session.cursor += 1;
        const committedFrame = {
          kind: "appendBatch",
          documentId: frame.documentId,
          cursor: session.cursor,
          batch: frame.batch
        } satisfies Extract<ProtocolFrame, { kind: "appendBatch" }>;
        session.batches.push({ cursor: session.cursor, frame: committedFrame });
        session.replicaSnapshot = session.engine.exportReplicaSnapshot();
        this.broadcast(frame.documentId, committedFrame);
        void this.relayUpstream(session, frame.batch);
        return [this.ack(frame.documentId, frame.batch.id, session.cursor)];
      }

      case "heartbeat":
        return [
          {
            kind: "heartbeat",
            documentId: frame.documentId,
            cursor: session.cursor,
            sentAtUnixMs: Date.now()
          } satisfies HeartbeatFrame
        ];

      case "cursorWatermark":
      case "ack":
        return [frame];

      case "snapshotChunk":
        return [
          {
            kind: "error",
            documentId: frame.documentId,
            code: "UNSUPPORTED_LOCAL_SNAPSHOT_UPLOAD",
            message: "Local server snapshot uploads are not wired yet",
            retryable: false
          } satisfies ErrorFrame
        ];

      case "error":
        return [frame];

      default:
        return [
          {
            kind: "error",
            documentId,
            code: "UNSUPPORTED_SYNC_FRAME",
            message: "Unsupported frame",
            retryable: false
          } satisfies ErrorFrame
        ];
    }
  }

  async handleAgentFrame(frame: AgentFrame): Promise<AgentFrame> {
    if (frame.kind !== "request") {
      return this.agentError("unknown", "INVALID_AGENT_FRAME", "Local server accepts only agent requests", false);
    }

    const request = frame.request;
    const requestId = request.id;
    let response: AgentResponse;
    try {
      switch (request.kind) {
        case "openWorkbookSession": {
          const session = this.ensureSession(request.documentId);
          const sessionId = normalizeSessionId(request.documentId, request.replicaId);
          session.agentSessions.set(sessionId, {
            sessionId,
            documentId: request.documentId,
            replicaId: request.replicaId,
            subscriptionIds: new Set()
          });
          response = { kind: "ok", id: request.id, sessionId };
          break;
        }

        case "closeWorkbookSession": {
          const session = this.getSessionByAgentSessionId(request.sessionId);
          const agentSession = session.agentSessions.get(request.sessionId);
          if (agentSession) {
            [...agentSession.subscriptionIds].forEach((subscriptionId) => {
              this.removeAgentSubscription(session, request.sessionId, subscriptionId);
            });
          }
          session.agentSessions.delete(request.sessionId);
          response = { kind: "ok", id: request.id };
          break;
        }

        case "readRange": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          response = {
            kind: "rangeValues",
            id: request.id,
            values: this.readRange(engine, request.range)
          };
          break;
        }

        case "writeRange": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          engine.setRangeValues(request.range, request.values);
          response = { kind: "ok", id: request.id };
          break;
        }

        case "setRangeFormulas": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          engine.setRangeFormulas(request.range, request.formulas);
          response = { kind: "ok", id: request.id };
          break;
        }

        case "clearRange": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          engine.clearRange(request.range);
          response = { kind: "ok", id: request.id };
          break;
        }

        case "fillRange": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          engine.fillRange(request.source, request.target);
          response = { kind: "ok", id: request.id };
          break;
        }

        case "copyRange": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          engine.copyRange(request.source, request.target);
          response = { kind: "ok", id: request.id };
          break;
        }

        case "pasteRange": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          engine.pasteRange(request.source, request.target);
          response = { kind: "ok", id: request.id };
          break;
        }

        case "getDependents": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          const dependencies = engine.getDependents(request.sheetName, request.address);
          response = {
            kind: "dependencies",
            id: request.id,
            addresses: dependencies.directDependents
          };
          break;
        }

        case "getPrecedents": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          const dependencies = engine.getDependencies(request.sheetName, request.address);
          response = {
            kind: "dependencies",
            id: request.id,
            addresses: dependencies.directPrecedents
          };
          break;
        }

        case "exportSnapshot": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          response = { kind: "snapshot", id: request.id, snapshot: engine.exportSnapshot() };
          break;
        }

        case "importSnapshot": {
          const session = this.getSessionByAgentSessionId(request.sessionId);
          session.engine.importSnapshot(request.snapshot);
          session.replicaSnapshot = session.engine.exportReplicaSnapshot();
          response = { kind: "ok", id: request.id };
          break;
        }

        case "getMetrics": {
          const { engine } = this.getSessionByAgentSessionId(request.sessionId);
          response = { kind: "metrics", id: request.id, value: engine.getLastMetrics() };
          break;
        }

        case "subscribeRange": {
          const session = this.getSessionByAgentSessionId(request.sessionId);
          const agentSession = session.agentSessions.get(request.sessionId);
          if (!agentSession) {
            throw new Error(`Unknown agent session: ${request.sessionId}`);
          }
          const existingOwner = this.agentSubscriptionOwners.get(request.subscriptionId);
          if (existingOwner) {
            throw new Error(`Subscription id already in use: ${request.subscriptionId}`);
          }

          const changedAddresses = iterateRange(request.range);
          const unsubscribe = session.engine.subscribeCells(request.range.sheetName, changedAddresses, () => {
            this.queueAgentEvent(session.documentId, {
              kind: "rangeChanged",
              subscriptionId: request.subscriptionId,
              range: request.range,
              changedAddresses: [...changedAddresses]
            });
          });

          session.agentSubscriptions.set(request.subscriptionId, {
            subscriptionId: request.subscriptionId,
            sessionId: request.sessionId,
            range: request.range,
            changedAddresses,
            unsubscribe
          });
          agentSession.subscriptionIds.add(request.subscriptionId);
          this.agentSubscriptionOwners.set(request.subscriptionId, request.sessionId);
          response = {
            kind: "ok",
            id: request.id,
            value: { subscriptionId: request.subscriptionId }
          };
          break;
        }

        case "unsubscribe": {
          const session = this.getSessionByAgentSessionId(request.sessionId);
          this.removeAgentSubscription(session, request.sessionId, request.subscriptionId);
          response = { kind: "ok", id: request.id };
          break;
        }

        default: {
          const exhaustiveRequest: never = request;
          response = {
            kind: "error",
            id: requestId,
            code: "UNSUPPORTED_AGENT_REQUEST",
            message: `Unsupported agent request ${(exhaustiveRequest as { kind: string }).kind}`,
            retryable: false
          };
          break;
        }
      }
    } catch (error) {
      response = {
        kind: "error",
        id: request.id,
        code: "LOCAL_SERVER_FAILURE",
        message: error instanceof Error ? error.message : String(error),
        retryable: false
      };
    }

    return { kind: "response", response };
  }

  getDocumentState(documentId: string): LocalDocumentStateSummary {
    const session = this.ensureSession(documentId);
    return {
      documentId,
      cursor: session.cursor,
      browserSessions: [...session.browserSubscribers.keys()],
      agentSessions: [...session.agentSessions.keys()],
      lastBatchId: session.batches.at(-1)?.frame.batch.id ?? null
    };
  }

  private readRange(engine: SpreadsheetEngine, range: CellRangeRef): CellValue[][] {
    const addresses = iterateRange(range);
    const [startColPart] = splitAddress(range.startAddress);
    const [endColPart] = splitAddress(range.endAddress);
    const width = decodeColumn(endColPart) - decodeColumn(startColPart) + 1;
    const rows: CellValue[][] = [];
    for (let index = 0; index < addresses.length; index += width) {
      rows.push(
        addresses.slice(index, index + width).map((address) => engine.getCellValue(range.sheetName, address))
      );
    }
    return rows;
  }

  private getSessionByAgentSessionId(sessionId: string): LocalWorkbookSession {
    const documentId = sessionId.split(":")[0];
    if (!documentId) {
      throw new Error(`Invalid session id: ${sessionId}`);
    }
    const session = this.sessions.get(documentId);
    if (!session || !session.agentSessions.has(sessionId)) {
      throw new Error(`Unknown agent session: ${sessionId}`);
    }
    return session;
  }

  private removeAgentSubscription(
    session: LocalWorkbookSession,
    sessionId: string,
    subscriptionId: string
  ): void {
    const subscription = session.agentSubscriptions.get(subscriptionId);
    if (!subscription) {
      return;
    }
    if (subscription.sessionId !== sessionId) {
      throw new Error(`Subscription ${subscriptionId} does not belong to agent session ${sessionId}`);
    }
    subscription.unsubscribe();
    session.agentSubscriptions.delete(subscriptionId);
    this.agentSubscriptionOwners.delete(subscriptionId);
    session.agentSessions.get(sessionId)?.subscriptionIds.delete(subscriptionId);
    session.eventBacklog = session.eventBacklog.filter((event) => {
      return event.kind !== "rangeChanged" || event.subscriptionId !== subscriptionId;
    });
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
    const session = this.sessions.get(documentId);
    if (!session) {
      return;
    }
    session.browserSubscribers.forEach((subscriber) => subscriber.send(frame));
  }

  private async relayUpstream(session: LocalWorkbookSession, batch: Extract<ProtocolFrame, { kind: "appendBatch" }>["batch"]): Promise<void> {
    if (!session.upstreamRelay) {
      return;
    }
    try {
      await session.upstreamRelay.send(batch);
    } catch (error) {
      console.error(`Failed to relay batch ${batch.id} for document ${session.documentId}:`, error);
    }
  }

  private ack(documentId: string, batchId: string, cursor: number): AckFrame {
    return {
      kind: "ack",
      documentId,
      batchId,
      cursor,
      acceptedAtUnixMs: Date.now()
    };
  }

  private agentError(id: string, code: string, message: string, retryable: boolean): AgentFrame {
    return {
      kind: "response",
      response: {
        kind: "error",
        id,
        code,
        message,
        retryable
      }
    };
  }
}
