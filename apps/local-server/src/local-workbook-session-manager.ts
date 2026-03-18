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
  PROTOCOL_VERSION,
  type ProtocolFrame
} from "@bilig/binary-protocol";
import type { AgentEvent, AgentFrame, AgentResponse } from "@bilig/agent-api";
import { shouldApplyBatch } from "@bilig/crdt";

interface BrowserSubscriber {
  id: string;
  send(frame: ProtocolFrame): void;
}

interface AgentSession {
  sessionId: string;
  documentId: string;
  replicaId: string;
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
  eventBacklog: AgentEvent[];
  batches: StoredBatch[];
  cursor: number;
  replicaSnapshot: EngineReplicaSnapshot | null;
  unsubscribeBatches: () => void;
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
      eventBacklog: [],
      batches: [],
      cursor: 0,
      replicaSnapshot: null,
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
            replicaId: request.replicaId
          });
          response = { kind: "ok", id: request.id, sessionId };
          break;
        }

        case "closeWorkbookSession": {
          const session = this.getSessionByAgentSessionId(request.sessionId);
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

        case "subscribeRange":
        case "unsubscribe":
          response = {
            kind: "error",
            id: request.id,
            code: "AGENT_STREAM_NOT_WIRED",
            message: `${request.kind} requires the agent event stream transport, which is not wired in this tranche`,
            retryable: false
          };
          break;

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
    const [startColPart, startRowPart] = splitAddress(range.startAddress);
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

  private broadcast(documentId: string, frame: ProtocolFrame): void {
    const session = this.sessions.get(documentId);
    if (!session) {
      return;
    }
    session.browserSubscribers.forEach((subscriber) => subscriber.send(frame));
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
