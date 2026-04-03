import { SpreadsheetEngine, type EngineReplicaSnapshot } from "@bilig/core";
import type { CellRangeRef, CellValue, EngineEvent } from "@bilig/protocol";
import type {
  AckFrame,
  CursorWatermarkFrame,
  ErrorFrame,
  HeartbeatFrame,
  ProtocolFrame,
  SnapshotChunkFrame,
} from "@bilig/binary-protocol";
import type {
  AgentEvent,
  AgentFrame,
  AgentResponse,
  LoadWorkbookFileRequest,
} from "@bilig/agent-api";
import { shouldApplyBatch } from "@bilig/crdt";
import type { UpstreamSyncRelay } from "../zero/sync-relay.js";
import {
  type AgentFrameContext,
  createWorkbookLoadedResponse,
  normalizeSessionId,
  prepareWorkbookLoad,
  routeAgentFrame,
  type WorksheetAgentRequest,
} from "./agent-routing.js";
import { createSnapshotPublication } from "./session-shared.js";

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
  unsubscribe: () => void;
}

interface StoredBatch {
  cursor: number;
  frame: Extract<ProtocolFrame, { kind: "appendBatch" }>;
}

interface StoredSnapshotPublication {
  cursor: number;
  snapshotId: string;
  contentType: string;
  bytes: Uint8Array;
  frames: SnapshotChunkFrame[];
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
  latestSnapshot: StoredSnapshotPublication | null;
  cursor: number;
  replicaSnapshot: EngineReplicaSnapshot | null;
  upstreamRelay: UpstreamSyncRelay | null;
  unsubscribeBatches: () => void;
  compactScheduled: boolean;
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

function cellCountForRange(range: CellRangeRef): number {
  const [startColPart, startRowPart] = splitAddress(range.startAddress);
  const [endColPart, endRowPart] = splitAddress(range.endAddress);
  const width = decodeColumn(endColPart) - decodeColumn(startColPart) + 1;
  const height = Number.parseInt(endRowPart, 10) - Number.parseInt(startRowPart, 10) + 1;
  return width * height;
}

function collectChangedAddressesInRange(
  engine: SpreadsheetEngine,
  range: CellRangeRef,
  changedCellIndices: readonly number[] | Uint32Array,
): string[] {
  const [startColPart, startRowPart] = splitAddress(range.startAddress);
  const [endColPart, endRowPart] = splitAddress(range.endAddress);
  const startCol = decodeColumn(startColPart);
  const endCol = decodeColumn(endColPart);
  const startRow = Number.parseInt(startRowPart, 10);
  const endRow = Number.parseInt(endRowPart, 10);
  const changedAddresses: string[] = [];

  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const qualifiedAddress = engine.workbook.getQualifiedAddress(changedCellIndices[index]!);
    if (!qualifiedAddress.startsWith(`${range.sheetName}!`)) {
      continue;
    }
    const address = qualifiedAddress.slice(range.sheetName.length + 1);
    const parsed = splitAddress(address);
    const col = decodeColumn(parsed[0]);
    const row = Number.parseInt(parsed[1], 10);
    if (col < startCol || col > endCol || row < startRow || row > endRow) {
      continue;
    }
    changedAddresses.push(address);
  }

  return changedAddresses;
}

function collectAddressesForIntersection(
  range: CellRangeRef,
  startAddress: string,
  endAddress: string,
): string[] {
  const rangeStart = splitAddress(range.startAddress);
  const rangeEnd = splitAddress(range.endAddress);
  const eventStart = splitAddress(startAddress);
  const eventEnd = splitAddress(endAddress);

  const startCol = Math.max(decodeColumn(rangeStart[0]), decodeColumn(eventStart[0]));
  const endCol = Math.min(decodeColumn(rangeEnd[0]), decodeColumn(eventEnd[0]));
  const startRow = Math.max(Number.parseInt(rangeStart[1], 10), Number.parseInt(eventStart[1], 10));
  const endRow = Math.min(Number.parseInt(rangeEnd[1], 10), Number.parseInt(eventEnd[1], 10));

  if (startCol > endCol || startRow > endRow) {
    return [];
  }

  const changedAddresses: string[] = [];
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      changedAddresses.push(`${encodeColumn(col)}${row}`);
    }
  }
  return changedAddresses;
}

function collectChangedAddressesForEvent(
  engine: SpreadsheetEngine,
  range: CellRangeRef,
  event: EngineEvent,
): string[] {
  if (event.invalidation === "full") {
    return iterateRange(range);
  }

  const changedAddresses = new Set(
    collectChangedAddressesInRange(engine, range, event.changedCellIndices),
  );
  for (let index = 0; index < event.invalidatedRanges.length; index += 1) {
    const invalidatedRange = event.invalidatedRanges[index]!;
    if (invalidatedRange.sheetName !== range.sheetName) {
      continue;
    }
    collectAddressesForIntersection(
      range,
      invalidatedRange.startAddress,
      invalidatedRange.endAddress,
    ).forEach((address) => {
      changedAddresses.add(address);
    });
  }
  return [...changedAddresses];
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
      replicaId: `worksheet-host:${documentId}`,
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
      latestSnapshot: null,
      cursor: 0,
      replicaSnapshot: null,
      upstreamRelay: this.options.createSyncRelay?.(documentId) ?? null,
      unsubscribeBatches: () => {},
      compactScheduled: false,
    };

    session.unsubscribeBatches = engine.subscribeBatches((batch) => {
      session.cursor += 1;
      const frame = {
        kind: "appendBatch",
        documentId,
        cursor: session.cursor,
        batch,
      } satisfies Extract<ProtocolFrame, { kind: "appendBatch" }>;
      session.batches.push({ cursor: session.cursor, frame });
      session.replicaSnapshot = engine.exportReplicaSnapshot();
      this.broadcast(documentId, frame);
      this.maybeCompactSession(session);
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
        const snapshotFrames =
          session.latestSnapshot && frame.lastServerCursor < session.latestSnapshot.cursor
            ? session.latestSnapshot.frames
            : [];
        const missed = session.batches
          .filter(
            (entry) =>
              entry.cursor > Math.max(frame.lastServerCursor, session.latestSnapshot?.cursor ?? 0),
          )
          .map((entry) => entry.frame);
        return [
          ...snapshotFrames,
          ...missed,
          {
            kind: "cursorWatermark",
            documentId: frame.documentId,
            cursor: session.cursor,
            compactedCursor:
              session.latestSnapshot?.cursor ??
              Math.max(0, session.cursor - session.batches.length),
          } satisfies CursorWatermarkFrame,
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
          batch: frame.batch,
        } satisfies Extract<ProtocolFrame, { kind: "appendBatch" }>;
        session.batches.push({ cursor: session.cursor, frame: committedFrame });
        session.replicaSnapshot = session.engine.exportReplicaSnapshot();
        this.broadcast(frame.documentId, committedFrame);
        this.maybeCompactSession(session);
        void this.relayUpstream(session, frame.batch);
        return [this.ack(frame.documentId, frame.batch.id, session.cursor)];
      }

      case "heartbeat":
        return [
          {
            kind: "heartbeat",
            documentId: frame.documentId,
            cursor: session.cursor,
            sentAtUnixMs: Date.now(),
          } satisfies HeartbeatFrame,
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
            retryable: false,
          } satisfies ErrorFrame,
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
            retryable: false,
          } satisfies ErrorFrame,
        ];
    }
  }

  async handleAgentFrame(frame: AgentFrame, context: AgentFrameContext = {}): Promise<AgentFrame> {
    return routeAgentFrame(frame, context, {
      invalidFrameMessage: "Local server accepts only agent requests",
      errorCode: "LOCAL_SERVER_FAILURE",
      loadWorkbookFile: (request, requestContext) => this.loadWorkbookFile(request, requestContext),
      openWorkbookSession: (request) => {
        const session = this.ensureSession(request.documentId);
        const sessionId = normalizeSessionId(request.documentId, request.replicaId);
        session.agentSessions.set(sessionId, {
          sessionId,
          documentId: request.documentId,
          replicaId: request.replicaId,
          subscriptionIds: new Set(),
        });
        return { kind: "ok", id: request.id, sessionId };
      },
      closeWorkbookSession: (request) => {
        const session = this.getSessionByAgentSessionId(request.sessionId);
        const agentSession = session.agentSessions.get(request.sessionId);
        if (agentSession) {
          [...agentSession.subscriptionIds].forEach((subscriptionId) => {
            this.removeAgentSubscription(session, request.sessionId, subscriptionId);
          });
        }
        session.agentSessions.delete(request.sessionId);
        return { kind: "ok", id: request.id };
      },
      getMetrics: (request) => {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        return { kind: "metrics", id: request.id, value: engine.getLastMetrics() };
      },
      handleWorksheetRequest: (_frame, request) => this.handleWorksheetRequest(request),
    });
  }

  getDocumentState(documentId: string): LocalDocumentStateSummary {
    const session = this.ensureSession(documentId);
    return {
      documentId,
      cursor: session.cursor,
      browserSessions: [...session.browserSubscribers.keys()],
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

  private readRange(engine: SpreadsheetEngine, range: CellRangeRef): CellValue[][] {
    const [startColPart] = splitAddress(range.startAddress);
    const [endColPart] = splitAddress(range.endAddress);
    const [, startRowPart] = splitAddress(range.startAddress);
    const [, endRowPart] = splitAddress(range.endAddress);
    const startCol = decodeColumn(startColPart);
    const endCol = decodeColumn(endColPart);
    const startRow = Number.parseInt(startRowPart, 10);
    const endRow = Number.parseInt(endRowPart, 10);
    const width = decodeColumn(endColPart) - decodeColumn(startColPart) + 1;
    const rows: CellValue[][] = [];
    for (let row = startRow; row <= endRow; row += 1) {
      const nextRow: CellValue[] = Array.from<CellValue>({ length: width });
      for (let col = startCol; col <= endCol; col += 1) {
        nextRow[col - startCol] = engine.getCellValue(
          range.sheetName,
          `${encodeColumn(col)}${row}`,
        );
      }
      rows.push(nextRow);
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

  private loadWorkbookFile(
    request: LoadWorkbookFileRequest,
    context: AgentFrameContext,
  ): AgentResponse {
    const prepared = prepareWorkbookLoad(request, context, {
      ...(this.options.maxImportBytes !== undefined
        ? { maxImportBytes: this.options.maxImportBytes }
        : {}),
      ...(this.options.publicServerUrl ? { publicServerUrl: this.options.publicServerUrl } : {}),
      ...(this.options.browserAppBaseUrl
        ? { browserAppBaseUrl: this.options.browserAppBaseUrl }
        : {}),
    });
    const session = this.ensureSession(prepared.documentId);
    session.agentSessions.set(prepared.sessionId, {
      sessionId: prepared.sessionId,
      documentId: prepared.documentId,
      replicaId: request.replicaId,
      subscriptionIds: session.agentSessions.get(prepared.sessionId)?.subscriptionIds ?? new Set(),
    });

    session.engine.importSnapshot(prepared.imported.snapshot);
    session.replicaSnapshot = session.engine.exportReplicaSnapshot();
    this.publishSnapshot(session, prepared.imported.snapshot);

    return createWorkbookLoadedResponse(request.id, prepared);
  }

  private handleWorksheetRequest(request: WorksheetAgentRequest): AgentResponse {
    switch (request.kind) {
      case "readRange": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        return {
          kind: "rangeValues",
          id: request.id,
          values: this.readRange(engine, request.range),
        };
      }
      case "writeRange": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.setRangeValues(request.range, request.values);
        return { kind: "ok", id: request.id };
      }
      case "setRangeFormulas": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.setRangeFormulas(request.range, request.formulas);
        return { kind: "ok", id: request.id };
      }
      case "setRangeStyle": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.setRangeStyle(request.range, request.patch);
        return { kind: "ok", id: request.id };
      }
      case "clearRangeStyle": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.clearRangeStyle(request.range, request.fields);
        return { kind: "ok", id: request.id };
      }
      case "setRangeNumberFormat": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.setRangeNumberFormat(request.range, request.format);
        return { kind: "ok", id: request.id };
      }
      case "clearRangeNumberFormat": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.clearRangeNumberFormat(request.range);
        return { kind: "ok", id: request.id };
      }
      case "clearRange": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.clearRange(request.range);
        return { kind: "ok", id: request.id };
      }
      case "fillRange": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.fillRange(request.source, request.target);
        return { kind: "ok", id: request.id };
      }
      case "copyRange": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.copyRange(request.source, request.target);
        return { kind: "ok", id: request.id };
      }
      case "pasteRange": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        engine.pasteRange(request.source, request.target);
        return { kind: "ok", id: request.id };
      }
      case "getDependents": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        const dependencies = engine.getDependents(request.sheetName, request.address);
        return {
          kind: "dependencies",
          id: request.id,
          addresses: dependencies.directDependents,
        };
      }
      case "getPrecedents": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        const dependencies = engine.getDependencies(request.sheetName, request.address);
        return {
          kind: "dependencies",
          id: request.id,
          addresses: dependencies.directPrecedents,
        };
      }
      case "exportSnapshot": {
        const { engine } = this.getSessionByAgentSessionId(request.sessionId);
        return { kind: "snapshot", id: request.id, snapshot: engine.exportSnapshot() };
      }
      case "importSnapshot": {
        const session = this.getSessionByAgentSessionId(request.sessionId);
        session.engine.importSnapshot(request.snapshot);
        session.replicaSnapshot = session.engine.exportReplicaSnapshot();
        return { kind: "ok", id: request.id };
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

        const rangeCellCount = cellCountForRange(request.range);
        const changedAddresses =
          rangeCellCount <= LARGE_RANGE_SUBSCRIPTION_THRESHOLD ? iterateRange(request.range) : null;
        const unsubscribe = changedAddresses
          ? session.engine.subscribeCells(request.range.sheetName, changedAddresses, () => {
              this.queueAgentEvent(session.documentId, {
                kind: "rangeChanged",
                subscriptionId: request.subscriptionId,
                range: request.range,
                changedAddresses: [...changedAddresses],
              });
            })
          : session.engine.subscribe((event: EngineEvent) => {
              const changedInRange = collectChangedAddressesForEvent(
                session.engine,
                request.range,
                event,
              );
              if (changedInRange.length === 0) {
                return;
              }
              this.queueAgentEvent(session.documentId, {
                kind: "rangeChanged",
                subscriptionId: request.subscriptionId,
                range: request.range,
                changedAddresses: changedInRange,
              });
            });

        session.agentSubscriptions.set(request.subscriptionId, {
          subscriptionId: request.subscriptionId,
          sessionId: request.sessionId,
          range: request.range,
          unsubscribe,
        });
        agentSession.subscriptionIds.add(request.subscriptionId);
        this.agentSubscriptionOwners.set(request.subscriptionId, request.sessionId);
        return {
          kind: "ok",
          id: request.id,
          value: { subscriptionId: request.subscriptionId },
        };
      }
      case "unsubscribe": {
        const session = this.getSessionByAgentSessionId(request.sessionId);
        this.removeAgentSubscription(session, request.sessionId, request.subscriptionId);
        return { kind: "ok", id: request.id };
      }
      case "createPivotTable": {
        const session = this.getSessionByAgentSessionId(request.sessionId);
        session.engine.setPivotTable(request.sheetName, request.address, {
          name: request.name,
          source: request.source,
          groupBy: request.groupBy,
          values: request.values,
        });
        return { kind: "ok", id: request.id };
      }
      default: {
        const exhaustiveRequest: never = request;
        return {
          kind: "error",
          id: "unknown",
          code: "UNSUPPORTED_AGENT_REQUEST",
          message: `Unsupported agent request ${(exhaustiveRequest as { kind: string }).kind}`,
          retryable: false,
        };
      }
    }
  }

  private removeAgentSubscription(
    session: LocalWorkbookSession,
    sessionId: string,
    subscriptionId: string,
  ): void {
    const subscription = session.agentSubscriptions.get(subscriptionId);
    if (!subscription) {
      return;
    }
    if (subscription.sessionId !== sessionId) {
      throw new Error(
        `Subscription ${subscriptionId} does not belong to agent session ${sessionId}`,
      );
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

  private publishSnapshot(
    session: LocalWorkbookSession,
    snapshot: Parameters<SpreadsheetEngine["importSnapshot"]>[0],
  ): void {
    const cursor = session.cursor + 1;
    const publication = createSnapshotPublication(session.documentId, cursor, snapshot);
    session.cursor = cursor;
    session.batches = [];
    session.latestSnapshot = {
      cursor,
      snapshotId: publication.snapshotId,
      contentType: publication.contentType,
      bytes: publication.bytes,
      frames: publication.frames,
    };
    publication.frames.forEach((frame) => this.broadcast(session.documentId, frame));
    this.broadcast(session.documentId, {
      kind: "cursorWatermark",
      documentId: session.documentId,
      cursor,
      compactedCursor: cursor,
    });
  }

  private maybeCompactSession(session: LocalWorkbookSession): void {
    if (session.batches.length <= MAX_BATCH_BACKLOG || session.compactScheduled) {
      return;
    }
    session.compactScheduled = true;
    setImmediate(() => {
      session.compactScheduled = false;
      const liveSession = this.sessions.get(session.documentId);
      if (!liveSession || liveSession.batches.length <= MAX_BATCH_BACKLOG) {
        return;
      }
      this.publishSnapshot(liveSession, liveSession.engine.exportSnapshot());
    });
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

  private ack(documentId: string, batchId: string, cursor: number): AckFrame {
    return {
      kind: "ack",
      documentId,
      batchId,
      cursor,
      acceptedAtUnixMs: Date.now(),
    };
  }
}
