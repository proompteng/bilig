import type { EngineReplicaSnapshot, SpreadsheetEngine } from "@bilig/core";
import type { ProtocolFrame, SnapshotChunkFrame } from "@bilig/binary-protocol";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  acceptSnapshotChunk,
  createSnapshotPublication,
  createSnapshotPublicationFromBytes,
  decodeWorkbookSnapshotBytes,
  type SnapshotAssemblyRegistry,
} from "./session-shared.js";
import { createCursorWatermarkFrame } from "./sync-frame-shared.js";

export interface StoredBatch {
  cursor: number;
  frame: Extract<ProtocolFrame, { kind: "appendBatch" }>;
}

export interface StoredSnapshotPublication {
  cursor: number;
  snapshotId: string;
  contentType: string;
  bytes: Uint8Array;
  frames: SnapshotChunkFrame[];
}

export interface LocalSnapshotSessionState {
  documentId: string;
  engine: SpreadsheetEngine;
  batches: StoredBatch[];
  latestSnapshot: StoredSnapshotPublication | null;
  snapshotCache: WorkbookSnapshot | null;
  snapshotDirty: boolean;
  cursor: number;
  replicaSnapshot: EngineReplicaSnapshot | null;
  compactScheduled: boolean;
}

export interface LocalSnapshotStoreContext<
  SessionState extends LocalSnapshotSessionState = LocalSnapshotSessionState,
> {
  broadcast(documentId: string, frame: ProtocolFrame): void;
  getSession(documentId: string): SessionState | undefined;
  snapshotAssemblies: SnapshotAssemblyRegistry;
  maxBatchBacklog: number;
  schedule?(callback: () => void): void;
}

export function invalidateLocalSnapshotCache<SessionState extends LocalSnapshotSessionState>(
  session: SessionState,
): void {
  session.snapshotDirty = true;
}

export function storeLocalCachedSnapshot<SessionState extends LocalSnapshotSessionState>(
  session: SessionState,
  snapshot: WorkbookSnapshot,
): void {
  session.snapshotCache = snapshot;
  session.snapshotDirty = false;
}

export function getLocalCachedSnapshot<SessionState extends LocalSnapshotSessionState>(
  session: SessionState,
): WorkbookSnapshot {
  if (session.snapshotCache && !session.snapshotDirty) {
    return session.snapshotCache;
  }
  const snapshot = session.engine.exportSnapshot();
  storeLocalCachedSnapshot(session, snapshot);
  return snapshot;
}

export function publishLocalSnapshot<SessionState extends LocalSnapshotSessionState>(
  session: SessionState,
  snapshot: WorkbookSnapshot,
  broadcast: LocalSnapshotStoreContext<SessionState>["broadcast"],
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
  storeLocalCachedSnapshot(session, snapshot);
  publication.frames.forEach((frame) => broadcast(session.documentId, frame));
  broadcast(session.documentId, createCursorWatermarkFrame(session.documentId, cursor, cursor));
}

export function acceptLocalSnapshotChunk<SessionState extends LocalSnapshotSessionState>(
  session: SessionState,
  frame: Extract<ProtocolFrame, { kind: "snapshotChunk" }>,
  context: Pick<LocalSnapshotStoreContext<SessionState>, "broadcast" | "snapshotAssemblies">,
): void {
  const assembled = acceptSnapshotChunk(context.snapshotAssemblies, frame);
  if (!assembled) {
    return;
  }

  const snapshot = decodeWorkbookSnapshotBytes(assembled);
  const publication = createSnapshotPublicationFromBytes(assembled);
  session.engine.importSnapshot(snapshot);
  session.replicaSnapshot = session.engine.exportReplicaSnapshot();
  session.cursor = assembled.cursor;
  session.batches = [];
  session.latestSnapshot = {
    cursor: assembled.cursor,
    snapshotId: publication.snapshotId,
    contentType: publication.contentType,
    bytes: publication.bytes,
    frames: publication.frames,
  };
  storeLocalCachedSnapshot(session, snapshot);
  publication.frames.forEach((snapshotFrame) =>
    context.broadcast(session.documentId, snapshotFrame),
  );
  context.broadcast(
    session.documentId,
    createCursorWatermarkFrame(session.documentId, assembled.cursor, assembled.cursor),
  );
}

export function maybeCompactLocalSession<SessionState extends LocalSnapshotSessionState>(
  session: SessionState,
  context: LocalSnapshotStoreContext<SessionState>,
): void {
  if (session.batches.length <= context.maxBatchBacklog || session.compactScheduled) {
    return;
  }
  session.compactScheduled = true;
  const schedule =
    context.schedule === undefined
      ? (callback: () => void) => setImmediate(callback)
      : (callback: () => void) => context.schedule?.(callback);
  const broadcast = (documentId: string, frame: ProtocolFrame) => {
    context.broadcast(documentId, frame);
  };
  schedule(() => {
    session.compactScheduled = false;
    const liveSession = context.getSession(session.documentId);
    if (!liveSession || liveSession.batches.length <= context.maxBatchBacklog) {
      return;
    }
    publishLocalSnapshot(liveSession, getLocalCachedSnapshot(liveSession), broadcast);
  });
}
