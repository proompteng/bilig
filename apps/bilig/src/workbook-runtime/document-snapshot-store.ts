import type { ProtocolFrame } from "@bilig/binary-protocol";
import type { WorkbookSnapshot } from "@bilig/protocol";
import type { InMemoryDocumentPersistence } from "@bilig/storage-server";
import {
  acceptSnapshotChunk,
  createSnapshotPublication,
  type SnapshotAssemblyRegistry,
} from "./session-shared.js";
import { createCursorWatermarkFrame } from "./sync-frame-shared.js";

export async function publishPersistedSnapshot(
  persistence: InMemoryDocumentPersistence,
  documentId: string,
  snapshot: WorkbookSnapshot,
  broadcast: (documentId: string, frame: ProtocolFrame) => void,
): Promise<void> {
  const cursor = (await persistence.batches.latestCursor(documentId)) + 1;
  const publication = createSnapshotPublication(documentId, cursor, snapshot);

  await persistence.batches.reset(documentId, cursor);
  await persistence.snapshots.put({
    documentId,
    snapshotId: publication.snapshotId,
    cursor,
    contentType: publication.contentType,
    bytes: publication.bytes,
    createdAtUnixMs: Date.now(),
  });
  publication.frames.forEach((frame) => broadcast(documentId, frame));
  broadcast(documentId, createCursorWatermarkFrame(documentId, cursor, cursor));
}

export async function acceptPersistedSnapshotChunk(
  persistence: InMemoryDocumentPersistence,
  snapshotAssemblies: SnapshotAssemblyRegistry,
  frame: Extract<ProtocolFrame, { kind: "snapshotChunk" }>,
): Promise<void> {
  const snapshot = acceptSnapshotChunk(snapshotAssemblies, frame);
  if (!snapshot) {
    return;
  }
  await persistence.snapshots.put({
    documentId: snapshot.documentId,
    snapshotId: snapshot.snapshotId,
    cursor: snapshot.cursor,
    contentType: snapshot.contentType,
    bytes: snapshot.bytes,
    createdAtUnixMs: Date.now(),
  });
}
