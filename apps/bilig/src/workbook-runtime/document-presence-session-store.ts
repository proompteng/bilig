import type { InMemoryDocumentPersistence } from "@bilig/storage-server";
import { normalizeSessionId } from "./agent-routing.js";
import { documentIdFromSessionId } from "./workbook-session-shared.js";

export async function openPresenceBackedWorkbookSession(
  persistence: InMemoryDocumentPersistence,
  documentId: string,
  replicaId: string,
): Promise<string> {
  const sessionId = normalizeSessionId(documentId, replicaId);
  await persistence.presence.join(documentId, sessionId);
  return sessionId;
}

export async function closePresenceBackedWorkbookSession(
  persistence: InMemoryDocumentPersistence,
  sessionId: string,
): Promise<void> {
  await persistence.presence.leave(documentIdFromSessionId(sessionId), sessionId);
}

export async function countPresenceBackedWorkbookSessions(
  persistence: InMemoryDocumentPersistence,
  sessionId: string,
): Promise<number> {
  return (await persistence.presence.sessions(documentIdFromSessionId(sessionId))).length;
}

export async function joinOwnedBrowserSession(
  persistence: InMemoryDocumentPersistence,
  ownerId: string,
  documentId: string,
  sessionId: string,
): Promise<void> {
  await persistence.presence.join(documentId, sessionId);
  await persistence.ownership.claim(documentId, ownerId, Date.now() + 60_000);
}
