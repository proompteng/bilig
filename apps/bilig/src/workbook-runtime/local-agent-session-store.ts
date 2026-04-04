import type { CellRangeRef } from "@bilig/protocol";
import type { RangeBounds } from "./range-subscription-utils.js";
import { normalizeSessionId } from "./agent-routing.js";
import { documentIdFromSessionId } from "./workbook-session-shared.js";

export interface LocalAgentSessionState {
  sessionId: string;
  documentId: string;
  replicaId: string;
  subscriptionIds: Set<string>;
}

export interface LocalAgentRangeSubscriptionState {
  subscriptionId: string;
  sessionId: string;
  range: CellRangeRef;
  bounds: RangeBounds;
  watchedAddresses: readonly string[] | null;
  unsubscribe: () => void;
}

export interface LocalAgentSessionContainer {
  documentId: string;
  agentSessions: Map<string, LocalAgentSessionState>;
  agentSubscriptions: Map<string, LocalAgentRangeSubscriptionState>;
}

export function openLocalAgentSession(
  session: LocalAgentSessionContainer,
  replicaId: string,
): string {
  const sessionId = normalizeSessionId(session.documentId, replicaId);
  session.agentSessions.set(sessionId, {
    sessionId,
    documentId: session.documentId,
    replicaId,
    subscriptionIds: session.agentSessions.get(sessionId)?.subscriptionIds ?? new Set(),
  });
  return sessionId;
}

export function getLocalSessionByAgentSessionId<SessionState extends LocalAgentSessionContainer>(
  sessions: Map<string, SessionState>,
  sessionId: string,
): SessionState {
  const documentId = documentIdFromSessionId(sessionId);
  const session = sessions.get(documentId);
  if (!session || !session.agentSessions.has(sessionId)) {
    throw new Error(`Unknown agent session: ${sessionId}`);
  }
  return session;
}

export function closeLocalAgentSession<SessionState extends LocalAgentSessionContainer>(
  session: SessionState,
  sessionId: string,
  removeAgentSubscription: (
    session: SessionState,
    sessionId: string,
    subscriptionId: string,
  ) => void,
): void {
  const agentSession = session.agentSessions.get(sessionId);
  if (agentSession) {
    [...agentSession.subscriptionIds].forEach((subscriptionId) => {
      removeAgentSubscription(session, sessionId, subscriptionId);
    });
  }
  session.agentSessions.delete(sessionId);
}

export function removeLocalAgentSubscription<SessionState extends LocalAgentSessionContainer>(
  session: SessionState,
  sessionId: string,
  subscriptionId: string,
  agentSubscriptionOwners: Map<string, string>,
  onRemove?: (subscriptionId: string) => void,
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
  agentSubscriptionOwners.delete(subscriptionId);
  session.agentSessions.get(sessionId)?.subscriptionIds.delete(subscriptionId);
  onRemove?.(subscriptionId);
}
