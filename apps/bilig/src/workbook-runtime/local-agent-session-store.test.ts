import { describe, expect, it, vi } from "vitest";
import {
  closeLocalAgentSession,
  getLocalSessionByAgentSessionId,
  openLocalAgentSession,
  removeLocalAgentSubscription,
  type LocalAgentSessionContainer,
} from "./local-agent-session-store.js";

function createSession(documentId: string): LocalAgentSessionContainer {
  return {
    documentId,
    agentSessions: new Map(),
    agentSubscriptions: new Map(),
  };
}

describe("local-agent-session-store", () => {
  it("opens and resolves local agent sessions by session id", () => {
    const session = createSession("doc-1");
    const sessions = new Map([[session.documentId, session]]);

    const sessionId = openLocalAgentSession(session, "replica-1");

    expect(sessionId).toBe("doc-1:replica-1");
    expect(getLocalSessionByAgentSessionId(sessions, sessionId)).toBe(session);
  });

  it("closes agent sessions and removes owned subscriptions", () => {
    const session = createSession("doc-2");
    const sessionId = openLocalAgentSession(session, "replica-2");
    session.agentSubscriptions.set("sub-1", {
      subscriptionId: "sub-1",
      sessionId,
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "A1",
      },
      unsubscribe: vi.fn(),
    });
    session.agentSessions.get(sessionId)?.subscriptionIds.add("sub-1");

    const removed: string[] = [];
    closeLocalAgentSession(session, sessionId, (_ownedSession, _ownedSessionId, subscriptionId) => {
      removed.push(subscriptionId);
    });

    expect(removed).toEqual(["sub-1"]);
    expect(session.agentSessions.has(sessionId)).toBe(false);
  });

  it("removes subscriptions and ownership bookkeeping", () => {
    const session = createSession("doc-3");
    const sessionId = openLocalAgentSession(session, "replica-3");
    const unsubscribe = vi.fn();
    const ownership = new Map<string, string>([["sub-3", sessionId]]);
    session.agentSubscriptions.set("sub-3", {
      subscriptionId: "sub-3",
      sessionId,
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "B2",
      },
      unsubscribe,
    });
    session.agentSessions.get(sessionId)?.subscriptionIds.add("sub-3");

    removeLocalAgentSubscription(session, sessionId, "sub-3", ownership);

    expect(unsubscribe).toHaveBeenCalledTimes(1);
    expect(session.agentSubscriptions.has("sub-3")).toBe(false);
    expect(session.agentSessions.get(sessionId)?.subscriptionIds.has("sub-3")).toBe(false);
    expect(ownership.has("sub-3")).toBe(false);
  });
});
