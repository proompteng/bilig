import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  WorkbookAgentSessionSnapshotSchema,
  WorkbookAgentStreamEventSchema,
  decodeUnknownSync,
  type WorkbookAgentSessionSnapshot,
  type WorkbookAgentStreamEvent,
  type WorkbookAgentUiContext,
} from "@bilig/contracts";
import { WorkbookAgentPanel } from "./WorkbookAgentPanel.js";

const STORAGE_KEY_PREFIX = "bilig:workbook-agent:";

interface StoredWorkbookAgentSession {
  sessionId: string;
  threadId: string;
}

function storageKey(documentId: string): string {
  return `${STORAGE_KEY_PREFIX}${documentId}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isStoredWorkbookAgentSession(value: unknown): value is StoredWorkbookAgentSession {
  return (
    isRecord(value) &&
    typeof value["sessionId"] === "string" &&
    typeof value["threadId"] === "string"
  );
}

function resolvePayloadMessage(payload: unknown, fallback: string): string {
  return isRecord(payload) && typeof payload["message"] === "string"
    ? payload["message"]
    : fallback;
}

function readMessageEventData(event: Event): string | null {
  return event instanceof MessageEvent && typeof event.data === "string" ? event.data : null;
}

function loadStoredSession(documentId: string): StoredWorkbookAgentSession | null {
  try {
    const raw = window.sessionStorage.getItem(storageKey(documentId));
    if (!raw) {
      return null;
    }
    const parsed = JSON.parse(raw) as unknown;
    if (isStoredWorkbookAgentSession(parsed)) {
      return parsed;
    }
  } catch {}
  return null;
}

function persistStoredSession(documentId: string, value: StoredWorkbookAgentSession): void {
  window.sessionStorage.setItem(storageKey(documentId), JSON.stringify(value));
}

function updateSnapshotFromDelta(
  snapshot: WorkbookAgentSessionSnapshot | null,
  event: Extract<WorkbookAgentStreamEvent, { type: "assistantDelta" | "planDelta" }>,
): WorkbookAgentSessionSnapshot | null {
  if (!snapshot) {
    return snapshot;
  }
  return {
    ...snapshot,
    entries: snapshot.entries.map((entry) => {
      if (entry.id !== event.itemId) {
        return entry;
      }
      return {
        ...entry,
        text: `${entry.text ?? ""}${event.delta}`,
      };
    }),
  };
}

export function useWorkbookAgentPane(input: {
  readonly documentId: string;
  readonly enabled: boolean;
  readonly contextKey: string;
  readonly getContext: () => WorkbookAgentUiContext;
}) {
  const { contextKey, documentId, enabled, getContext } = input;
  const [snapshot, setSnapshot] = useState<WorkbookAgentSessionSnapshot | null>(null);
  const [draft, setDraft] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isOpen, setIsOpen] = useState(true);
  const sessionRef = useRef<StoredWorkbookAgentSession | null>(null);
  const eventSourceRef = useRef<EventSource | null>(null);
  const lastContextKeyRef = useRef<string>("");
  const getContextRef = useRef(getContext);
  const currentContext = getContextRef.current();

  useEffect(() => {
    getContextRef.current = getContext;
  }, [getContext]);

  const closeStream = useCallback(() => {
    eventSourceRef.current?.close();
    eventSourceRef.current = null;
  }, []);

  const connectStream = useCallback(
    (sessionId: string) => {
      closeStream();
      const source = new EventSource(
        `/v2/documents/${encodeURIComponent(documentId)}/agent/sessions/${encodeURIComponent(sessionId)}/events`,
      );
      source.addEventListener("message", (message) => {
        try {
          const payloadText = readMessageEventData(message);
          if (payloadText === null) {
            return;
          }
          const event = decodeUnknownSync(WorkbookAgentStreamEventSchema, JSON.parse(payloadText));
          if (event.type === "snapshot") {
            setSnapshot(event.snapshot);
            persistStoredSession(documentId, {
              sessionId: event.snapshot.sessionId,
              threadId: event.snapshot.threadId,
            });
            sessionRef.current = {
              sessionId: event.snapshot.sessionId,
              threadId: event.snapshot.threadId,
            };
            setError(null);
            return;
          }
          setSnapshot((current: WorkbookAgentSessionSnapshot | null) =>
            updateSnapshotFromDelta(current, event),
          );
        } catch (nextError) {
          setError(nextError instanceof Error ? nextError.message : String(nextError));
        }
      });
      source.addEventListener("error", () => {
        setError("Assistant stream disconnected. Retrying...");
      });
      eventSourceRef.current = source;
    },
    [closeStream, documentId],
  );

  useEffect(() => {
    if (!enabled) {
      setIsLoading(false);
      return;
    }
    let cancelled = false;
    const context = getContextRef.current();
    const bootstrap = async () => {
      try {
        setIsLoading(true);
        const storedSession = loadStoredSession(documentId);
        sessionRef.current = storedSession;
        const response = await fetch(
          `/v2/documents/${encodeURIComponent(documentId)}/agent/sessions`,
          {
            method: "POST",
            headers: {
              "content-type": "application/json",
            },
            body: JSON.stringify({
              ...storedSession,
              context,
            }),
          },
        );
        const payload = (await response.json()) as unknown;
        if (!response.ok) {
          throw new Error(
            resolvePayloadMessage(
              payload,
              `Workbook agent request failed with status ${response.status}`,
            ),
          );
        }
        const nextSnapshot = decodeUnknownSync(WorkbookAgentSessionSnapshotSchema, payload);
        if (cancelled) {
          return;
        }
        setSnapshot(nextSnapshot);
        persistStoredSession(documentId, {
          sessionId: nextSnapshot.sessionId,
          threadId: nextSnapshot.threadId,
        });
        sessionRef.current = {
          sessionId: nextSnapshot.sessionId,
          threadId: nextSnapshot.threadId,
        };
        connectStream(nextSnapshot.sessionId);
        setError(null);
      } catch (nextError) {
        if (!cancelled) {
          setError(nextError instanceof Error ? nextError.message : String(nextError));
        }
      } finally {
        if (!cancelled) {
          setIsLoading(false);
        }
      }
    };
    void bootstrap();
    return () => {
      cancelled = true;
      closeStream();
    };
  }, [closeStream, connectStream, documentId, enabled, contextKey]);

  useEffect(() => {
    if (!enabled || !snapshot) {
      return;
    }
    const nextContextKey = JSON.stringify(currentContext);
    if (lastContextKeyRef.current === nextContextKey) {
      return;
    }
    lastContextKeyRef.current = nextContextKey;
    const timeout = window.setTimeout(() => {
      const activeSession = sessionRef.current;
      if (!activeSession) {
        return;
      }
      void fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/agent/sessions/${encodeURIComponent(activeSession.sessionId)}/context`,
        {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            context: currentContext,
          }),
        },
      ).catch(() => undefined);
    }, 150);
    return () => {
      window.clearTimeout(timeout);
    };
  }, [currentContext, documentId, enabled, snapshot]);

  const sendPrompt = useCallback(async () => {
    const prompt = draft.trim();
    const activeSession = sessionRef.current;
    if (!activeSession || prompt.length === 0) {
      return;
    }
    try {
      setError(null);
      const response = await fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/agent/sessions/${encodeURIComponent(activeSession.sessionId)}/turns`,
        {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            prompt,
            context: getContextRef.current(),
          }),
        },
      );
      const payload = (await response.json()) as unknown;
      if (!response.ok) {
        throw new Error(
          resolvePayloadMessage(
            payload,
            `Workbook agent request failed with status ${response.status}`,
          ),
        );
      }
      setSnapshot(decodeUnknownSync(WorkbookAgentSessionSnapshotSchema, payload));
      setDraft("");
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : String(nextError));
    }
  }, [documentId, draft]);

  const interrupt = useCallback(async () => {
    const activeSession = sessionRef.current;
    if (!activeSession) {
      return;
    }
    try {
      const response = await fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/agent/sessions/${encodeURIComponent(activeSession.sessionId)}/interrupt`,
        {
          method: "POST",
        },
      );
      const payload = (await response.json()) as unknown;
      if (!response.ok) {
        throw new Error(
          resolvePayloadMessage(
            payload,
            `Workbook agent request failed with status ${response.status}`,
          ),
        );
      }
      setSnapshot(decodeUnknownSync(WorkbookAgentSessionSnapshotSchema, payload));
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : String(nextError));
    }
  }, [documentId]);

  const agentToggle = useMemo(
    () => (
      <button
        aria-controls="workbook-agent-panel"
        aria-expanded={isOpen}
        aria-label="Toggle workbook assistant"
        className="inline-flex h-8 items-center gap-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
        data-testid="workbook-agent-toggle"
        type="button"
        onClick={() => {
          setIsOpen((current) => !current);
        }}
      >
        <span>Assistant</span>
      </button>
    ),
    [isOpen],
  );

  const agentPanel = useMemo(
    () => (
      <WorkbookAgentPanel
        draft={draft}
        error={error}
        isLoading={isLoading}
        isOpen={isOpen}
        snapshot={snapshot}
        onClose={() => {
          setIsOpen(false);
        }}
        onDraftChange={setDraft}
        onInterrupt={() => {
          void interrupt();
        }}
        onSubmit={() => {
          void sendPrompt();
        }}
      />
    ),
    [draft, error, interrupt, isLoading, isOpen, sendPrompt, snapshot],
  );

  return {
    agentPanel,
    agentToggle,
  };
}
