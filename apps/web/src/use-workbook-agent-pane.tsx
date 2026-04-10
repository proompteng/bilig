import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  isFullWorkbookAgentCommandSelection,
  isWorkbookAgentCommandBundle,
  isWorkbookAgentExecutionRecord,
  normalizeWorkbookAgentCommandIndexes,
  projectWorkbookAgentBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentExecutionRecord,
  type WorkbookAgentPreviewSummary,
} from "@bilig/agent-api";
import {
  WorkbookAgentSessionSnapshotSchema,
  WorkbookAgentStreamEventSchema,
  WorkbookAgentThreadScopeSchema,
  WorkbookAgentThreadSummarySchema,
  decodeUnknownSync,
  type WorkbookAgentSessionSnapshot,
  type WorkbookAgentStreamEvent,
  type WorkbookAgentThreadScope,
  type WorkbookAgentThreadSummary,
  type WorkbookAgentUiContext,
} from "@bilig/contracts";
import { Schema } from "effect";
import { WorkbookAgentPanel } from "./WorkbookAgentPanel.js";

const STORAGE_KEY_PREFIX = "bilig:workbook-agent:";
const DRAFT_STORAGE_KEY_PREFIX = "bilig:workbook-agent-drafts:";
const WorkbookAgentThreadSummaryListSchema = Schema.Array(WorkbookAgentThreadSummarySchema);

interface StoredWorkbookAgentThreadRef {
  threadId: string;
  sessionId?: string;
}

interface WorkbookAgentLiveSession {
  sessionId: string;
  threadId: string;
}

function storageKey(documentId: string): string {
  return `${STORAGE_KEY_PREFIX}${documentId}`;
}

function draftStorageKey(documentId: string): string {
  return `${DRAFT_STORAGE_KEY_PREFIX}${documentId}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isStoredWorkbookAgentSession(value: unknown): value is StoredWorkbookAgentThreadRef {
  return (
    isRecord(value) &&
    (value["sessionId"] === undefined || typeof value["sessionId"] === "string") &&
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

function loadStoredSession(documentId: string): StoredWorkbookAgentThreadRef | null {
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

function persistStoredSession(documentId: string, value: StoredWorkbookAgentThreadRef): void {
  window.sessionStorage.setItem(storageKey(documentId), JSON.stringify(value));
}

function clearStoredSession(documentId: string): void {
  window.sessionStorage.removeItem(storageKey(documentId));
}

function loadStoredDrafts(documentId: string): Record<string, string> {
  try {
    const raw = window.sessionStorage.getItem(draftStorageKey(documentId));
    if (!raw) {
      return {};
    }
    const parsed = JSON.parse(raw) as unknown;
    if (!isRecord(parsed)) {
      return {};
    }
    return Object.fromEntries(
      Object.entries(parsed).flatMap(([key, value]) =>
        typeof value === "string" ? ([[key, value]] as const) : [],
      ),
    );
  } catch {
    return {};
  }
}

function persistStoredDrafts(documentId: string, drafts: Record<string, string>): void {
  const entries = Object.entries(drafts).filter((entry) => entry[1].length > 0);
  if (entries.length === 0) {
    window.sessionStorage.removeItem(draftStorageKey(documentId));
    return;
  }
  window.sessionStorage.setItem(
    draftStorageKey(documentId),
    JSON.stringify(Object.fromEntries(entries)),
  );
}

function clearStoredDraft(documentId: string, key: string): void {
  const drafts = loadStoredDrafts(documentId);
  if (!(key in drafts)) {
    return;
  }
  delete drafts[key];
  persistStoredDrafts(documentId, drafts);
}

function draftKey(threadId: string | null, scope: WorkbookAgentThreadScope): string {
  return threadId ? `thread:${threadId}` : `new:${scope}`;
}

function createSessionResumeBody(
  storedSession: StoredWorkbookAgentThreadRef | null,
  context: WorkbookAgentUiContext,
  scope: WorkbookAgentThreadScope,
): {
  readonly threadId?: string;
  readonly context: WorkbookAgentUiContext;
  readonly scope?: WorkbookAgentThreadScope;
} {
  return storedSession?.threadId
    ? {
        threadId: storedSession.threadId,
        context,
      }
    : {
        context,
        scope: decodeUnknownSync(WorkbookAgentThreadScopeSchema, scope),
      };
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

function normalizeWorkbookAgentErrorMessage(error: string): string {
  if (error.includes("thread/start.dynamicTools requires experimentalApi capability")) {
    return "Retry in a moment.";
  }
  if (error.includes("Invalid Codex initialize response")) {
    return "Retry in a moment.";
  }
  return error;
}

export function useWorkbookAgentPane(input: {
  readonly documentId: string;
  readonly enabled: boolean;
  readonly getContext: () => WorkbookAgentUiContext;
  readonly previewBundle: (
    bundle: WorkbookAgentCommandBundle,
  ) => Promise<WorkbookAgentPreviewSummary>;
}) {
  const { documentId, enabled, getContext, previewBundle } = input;
  const [snapshot, setSnapshot] = useState<WorkbookAgentSessionSnapshot | null>(null);
  const [draft, setDraft] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isApplyingBundle, setIsApplyingBundle] = useState(false);
  const [preview, setPreview] = useState<WorkbookAgentPreviewSummary | null>(null);
  const [selectedCommandIndexes, setSelectedCommandIndexes] = useState<number[]>([]);
  const [threadSummaries, setThreadSummaries] = useState<readonly WorkbookAgentThreadSummary[]>([]);
  const [threadScope, setThreadScope] = useState<WorkbookAgentThreadScope>("private");
  const autoApplyBundleIdRef = useRef<string | null>(null);
  const sessionRef = useRef<WorkbookAgentLiveSession | null>(null);
  const eventSourceRef = useRef<EventSource | null>(null);
  const recoveringStreamRef = useRef(false);
  const lastContextKeyRef = useRef<string>("");
  const lastDraftKeyRef = useRef<string | null>(null);
  const getContextRef = useRef(getContext);
  const currentContext = getContextRef.current();
  const activeDraftKey = draftKey(snapshot?.threadId ?? null, threadScope);

  useEffect(() => {
    getContextRef.current = getContext;
  }, [getContext]);

  useEffect(() => {
    if (lastDraftKeyRef.current === activeDraftKey) {
      return;
    }
    lastDraftKeyRef.current = activeDraftKey;
    setDraft(loadStoredDrafts(documentId)[activeDraftKey] ?? "");
  }, [activeDraftKey, documentId]);

  useEffect(() => {
    const drafts = loadStoredDrafts(documentId);
    if (draft.length === 0) {
      if (!(activeDraftKey in drafts)) {
        return;
      }
      delete drafts[activeDraftKey];
    } else {
      drafts[activeDraftKey] = draft;
    }
    persistStoredDrafts(documentId, drafts);
  }, [activeDraftKey, documentId, draft]);

  const loadThreadSummaries = useCallback(async (): Promise<
    readonly WorkbookAgentThreadSummary[]
  > => {
    const response = await fetch(`/v2/documents/${encodeURIComponent(documentId)}/agent/threads`);
    const payload = (await response.json()) as unknown;
    if (!response.ok) {
      throw new Error(
        resolvePayloadMessage(
          payload,
          `Workbook agent request failed with status ${response.status}`,
        ),
      );
    }
    return decodeUnknownSync(WorkbookAgentThreadSummaryListSchema, payload);
  }, [documentId]);

  const pendingBundle = useMemo<WorkbookAgentCommandBundle | null>(() => {
    const candidate = snapshot?.pendingBundle;
    return candidate && isWorkbookAgentCommandBundle(candidate) ? candidate : null;
  }, [snapshot?.pendingBundle]);
  const pendingCommandCount = pendingBundle?.commands.length ?? 0;

  const normalizedCommandIndexes = useMemo(
    () =>
      pendingBundle
        ? normalizeWorkbookAgentCommandIndexes(pendingBundle, selectedCommandIndexes)
        : [],
    [pendingBundle, selectedCommandIndexes],
  );

  const selectedPendingBundle = useMemo<WorkbookAgentCommandBundle | null>(
    () =>
      pendingBundle
        ? projectWorkbookAgentBundle({
            bundle: pendingBundle,
            commandIndexes: normalizedCommandIndexes,
            bundleId: pendingBundle.id,
          })
        : null,
    [normalizedCommandIndexes, pendingBundle],
  );

  const executionRecords = useMemo<WorkbookAgentExecutionRecord[]>(() => {
    const candidates = snapshot?.executionRecords ?? [];
    return candidates.flatMap((entry) => (isWorkbookAgentExecutionRecord(entry) ? [entry] : []));
  }, [snapshot?.executionRecords]);

  const closeStream = useCallback(() => {
    eventSourceRef.current?.close();
    eventSourceRef.current = null;
  }, []);

  const persistSessionSnapshot = useCallback(
    (nextSnapshot: WorkbookAgentSessionSnapshot) => {
      setSnapshot(nextSnapshot);
      setThreadScope(nextSnapshot.scope);
      persistStoredSession(documentId, {
        sessionId: nextSnapshot.sessionId,
        threadId: nextSnapshot.threadId,
      });
      sessionRef.current = {
        sessionId: nextSnapshot.sessionId,
        threadId: nextSnapshot.threadId,
      };
      void loadThreadSummaries()
        .then((nextThreadSummaries) => {
          setThreadSummaries(nextThreadSummaries);
          return nextThreadSummaries;
        })
        .catch(() => undefined);
    },
    [documentId, loadThreadSummaries],
  );

  const connectStream = useCallback(
    (threadId: string) => {
      closeStream();
      const source = new EventSource(
        `/v2/documents/${encodeURIComponent(documentId)}/agent/threads/${encodeURIComponent(threadId)}/events`,
      );
      source.addEventListener("message", (message) => {
        try {
          const payloadText = readMessageEventData(message);
          if (payloadText === null) {
            return;
          }
          const event = decodeUnknownSync(WorkbookAgentStreamEventSchema, JSON.parse(payloadText));
          if (event.type === "snapshot") {
            recoveringStreamRef.current = false;
            persistSessionSnapshot(event.snapshot);
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
        if (eventSourceRef.current === source) {
          source.close();
          eventSourceRef.current = null;
        }
        if (recoveringStreamRef.current) {
          return;
        }
        const storedSession = sessionRef.current;
        if (!storedSession) {
          setError("Assistant stream disconnected.");
          return;
        }
        recoveringStreamRef.current = true;
        setError(null);
        void (async () => {
          try {
            setIsLoading(true);
            const response = await fetch(
              `/v2/documents/${encodeURIComponent(documentId)}/agent/sessions`,
              {
                method: "POST",
                headers: {
                  "content-type": "application/json",
                },
                body: JSON.stringify(
                  createSessionResumeBody(storedSession, getContextRef.current(), threadScope),
                ),
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
            persistSessionSnapshot(nextSnapshot);
            connectStream(nextSnapshot.threadId);
          } catch (nextError) {
            recoveringStreamRef.current = false;
            setError(nextError instanceof Error ? nextError.message : String(nextError));
          } finally {
            setIsLoading(false);
          }
        })();
      });
      eventSourceRef.current = source;
    },
    [closeStream, documentId, persistSessionSnapshot, threadScope],
  );

  const createOrResumeSession = useCallback(
    async (
      storedSession: StoredWorkbookAgentThreadRef | null,
      context: WorkbookAgentUiContext,
      scope: WorkbookAgentThreadScope,
    ) => {
      const response = await fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/agent/sessions`,
        {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(createSessionResumeBody(storedSession, context, scope)),
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
      return decodeUnknownSync(WorkbookAgentSessionSnapshotSchema, payload);
    },
    [documentId],
  );

  const ensureSession = useCallback(async (): Promise<WorkbookAgentLiveSession> => {
    const activeSession = sessionRef.current;
    if (activeSession) {
      return activeSession;
    }
    setIsLoading(true);
    try {
      const nextSnapshot = await createOrResumeSession(null, getContextRef.current(), threadScope);
      persistSessionSnapshot(nextSnapshot);
      connectStream(nextSnapshot.threadId);
      setError(null);
      const nextSession = {
        sessionId: nextSnapshot.sessionId,
        threadId: nextSnapshot.threadId,
      };
      sessionRef.current = nextSession;
      return nextSession;
    } finally {
      setIsLoading(false);
    }
  }, [connectStream, createOrResumeSession, persistSessionSnapshot, threadScope]);

  useEffect(() => {
    autoApplyBundleIdRef.current = null;
  }, [pendingBundle?.id]);

  useEffect(() => {
    setSelectedCommandIndexes(
      Array.from({ length: pendingCommandCount }, (_unused, index) => index),
    );
  }, [pendingBundle?.id, pendingCommandCount]);

  useEffect(() => {
    if (!enabled || selectedPendingBundle === null) {
      setPreview(null);
      return;
    }
    let cancelled = false;
    void (async () => {
      try {
        const nextPreview = await previewBundle(selectedPendingBundle);
        if (!cancelled) {
          setPreview(nextPreview);
        }
      } catch (nextError: unknown) {
        if (!cancelled) {
          setPreview(null);
          setError(nextError instanceof Error ? nextError.message : String(nextError));
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [enabled, previewBundle, selectedPendingBundle]);

  const applyPendingBundle = useCallback(
    async (appliedBy: "user" | "auto" = "user") => {
      const activeSession = sessionRef.current;
      if (!activeSession || !pendingBundle || !selectedPendingBundle || !preview) {
        return;
      }
      try {
        setIsApplyingBundle(true);
        const response = await fetch(
          `/v2/documents/${encodeURIComponent(documentId)}/agent/threads/${encodeURIComponent(activeSession.threadId)}/bundles/${encodeURIComponent(pendingBundle.id)}/apply`,
          {
            method: "POST",
            headers: {
              "content-type": "application/json",
            },
            body: JSON.stringify({
              appliedBy,
              commandIndexes: normalizedCommandIndexes,
              preview,
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
        persistSessionSnapshot(decodeUnknownSync(WorkbookAgentSessionSnapshotSchema, payload));
        setError(null);
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError));
      } finally {
        setIsApplyingBundle(false);
      }
    },
    [
      documentId,
      normalizedCommandIndexes,
      pendingBundle,
      persistSessionSnapshot,
      preview,
      selectedPendingBundle,
    ],
  );

  useEffect(() => {
    if (
      !enabled ||
      !pendingBundle ||
      !selectedPendingBundle ||
      !preview ||
      pendingBundle.approvalMode !== "auto" ||
      !isFullWorkbookAgentCommandSelection({
        bundle: pendingBundle,
        commandIndexes: normalizedCommandIndexes,
      }) ||
      isApplyingBundle ||
      autoApplyBundleIdRef.current === pendingBundle.id
    ) {
      return;
    }
    autoApplyBundleIdRef.current = pendingBundle.id;
    void applyPendingBundle("auto");
  }, [
    applyPendingBundle,
    enabled,
    isApplyingBundle,
    normalizedCommandIndexes,
    pendingBundle,
    preview,
    selectedPendingBundle,
  ]);

  const togglePendingCommand = useCallback(
    (commandIndex: number) => {
      setSelectedCommandIndexes((current) => {
        if (!pendingBundle || commandIndex < 0 || commandIndex >= pendingBundle.commands.length) {
          return current;
        }
        const selected = new Set(normalizeWorkbookAgentCommandIndexes(pendingBundle, current));
        if (selected.has(commandIndex)) {
          selected.delete(commandIndex);
        } else {
          selected.add(commandIndex);
        }
        return pendingBundle.commands.flatMap((_command, index) =>
          selected.has(index) ? [index] : [],
        );
      });
    },
    [pendingBundle],
  );

  const selectAllPendingCommands = useCallback(() => {
    setSelectedCommandIndexes(
      pendingBundle ? pendingBundle.commands.map((_command, index) => index) : [],
    );
  }, [pendingBundle]);

  useEffect(() => {
    if (!enabled) {
      closeStream();
      sessionRef.current = null;
      recoveringStreamRef.current = false;
      setThreadSummaries([]);
      setSnapshot(null);
      setIsLoading(false);
      return;
    }
    let cancelled = false;
    lastContextKeyRef.current = "";
    const storedSession = loadStoredSession(documentId);
    sessionRef.current = null;

    const bootstrapThreadSummaries = async () => {
      try {
        const nextThreadSummaries = await loadThreadSummaries();
        if (!cancelled) {
          setThreadSummaries(nextThreadSummaries);
        }
      } catch {
        if (!cancelled) {
          setThreadSummaries([]);
        }
      }
    };
    void bootstrapThreadSummaries();

    if (!storedSession) {
      setIsLoading(false);
      return () => {
        cancelled = true;
        closeStream();
      };
    }

    const bootstrapStoredSession = async () => {
      try {
        setIsLoading(true);
        const nextSnapshot = await createOrResumeSession(
          storedSession,
          getContextRef.current(),
          threadScope,
        );
        if (cancelled) {
          return;
        }
        persistSessionSnapshot(nextSnapshot);
        connectStream(nextSnapshot.threadId);
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
    void bootstrapStoredSession();
    return () => {
      cancelled = true;
      closeStream();
    };
  }, [
    closeStream,
    connectStream,
    createOrResumeSession,
    documentId,
    enabled,
    loadThreadSummaries,
    persistSessionSnapshot,
    threadScope,
  ]);

  const selectThread = useCallback(
    async (threadId: string) => {
      if (sessionRef.current?.threadId === threadId) {
        return;
      }
      try {
        setIsLoading(true);
        setError(null);
        const nextSnapshot = await createOrResumeSession(
          { threadId },
          getContextRef.current(),
          threadScope,
        );
        persistSessionSnapshot(nextSnapshot);
        connectStream(nextSnapshot.threadId);
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError));
      } finally {
        setIsLoading(false);
      }
    },
    [connectStream, createOrResumeSession, persistSessionSnapshot, threadScope],
  );

  const startNewThread = useCallback(() => {
    closeStream();
    clearStoredSession(documentId);
    recoveringStreamRef.current = false;
    sessionRef.current = null;
    setSnapshot(null);
    setPreview(null);
    setSelectedCommandIndexes([]);
    setError(null);
  }, [closeStream, documentId]);

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
        `/v2/documents/${encodeURIComponent(documentId)}/agent/threads/${encodeURIComponent(activeSession.threadId)}/context`,
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
    if (prompt.length === 0) {
      return;
    }
    try {
      setError(null);
      const existingSession = sessionRef.current;
      const activeSession = existingSession ?? (await ensureSession());
      const response = await fetch(
        existingSession
          ? `/v2/documents/${encodeURIComponent(documentId)}/agent/threads/${encodeURIComponent(activeSession.threadId)}/turns`
          : `/v2/documents/${encodeURIComponent(documentId)}/agent/sessions/${encodeURIComponent(activeSession.sessionId)}/turns`,
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
      persistSessionSnapshot(decodeUnknownSync(WorkbookAgentSessionSnapshotSchema, payload));
      clearStoredDraft(documentId, activeDraftKey);
      setDraft("");
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : String(nextError));
    }
  }, [activeDraftKey, documentId, draft, ensureSession, persistSessionSnapshot]);

  const interrupt = useCallback(async () => {
    const activeSession = sessionRef.current;
    if (!activeSession) {
      return;
    }
    try {
      const response = await fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/agent/threads/${encodeURIComponent(activeSession.threadId)}/interrupt`,
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

  const dismissPendingBundle = useCallback(async () => {
    const activeSession = sessionRef.current;
    if (!activeSession || !pendingBundle) {
      return;
    }
    try {
      const response = await fetch(
        `/v2/documents/${encodeURIComponent(documentId)}/agent/threads/${encodeURIComponent(activeSession.threadId)}/bundles/${encodeURIComponent(pendingBundle.id)}/dismiss`,
        {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: "{}",
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
      persistSessionSnapshot(decodeUnknownSync(WorkbookAgentSessionSnapshotSchema, payload));
      setError(null);
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : String(nextError));
    }
  }, [documentId, pendingBundle, persistSessionSnapshot]);

  const replayExecutionRecord = useCallback(
    async (recordId: string) => {
      const activeSession = await ensureSession();
      try {
        const response = await fetch(
          `/v2/documents/${encodeURIComponent(documentId)}/agent/threads/${encodeURIComponent(activeSession.threadId)}/runs/${encodeURIComponent(recordId)}/replay`,
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
        persistSessionSnapshot(decodeUnknownSync(WorkbookAgentSessionSnapshotSchema, payload));
        setError(null);
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError));
      }
    },
    [documentId, ensureSession, persistSessionSnapshot],
  );

  const clearAgentError = useCallback(() => {
    setError(null);
  }, []);

  const agentPanel = useMemo(
    () => (
      <WorkbookAgentPanel
        activeThreadId={snapshot?.threadId ?? sessionRef.current?.threadId ?? null}
        currentContext={currentContext}
        draft={draft}
        executionRecords={executionRecords}
        isApplyingBundle={isApplyingBundle}
        isLoading={isLoading}
        pendingBundle={pendingBundle}
        preview={preview}
        selectedCommandIndexes={normalizedCommandIndexes}
        snapshot={snapshot}
        threadScope={threadScope}
        threadSummaries={threadSummaries}
        onApplyPendingBundle={() => {
          void applyPendingBundle("user");
        }}
        onDraftChange={setDraft}
        onDismissPendingBundle={() => {
          void dismissPendingBundle();
        }}
        onInterrupt={() => {
          void interrupt();
        }}
        onSelectAllPendingCommands={selectAllPendingCommands}
        onSelectThreadScope={setThreadScope}
        onTogglePendingCommand={togglePendingCommand}
        onReplayExecutionRecord={(recordId) => {
          void replayExecutionRecord(recordId);
        }}
        onSelectThread={(threadId) => {
          void selectThread(threadId);
        }}
        onStartNewThread={startNewThread}
        onSubmit={() => {
          void sendPrompt();
        }}
      />
    ),
    [
      applyPendingBundle,
      currentContext,
      dismissPendingBundle,
      draft,
      executionRecords,
      interrupt,
      isApplyingBundle,
      isLoading,
      normalizedCommandIndexes,
      pendingBundle,
      preview,
      replayExecutionRecord,
      sendPrompt,
      selectThread,
      snapshot,
      selectAllPendingCommands,
      setThreadScope,
      startNewThread,
      threadScope,
      threadSummaries,
      togglePendingCommand,
    ],
  );

  return {
    agentPanel,
    agentError: error ? normalizeWorkbookAgentErrorMessage(error) : null,
    clearAgentError,
    pendingCommandCount,
    previewRanges: preview?.ranges ?? pendingBundle?.affectedRanges ?? [],
  };
}
