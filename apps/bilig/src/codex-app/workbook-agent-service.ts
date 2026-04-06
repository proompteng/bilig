import type {
  WorkbookAgentSessionSnapshot,
  WorkbookAgentStreamEvent,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
} from "@bilig/contracts";
import type {
  CodexServerNotification,
  WorkbookAgentAppliedBy,
  WorkbookAgentCommandBundle,
  WorkbookAgentContextRef,
  WorkbookAgentExecutionRecord,
} from "@bilig/agent-api";
import {
  appendWorkbookAgentCommandToBundle,
  buildWorkbookAgentExecutionRecord,
  createWorkbookAgentCommandBundle,
  decodeWorkbookAgentPreviewSummary,
  describeWorkbookAgentBundle,
  splitWorkbookAgentCommandBundle,
} from "@bilig/agent-api";
import type { SessionIdentity } from "../http/session.js";
import type { ZeroSyncService } from "../zero/service.js";
import { createWorkbookAgentServiceError } from "../workbook-agent-errors.js";
import {
  CodexAppServerClient,
  type CodexAppServerTransport,
  type CodexAppServerClientOptions,
} from "./codex-app-server-client.js";
import {
  handleWorkbookAgentToolCall,
  workbookAgentDynamicToolSpecs,
} from "./workbook-agent-tools.js";
import {
  buildEntriesFromThread,
  cloneSnapshot,
  createSessionBodySchema,
  createSystemEntry,
  createWorkbookAgentBaseInstructions,
  createWorkbookAgentDeveloperInstructions,
  mapThreadItemToEntry,
  startTurnBodySchema,
  updateContextBodySchema,
} from "./workbook-agent-session-model.js";

const DEFAULT_MODEL = process.env["BILIG_CODEX_MODEL"]?.trim() || "gpt-5.4";

function upsertEntry(
  entries: readonly WorkbookAgentTimelineEntry[],
  nextEntry: WorkbookAgentTimelineEntry,
): WorkbookAgentTimelineEntry[] {
  const index = entries.findIndex((entry) => entry.id === nextEntry.id);
  if (index < 0) {
    return [...entries, nextEntry];
  }
  const nextEntries = [...entries];
  nextEntries[index] = nextEntry;
  return nextEntries;
}

function removeEntry(
  entries: readonly WorkbookAgentTimelineEntry[],
  entryId: string,
): WorkbookAgentTimelineEntry[] {
  return entries.filter((entry) => entry.id !== entryId);
}

type MutableWorkbookAgentSessionSnapshot = {
  -readonly [Key in Exclude<
    keyof WorkbookAgentSessionSnapshot,
    "entries" | "pendingBundle" | "executionRecords"
  >]: WorkbookAgentSessionSnapshot[Key];
} & {
  entries: WorkbookAgentTimelineEntry[];
  pendingBundle: WorkbookAgentCommandBundle | null;
  executionRecords: WorkbookAgentExecutionRecord[];
};

interface WorkbookAgentSessionState {
  readonly sessionId: string;
  readonly documentId: string;
  readonly userId: string;
  threadId: string;
  snapshot: MutableWorkbookAgentSessionSnapshot;
  optimisticUserEntryIdByTurn: Map<string, string>;
  promptByTurn: Map<string, string>;
  lastAccessedAt: number;
}

function toContextRef(context: WorkbookAgentUiContext | null): WorkbookAgentContextRef | null {
  return context
    ? {
        selection: {
          sheetName: context.selection.sheetName,
          address: context.selection.address,
        },
        viewport: { ...context.viewport },
      }
    : null;
}

export interface WorkbookAgentService {
  readonly enabled: boolean;
  createSession(input: {
    documentId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot>;
  updateContext(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot>;
  startTurn(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot>;
  interruptTurn(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentSessionSnapshot>;
  applyPendingBundle(input: {
    documentId: string;
    sessionId: string;
    bundleId: string;
    session: SessionIdentity;
    appliedBy: WorkbookAgentAppliedBy;
    commandIndexes?: readonly number[] | null;
    preview: unknown;
  }): Promise<WorkbookAgentSessionSnapshot>;
  dismissPendingBundle(input: {
    documentId: string;
    sessionId: string;
    bundleId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentSessionSnapshot>;
  replayExecutionRecord(input: {
    documentId: string;
    sessionId: string;
    recordId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentSessionSnapshot>;
  getSnapshot(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
  }): WorkbookAgentSessionSnapshot;
  subscribe(sessionId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void;
  close(): Promise<void>;
}

class DisabledWorkbookAgentService implements WorkbookAgentService {
  readonly enabled = false;

  async createSession(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async updateContext(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async startTurn(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async interruptTurn(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async applyPendingBundle(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async dismissPendingBundle(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async replayExecutionRecord(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  getSnapshot(): never {
    throw new Error("Workbook agent service is not configured");
  }

  subscribe(): () => void {
    return () => {};
  }

  async close(): Promise<void> {}
}

export interface EnabledWorkbookAgentServiceOptions {
  zeroSyncService: ZeroSyncService;
  codexClientFactory?: (options: CodexAppServerClientOptions) => CodexAppServerTransport;
  now?: () => number;
  maxSessions?: number;
}

class EnabledWorkbookAgentService implements WorkbookAgentService {
  readonly enabled = true;
  private readonly zeroSyncService: ZeroSyncService;
  private readonly codexClientFactory: (
    options: CodexAppServerClientOptions,
  ) => CodexAppServerTransport;
  private readonly now: () => number;
  private readonly maxSessions: number;
  private readonly sessions = new Map<string, WorkbookAgentSessionState>();
  private readonly threadToSessionId = new Map<string, string>();
  private readonly subscribers = new Map<string, Set<(event: WorkbookAgentStreamEvent) => void>>();
  private codexClient: CodexAppServerTransport | null = null;
  private unsubscribeCodex: (() => void) | null = null;

  constructor(options: EnabledWorkbookAgentServiceOptions) {
    this.zeroSyncService = options.zeroSyncService;
    this.codexClientFactory =
      options.codexClientFactory ?? ((clientOptions) => new CodexAppServerClient(clientOptions));
    this.now = options.now ?? (() => Date.now());
    this.maxSessions = options.maxSessions ?? 64;
  }

  async createSession(input: {
    documentId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const parsed = createSessionBodySchema.parse(input.body);
    const sessionId = parsed.sessionId ?? crypto.randomUUID();
    const existing = this.sessions.get(sessionId);
    if (existing) {
      const sessionState = this.requireOwnedSession(
        existing,
        input.documentId,
        input.session.userID,
      );
      if (parsed.context) {
        sessionState.snapshot.context = parsed.context;
        this.emitSnapshot(sessionId);
      }
      this.touch(sessionState);
      return cloneSnapshot(sessionState.snapshot);
    }

    const codexClient = await this.getCodexClient();
    const executionRecords = await this.zeroSyncService.listWorkbookAgentRuns(
      input.documentId,
      input.session.userID,
    );
    const thread =
      parsed.threadId === undefined
        ? await codexClient.threadStart({
            model: DEFAULT_MODEL,
            approvalPolicy: "never",
            sandbox: "read-only",
            baseInstructions: createWorkbookAgentBaseInstructions(),
            developerInstructions: createWorkbookAgentDeveloperInstructions(),
            dynamicTools: workbookAgentDynamicToolSpecs,
          })
        : await codexClient.threadResume({
            threadId: parsed.threadId,
            baseInstructions: createWorkbookAgentBaseInstructions(),
            developerInstructions: createWorkbookAgentDeveloperInstructions(),
          });

    const snapshot: MutableWorkbookAgentSessionSnapshot = {
      sessionId,
      documentId: input.documentId,
      threadId: thread.id,
      status: thread.turns.some((turn) => turn.status === "failed")
        ? "failed"
        : thread.turns.some((turn) => turn.status === "inProgress")
          ? "inProgress"
          : "idle",
      activeTurnId: thread.turns.findLast((turn) => turn.status === "inProgress")?.id ?? null,
      lastError: thread.turns.findLast((turn) => turn.error?.message)?.error?.message ?? null,
      context: parsed.context ?? null,
      entries: buildEntriesFromThread(thread),
      pendingBundle: null,
      executionRecords,
    };
    const sessionState: WorkbookAgentSessionState = {
      sessionId,
      documentId: input.documentId,
      userId: input.session.userID,
      threadId: thread.id,
      snapshot,
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      lastAccessedAt: this.now(),
    };
    this.sessions.set(sessionId, sessionState);
    this.threadToSessionId.set(thread.id, sessionId);
    this.evictIfNeeded();
    return cloneSnapshot(snapshot);
  }

  async updateContext(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const parsed = updateContextBodySchema.parse(input.body);
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    sessionState.snapshot.context = parsed.context;
    this.touch(sessionState);
    this.emitSnapshot(input.sessionId);
    return cloneSnapshot(sessionState.snapshot);
  }

  async startTurn(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const parsed = startTurnBodySchema.parse(input.body);
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    if (parsed.context) {
      sessionState.snapshot.context = parsed.context;
    }
    const codexClient = await this.getCodexClient();
    const turn = await codexClient.turnStart({
      threadId: sessionState.threadId,
      prompt: parsed.prompt,
    });
    const optimisticEntryId = `optimistic-user:${turn.id}`;
    sessionState.snapshot.entries = upsertEntry(sessionState.snapshot.entries, {
      id: optimisticEntryId,
      kind: "user",
      turnId: turn.id,
      text: parsed.prompt,
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
    });
    sessionState.optimisticUserEntryIdByTurn.set(turn.id, optimisticEntryId);
    sessionState.promptByTurn.set(turn.id, parsed.prompt);
    sessionState.snapshot.activeTurnId = turn.id;
    sessionState.snapshot.status = "inProgress";
    sessionState.snapshot.lastError = null;
    this.touch(sessionState);
    this.emitSnapshot(input.sessionId);
    return cloneSnapshot(sessionState.snapshot);
  }

  async interruptTurn(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    const codexClient = await this.getCodexClient();
    await codexClient.turnInterrupt(sessionState.threadId);
    return cloneSnapshot(sessionState.snapshot);
  }

  async applyPendingBundle(input: {
    documentId: string;
    sessionId: string;
    bundleId: string;
    session: SessionIdentity;
    appliedBy: WorkbookAgentAppliedBy;
    commandIndexes?: readonly number[] | null;
    preview: unknown;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    const pendingBundle = sessionState.snapshot.pendingBundle;
    if (!pendingBundle || pendingBundle.id !== input.bundleId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_BUNDLE_NOT_FOUND",
        message: "Workbook agent preview bundle not found",
        statusCode: 404,
        retryable: false,
      });
    }
    const preview = decodeWorkbookAgentPreviewSummary(input.preview);
    if (!preview) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_PREVIEW_REQUIRED",
        message: "Workbook agent preview summary is required before apply",
        statusCode: 400,
        retryable: false,
      });
    }
    if (input.appliedBy === "auto" && pendingBundle.approvalMode !== "auto") {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_MANUAL_APPROVAL_REQUIRED",
        message: "Workbook agent bundle requires manual approval",
        statusCode: 409,
        retryable: false,
      });
    }
    const selection = splitWorkbookAgentCommandBundle({
      bundle: pendingBundle,
      acceptedCommandIndexes: input.commandIndexes,
    });
    if (!selection.acceptedBundle || !selection.acceptedScope) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_COMMAND_SELECTION_REQUIRED",
        message: "Select at least one staged workbook change before apply",
        statusCode: 400,
        retryable: false,
      });
    }
    if (input.appliedBy === "auto" && selection.acceptedScope !== "full") {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_MANUAL_APPROVAL_REQUIRED",
        message: "Partial workbook agent apply requires manual approval",
        statusCode: 409,
        retryable: false,
      });
    }
    const result = await this.zeroSyncService.applyAgentCommandBundle(
      input.documentId,
      selection.acceptedBundle,
      preview,
      input.session,
    );
    const executionRecord = buildWorkbookAgentExecutionRecord({
      bundle: selection.acceptedBundle,
      actorUserId: input.session.userID,
      planText: this.collectPlanTextForTurn(sessionState, pendingBundle.turnId),
      preview: result.preview,
      appliedRevision: result.revision,
      appliedBy: input.appliedBy,
      acceptedScope: selection.acceptedScope,
      now: this.now(),
    });
    await this.zeroSyncService.appendWorkbookAgentRun(executionRecord);
    sessionState.snapshot.executionRecords = [
      executionRecord,
      ...sessionState.snapshot.executionRecords.filter(
        (record) => record.id !== executionRecord.id,
      ),
    ];
    sessionState.snapshot.pendingBundle =
      selection.remainingBundle === null
        ? null
        : createWorkbookAgentCommandBundle({
            documentId: selection.remainingBundle.documentId,
            threadId: selection.remainingBundle.threadId,
            turnId: selection.remainingBundle.turnId,
            goalText: selection.remainingBundle.goalText,
            baseRevision: result.revision,
            context: selection.remainingBundle.context,
            commands: selection.remainingBundle.commands,
            now: this.now(),
          });
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createSystemEntry(
        `system-apply:${executionRecord.id}`,
        pendingBundle.turnId,
        `${input.appliedBy === "auto" ? "Auto-applied" : "Applied"} ${
          selection.acceptedScope === "partial" ? "selected " : ""
        }preview bundle at revision r${String(result.revision)}: ${selection.acceptedBundle.summary}`,
      ),
    );
    this.touch(sessionState);
    this.emitSnapshot(sessionState.sessionId);
    return cloneSnapshot(sessionState.snapshot);
  }

  async dismissPendingBundle(input: {
    documentId: string;
    sessionId: string;
    bundleId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    const pendingBundle = sessionState.snapshot.pendingBundle;
    if (!pendingBundle || pendingBundle.id !== input.bundleId) {
      throw new Error("Workbook agent preview bundle not found");
    }
    sessionState.snapshot.pendingBundle = null;
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createSystemEntry(
        `system-dismiss:${pendingBundle.id}:${this.now()}`,
        pendingBundle.turnId,
        `Dismissed preview bundle: ${pendingBundle.summary}`,
      ),
    );
    this.touch(sessionState);
    this.emitSnapshot(sessionState.sessionId);
    return cloneSnapshot(sessionState.snapshot);
  }

  async replayExecutionRecord(input: {
    documentId: string;
    sessionId: string;
    recordId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    const record = sessionState.snapshot.executionRecords.find(
      (entry) => entry.id === input.recordId,
    );
    if (!record) {
      throw new Error("Workbook agent execution record not found");
    }
    const baseRevision = await this.zeroSyncService.getWorkbookHeadRevision(input.documentId);
    const replayedBundle = createWorkbookAgentCommandBundle({
      documentId: input.documentId,
      threadId: sessionState.threadId,
      turnId: `replay:${record.id}:${String(this.now())}`,
      goalText: record.goalText,
      baseRevision,
      context: toContextRef(sessionState.snapshot.context) ?? record.context,
      commands: record.commands,
      now: this.now(),
    });
    sessionState.snapshot.pendingBundle = replayedBundle;
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createSystemEntry(
        `system-replay:${record.id}:${String(this.now())}`,
        replayedBundle.turnId,
        `Replayed prior agent plan as preview bundle: ${replayedBundle.summary}`,
      ),
    );
    this.touch(sessionState);
    this.emitSnapshot(sessionState.sessionId);
    return cloneSnapshot(sessionState.snapshot);
  }

  getSnapshot(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
  }): WorkbookAgentSessionSnapshot {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    this.touch(sessionState);
    return cloneSnapshot(sessionState.snapshot);
  }

  subscribe(sessionId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void {
    const listeners = this.subscribers.get(sessionId) ?? new Set();
    listeners.add(listener);
    this.subscribers.set(sessionId, listeners);
    return () => {
      const current = this.subscribers.get(sessionId);
      if (!current) {
        return;
      }
      current.delete(listener);
      if (current.size === 0) {
        this.subscribers.delete(sessionId);
      }
    };
  }

  async close(): Promise<void> {
    this.unsubscribeCodex?.();
    this.unsubscribeCodex = null;
    await this.codexClient?.close();
    this.codexClient = null;
    this.sessions.clear();
    this.threadToSessionId.clear();
    this.subscribers.clear();
  }

  private async getCodexClient(): Promise<CodexAppServerTransport> {
    if (!this.codexClient) {
      this.codexClient = this.codexClientFactory({
        command: process.env["BILIG_CODEX_BIN"]?.trim() || "codex",
        args: ["app-server"],
        cwd: process.cwd(),
        env: process.env,
        onLog: (message) => {
          if (message.length > 0) {
            console.error(message);
          }
        },
        handleDynamicToolCall: (request) => {
          const sessionState = this.getSessionByThreadId(request.threadId);
          return handleWorkbookAgentToolCall(
            {
              documentId: sessionState.documentId,
              session: {
                userID: sessionState.userId,
                roles: ["editor"],
              },
              uiContext: sessionState.snapshot.context,
              zeroSyncService: this.zeroSyncService,
              stageCommand: async (command) => {
                const baseRevision = await this.zeroSyncService.getWorkbookHeadRevision(
                  sessionState.documentId,
                );
                const bundle = appendWorkbookAgentCommandToBundle({
                  previousBundle: sessionState.snapshot.pendingBundle,
                  documentId: sessionState.documentId,
                  threadId: sessionState.threadId,
                  turnId: request.turnId,
                  goalText:
                    sessionState.promptByTurn.get(request.turnId) ??
                    "Update workbook from assistant request",
                  baseRevision,
                  context: toContextRef(sessionState.snapshot.context),
                  command,
                  now: this.now(),
                });
                sessionState.snapshot.pendingBundle = bundle;
                sessionState.snapshot.entries = upsertEntry(
                  sessionState.snapshot.entries,
                  createSystemEntry(
                    `system-preview:${bundle.id}`,
                    request.turnId,
                    describeWorkbookAgentBundle(bundle),
                  ),
                );
                this.emitSnapshot(sessionState.sessionId);
                return bundle;
              },
            },
            request,
          );
        },
      });
      await this.codexClient.ensureReady();
      this.unsubscribeCodex = this.codexClient.subscribe((notification) => {
        this.handleCodexNotification(notification);
      });
    }
    return this.codexClient;
  }

  private handleCodexNotification(notification: CodexServerNotification): void {
    switch (notification.method) {
      case "thread/started":
        return;
      case "turn/started": {
        const sessionState = this.tryGetSessionByThreadId(notification.params.threadId);
        if (!sessionState) {
          return;
        }
        sessionState.snapshot.activeTurnId = notification.params.turn.id;
        sessionState.snapshot.status = "inProgress";
        sessionState.snapshot.lastError = null;
        this.emitSnapshot(sessionState.sessionId);
        return;
      }
      case "turn/completed": {
        const sessionState = this.tryGetSessionByThreadId(notification.params.threadId);
        if (!sessionState) {
          return;
        }
        sessionState.snapshot.activeTurnId = null;
        sessionState.snapshot.status =
          notification.params.turn.status === "failed" ? "failed" : "idle";
        sessionState.snapshot.lastError = notification.params.turn.error?.message ?? null;
        this.emitSnapshot(sessionState.sessionId);
        return;
      }
      case "item/started":
      case "item/completed": {
        const sessionState = this.tryGetSessionByThreadId(notification.params.threadId);
        if (!sessionState) {
          return;
        }
        const optimisticUserEntryId = sessionState.optimisticUserEntryIdByTurn.get(
          notification.params.turnId,
        );
        if (notification.params.item.type === "userMessage" && optimisticUserEntryId) {
          sessionState.snapshot.entries = removeEntry(
            sessionState.snapshot.entries,
            optimisticUserEntryId,
          );
          sessionState.optimisticUserEntryIdByTurn.delete(notification.params.turnId);
        }
        sessionState.snapshot.entries = upsertEntry(
          sessionState.snapshot.entries,
          mapThreadItemToEntry(notification.params.item, notification.params.turnId),
        );
        this.emitSnapshot(sessionState.sessionId);
        return;
      }
      case "item/agentMessage/delta": {
        const sessionState = this.tryGetSessionByThreadId(notification.params.threadId);
        if (!sessionState) {
          return;
        }
        const existing =
          sessionState.snapshot.entries.find((entry) => entry.id === notification.params.itemId) ??
          ({
            id: notification.params.itemId,
            kind: "assistant",
            turnId: notification.params.turnId,
            text: "",
            phase: null,
            toolName: null,
            toolStatus: null,
            argumentsText: null,
            outputText: null,
            success: null,
          } satisfies WorkbookAgentTimelineEntry);
        sessionState.snapshot.entries = upsertEntry(sessionState.snapshot.entries, {
          ...existing,
          text: `${existing.text ?? ""}${notification.params.delta}`,
        });
        this.emit(sessionState.sessionId, {
          type: "assistantDelta",
          itemId: notification.params.itemId,
          delta: notification.params.delta,
        });
        return;
      }
      case "item/plan/delta": {
        const sessionState = this.tryGetSessionByThreadId(notification.params.threadId);
        if (!sessionState) {
          return;
        }
        const existing =
          sessionState.snapshot.entries.find((entry) => entry.id === notification.params.itemId) ??
          ({
            id: notification.params.itemId,
            kind: "plan",
            turnId: notification.params.turnId,
            text: "",
            phase: null,
            toolName: null,
            toolStatus: null,
            argumentsText: null,
            outputText: null,
            success: null,
          } satisfies WorkbookAgentTimelineEntry);
        sessionState.snapshot.entries = upsertEntry(sessionState.snapshot.entries, {
          ...existing,
          text: `${existing.text ?? ""}${notification.params.delta}`,
        });
        this.emit(sessionState.sessionId, {
          type: "planDelta",
          itemId: notification.params.itemId,
          delta: notification.params.delta,
        });
        return;
      }
      case "error": {
        const message =
          typeof notification.params.message === "string"
            ? notification.params.message
            : "Codex app-server error";
        this.sessions.forEach((sessionState) => {
          sessionState.snapshot.lastError = message;
          sessionState.snapshot.status = "failed";
          sessionState.snapshot.entries = upsertEntry(
            sessionState.snapshot.entries,
            createSystemEntry(
              `system-error:${this.now()}`,
              sessionState.snapshot.activeTurnId,
              message,
            ),
          );
          this.emitSnapshot(sessionState.sessionId);
        });
      }
    }
  }

  private collectPlanTextForTurn(
    sessionState: WorkbookAgentSessionState,
    turnId: string,
  ): string | null {
    const planText = sessionState.snapshot.entries
      .filter((entry) => entry.turnId === turnId && entry.kind === "plan" && entry.text)
      .map((entry) => entry.text?.trim() ?? "")
      .filter((text) => text.length > 0)
      .join("\n\n");
    return planText.length > 0 ? planText : null;
  }

  private getOwnedSession(
    documentId: string,
    sessionId: string,
    userId: string,
  ): WorkbookAgentSessionState {
    const sessionState = this.sessions.get(sessionId);
    if (!sessionState) {
      throw new Error("Workbook agent session not found");
    }
    return this.requireOwnedSession(sessionState, documentId, userId);
  }

  private requireOwnedSession(
    sessionState: WorkbookAgentSessionState,
    documentId: string,
    userId: string,
  ): WorkbookAgentSessionState {
    if (sessionState.documentId !== documentId) {
      throw new Error("Workbook agent session document mismatch");
    }
    if (sessionState.userId !== userId) {
      throw new Error("Workbook agent session user mismatch");
    }
    return sessionState;
  }

  private getSessionByThreadId(threadId: string): WorkbookAgentSessionState {
    const sessionState = this.tryGetSessionByThreadId(threadId);
    if (!sessionState) {
      throw new Error(`Workbook agent session not found for thread ${threadId}`);
    }
    return sessionState;
  }

  private tryGetSessionByThreadId(threadId: string): WorkbookAgentSessionState | null {
    const sessionId = this.threadToSessionId.get(threadId);
    if (!sessionId) {
      return null;
    }
    return this.sessions.get(sessionId) ?? null;
  }

  private emitSnapshot(sessionId: string): void {
    const sessionState = this.sessions.get(sessionId);
    if (!sessionState) {
      return;
    }
    this.emit(sessionId, {
      type: "snapshot",
      snapshot: cloneSnapshot(sessionState.snapshot),
    });
  }

  private emit(sessionId: string, event: WorkbookAgentStreamEvent): void {
    const listeners = this.subscribers.get(sessionId);
    if (!listeners) {
      return;
    }
    listeners.forEach((listener) => {
      listener(event);
    });
  }

  private touch(sessionState: WorkbookAgentSessionState): void {
    sessionState.lastAccessedAt = this.now();
  }

  private evictIfNeeded(): void {
    if (this.sessions.size <= this.maxSessions) {
      return;
    }
    const candidates = [...this.sessions.values()]
      .filter((sessionState) => {
        const listeners = this.subscribers.get(sessionState.sessionId);
        return sessionState.snapshot.status === "idle" && (!listeners || listeners.size === 0);
      })
      .toSorted((left, right) => left.lastAccessedAt - right.lastAccessedAt);
    while (this.sessions.size > this.maxSessions && candidates.length > 0) {
      const evicted = candidates.shift();
      if (!evicted) {
        return;
      }
      this.sessions.delete(evicted.sessionId);
      this.threadToSessionId.delete(evicted.threadId);
      this.subscribers.delete(evicted.sessionId);
    }
  }
}

export function createWorkbookAgentService(
  zeroSyncService: ZeroSyncService,
  options: Omit<EnabledWorkbookAgentServiceOptions, "zeroSyncService"> = {},
): WorkbookAgentService {
  if (!zeroSyncService.enabled) {
    return new DisabledWorkbookAgentService();
  }
  return new EnabledWorkbookAgentService({
    zeroSyncService,
    ...options,
  });
}

export const WorkbookAgentRouteSchemas = {
  createSessionBodySchema,
  updateContextBodySchema,
  startTurnBodySchema,
};
