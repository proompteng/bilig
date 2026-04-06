import type {
  WorkbookAgentSessionSnapshot,
  WorkbookAgentStreamEvent,
  WorkbookAgentTimelineEntry,
} from "@bilig/contracts";
import { z } from "zod";
import type { SessionIdentity } from "../http/session.js";
import type { ZeroSyncService } from "../zero/service.js";
import {
  CodexAppServerClient,
  type CodexAppServerTransport,
  type CodexAppServerClientOptions,
} from "./codex-app-server-client.js";
import type {
  CodexServerNotification,
  CodexThread,
  CodexThreadItem,
} from "./codex-app-server-types.js";
import {
  handleWorkbookAgentToolCall,
  workbookAgentDynamicToolSpecs,
} from "./workbook-agent-tools.js";

const DEFAULT_MODEL = process.env["BILIG_CODEX_MODEL"]?.trim() || "gpt-5.4";

const createSessionBodySchema = z.object({
  sessionId: z.string().min(1).optional(),
  threadId: z.string().min(1).optional(),
  context: z
    .object({
      selection: z.object({
        sheetName: z.string().min(1),
        address: z.string().min(1),
      }),
      viewport: z.object({
        rowStart: z.number().int().nonnegative(),
        rowEnd: z.number().int().nonnegative(),
        colStart: z.number().int().nonnegative(),
        colEnd: z.number().int().nonnegative(),
      }),
    })
    .optional(),
});

const updateContextBodySchema = z.object({
  context: z.object({
    selection: z.object({
      sheetName: z.string().min(1),
      address: z.string().min(1),
    }),
    viewport: z.object({
      rowStart: z.number().int().nonnegative(),
      rowEnd: z.number().int().nonnegative(),
      colStart: z.number().int().nonnegative(),
      colEnd: z.number().int().nonnegative(),
    }),
  }),
});

const startTurnBodySchema = z.object({
  prompt: z.string().trim().min(1),
  context: z
    .object({
      selection: z.object({
        sheetName: z.string().min(1),
        address: z.string().min(1),
      }),
      viewport: z.object({
        rowStart: z.number().int().nonnegative(),
        rowEnd: z.number().int().nonnegative(),
        colStart: z.number().int().nonnegative(),
        colEnd: z.number().int().nonnegative(),
      }),
    })
    .optional(),
});

function createWorkbookAgentBaseInstructions(): string {
  return [
    "You are the bilig workbook assistant embedded inside a spreadsheet product.",
    "Stay narrowly focused on inspecting and editing the active workbook.",
    "Use the provided bilig.* dynamic tools for workbook work.",
    "Do not use filesystem, shell, web, connector, or unrelated tools.",
  ].join(" ");
}

function createWorkbookAgentDeveloperInstructions(): string {
  return [
    "Before changing cells you have not inspected, read the relevant workbook range first.",
    "When the user refers to the current cell, selection, or visible area, call bilig.get_context.",
    "For clear workbook edit requests, make the edit directly and then summarize what changed.",
    "Prefer one structured workbook tool call over many tiny calls when the edit is rectangular or atomic.",
    "If the requested action is outside the available bilig.* tools, say exactly which workbook capability is missing instead of improvising.",
  ].join(" ");
}

function stringifyJson(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

function formatToolContentItems(
  contentItems:
    | Array<
        | {
            type: "inputText";
            text: string;
          }
        | {
            type: "inputImage";
            imageUrl: string;
          }
      >
    | null
    | undefined,
): string | null {
  if (!contentItems || contentItems.length === 0) {
    return null;
  }
  return contentItems
    .map((item) => (item.type === "inputText" ? item.text : `[image] ${item.imageUrl}`))
    .join("\n");
}

function textFromUserContent(
  content: readonly {
    type: "text";
    text: string;
  }[],
): string {
  return content.map((item) => item.text).join("\n");
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isUserTextContentItem(value: unknown): value is { type: "text"; text: string } {
  return isRecord(value) && value["type"] === "text" && typeof value["text"] === "string";
}

function isUserMessageItem(
  item: CodexThreadItem,
): item is Extract<CodexThreadItem, { type: "userMessage" }> {
  return (
    item.type === "userMessage" &&
    Array.isArray(item.content) &&
    item.content.every((entry) => isUserTextContentItem(entry))
  );
}

function isAgentMessageItem(
  item: CodexThreadItem,
): item is Extract<CodexThreadItem, { type: "agentMessage" }> {
  return (
    item.type === "agentMessage" &&
    typeof item.text === "string" &&
    (item.phase === null || typeof item.phase === "string")
  );
}

function isPlanItem(item: CodexThreadItem): item is Extract<CodexThreadItem, { type: "plan" }> {
  return item.type === "plan" && typeof item.text === "string";
}

function isToolContentItem(item: unknown): item is
  | {
      type: "inputText";
      text: string;
    }
  | {
      type: "inputImage";
      imageUrl: string;
    } {
  return (
    typeof item === "object" &&
    item !== null &&
    (("type" in item &&
      item.type === "inputText" &&
      "text" in item &&
      typeof item.text === "string") ||
      ("type" in item &&
        item.type === "inputImage" &&
        "imageUrl" in item &&
        typeof item.imageUrl === "string"))
  );
}

function isDynamicToolCallItem(
  item: CodexThreadItem,
): item is Extract<CodexThreadItem, { type: "dynamicToolCall" }> {
  return (
    item.type === "dynamicToolCall" &&
    typeof item.tool === "string" &&
    (item.status === "inProgress" || item.status === "completed" || item.status === "failed") &&
    (item.contentItems === null ||
      (Array.isArray(item.contentItems) &&
        item.contentItems.every((entry) => isToolContentItem(entry)))) &&
    (item.success === null || typeof item.success === "boolean")
  );
}

function createSystemEntry(
  id: string,
  turnId: string | null,
  text: string,
): WorkbookAgentTimelineEntry {
  return {
    id,
    kind: "system",
    turnId,
    text,
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
  };
}

function mapThreadItemToEntry(
  item: CodexThreadItem,
  turnId: string | null,
): WorkbookAgentTimelineEntry {
  if (isUserMessageItem(item)) {
    return {
      id: item.id,
      kind: "user",
      turnId,
      text: textFromUserContent(item.content),
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
    };
  }

  if (isAgentMessageItem(item)) {
    return {
      id: item.id,
      kind: "assistant",
      turnId,
      text: item.text,
      phase: item.phase,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
    };
  }

  if (isPlanItem(item)) {
    return {
      id: item.id,
      kind: "plan",
      turnId,
      text: item.text,
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
    };
  }

  if (isDynamicToolCallItem(item)) {
    return {
      id: item.id,
      kind: "tool",
      turnId,
      text: null,
      phase: null,
      toolName: item.tool,
      toolStatus: item.status,
      argumentsText: stringifyJson(item.arguments),
      outputText: formatToolContentItems(item.contentItems),
      success: item.success,
    };
  }

  return createSystemEntry(item.id, turnId, `Codex emitted ${item.type}.`);
}

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
  -readonly [Key in keyof WorkbookAgentSessionSnapshot]: Key extends "entries"
    ? WorkbookAgentTimelineEntry[]
    : WorkbookAgentSessionSnapshot[Key];
};

function cloneSnapshot(
  snapshot: MutableWorkbookAgentSessionSnapshot,
): WorkbookAgentSessionSnapshot {
  return {
    ...snapshot,
    entries: snapshot.entries.map((entry) => ({ ...entry })),
    ...(snapshot.context ? { context: structuredClone(snapshot.context) } : { context: null }),
  };
}

function buildEntriesFromThread(thread: CodexThread): WorkbookAgentTimelineEntry[] {
  const entries: WorkbookAgentTimelineEntry[] = [];
  for (const turn of thread.turns) {
    for (const item of turn.items) {
      entries.push(mapThreadItemToEntry(item, turn.id));
    }
  }
  return entries;
}

interface WorkbookAgentSessionState {
  readonly sessionId: string;
  readonly documentId: string;
  readonly userId: string;
  threadId: string;
  snapshot: MutableWorkbookAgentSessionSnapshot;
  optimisticUserEntryIdByTurn: Map<string, string>;
  lastAccessedAt: number;
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
    };
    const sessionState: WorkbookAgentSessionState = {
      sessionId,
      documentId: input.documentId,
      userId: input.session.userID,
      threadId: thread.id,
      snapshot,
      optimisticUserEntryIdByTurn: new Map(),
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
