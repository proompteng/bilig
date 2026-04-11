import type {
  WorkbookAgentSessionSnapshot,
  WorkbookAgentStreamEvent,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowRun,
} from "@bilig/contracts";
import type {
  CodexServerNotification,
  WorkbookAgentAppliedBy,
  WorkbookAgentCommandBundle,
  WorkbookAgentContextRef,
  WorkbookAgentExecutionRecord,
  WorkbookAgentSharedReviewState,
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
  CodexAppServerClientPool,
  type CodexAppServerClientPoolStats,
  isCodexAppServerPoolBackpressureError,
} from "./codex-app-server-pool.js";
import {
  getWorkbookAgentWorkflowFamily,
  isWorkbookAgentRolloutAllowed,
  resolveWorkbookAgentFeatureFlags,
  type WorkbookAgentFeatureFlags,
} from "./workbook-agent-feature-flags.js";
import {
  handleWorkbookAgentToolCall,
  type WorkbookAgentStartWorkflowRequest,
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
  reviewPendingBundleBodySchema,
  startWorkflowBodySchema,
  startTurnBodySchema,
  updateContextBodySchema,
} from "./workbook-agent-session-model.js";
import {
  cancelWorkflowSteps,
  completeWorkflowSteps,
  createWorkflowRunRecord,
  createRunningWorkflowSteps,
  describeWorkbookAgentWorkflowTemplate,
  executeWorkbookAgentWorkflow,
  failWorkflowSteps,
} from "./workbook-agent-workflows.js";
import { isWorkflowAbortError, throwIfWorkflowCancelled } from "./workbook-agent-workflow-abort.js";
import {
  appendRevisionCitation,
  attachSharedReviewState,
  createBundleRangeCitations,
  createPendingSharedReviewState,
  createWorkflowTurnId,
  needsSharedOwnerReview,
  normalizeSharedReviewState,
} from "./workbook-agent-bundle-state.js";

const DEFAULT_MODEL = process.env["BILIG_CODEX_MODEL"]?.trim() || "gpt-5.4";
const CODEX_APP_SERVER_ARGS = ["app-server", "-c", "analytics.enabled=false"] as const;

function parsePositiveIntegerEnv(value: string | undefined, fallback: number): number {
  if (!value) {
    return fallback;
  }
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

const DEFAULT_MAX_CODEX_CLIENTS = parsePositiveIntegerEnv(
  process.env["BILIG_CODEX_MAX_CLIENTS"],
  4,
);
const DEFAULT_MAX_CODEX_CONCURRENT_TURNS_PER_CLIENT = parsePositiveIntegerEnv(
  process.env["BILIG_CODEX_MAX_CONCURRENT_TURNS_PER_CLIENT"],
  1,
);
const DEFAULT_MAX_CODEX_QUEUED_TURNS_PER_CLIENT = parsePositiveIntegerEnv(
  process.env["BILIG_CODEX_MAX_QUEUED_TURNS_PER_CLIENT"],
  8,
);
const DEFAULT_MAX_ACTIVE_TURNS_PER_USER = parsePositiveIntegerEnv(
  process.env["BILIG_CODEX_MAX_ACTIVE_TURNS_PER_USER"],
  4,
);
const DEFAULT_MAX_ACTIVE_TURNS_PER_DOCUMENT = parsePositiveIntegerEnv(
  process.env["BILIG_CODEX_MAX_ACTIVE_TURNS_PER_DOCUMENT"],
  16,
);

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

function upsertWorkflowRun(
  runs: readonly WorkbookAgentWorkflowRun[],
  nextRun: WorkbookAgentWorkflowRun,
): WorkbookAgentWorkflowRun[] {
  const index = runs.findIndex((run) => run.runId === nextRun.runId);
  if (index < 0) {
    return [nextRun, ...runs];
  }
  const nextRuns = [...runs];
  nextRuns[index] = nextRun;
  return nextRuns;
}

function mergeTimelineEntries(
  codexEntries: readonly WorkbookAgentTimelineEntry[],
  durableEntries: readonly WorkbookAgentTimelineEntry[],
): WorkbookAgentTimelineEntry[] {
  const merged = [...codexEntries];
  const indexById = new Map(merged.map((entry, index) => [entry.id, index]));
  for (const entry of durableEntries) {
    const existingIndex = indexById.get(entry.id);
    if (existingIndex === undefined) {
      indexById.set(entry.id, merged.length);
      merged.push(entry);
      continue;
    }
    merged[existingIndex] = entry;
  }
  return merged;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function readNonEmptyString(value: unknown): string | null {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : null;
}

function extractCodexNotificationErrorMessage(value: unknown): string | null {
  const direct = readNonEmptyString(value);
  if (direct) {
    return direct;
  }
  if (!isRecord(value)) {
    return null;
  }

  for (const key of ["message", "detail", "details", "reason", "hint", "title", "errorMessage"]) {
    const nested = readNonEmptyString(value[key]);
    if (nested) {
      return nested;
    }
  }

  for (const key of ["error", "cause", "data"]) {
    const nested = extractCodexNotificationErrorMessage(value[key]);
    if (nested) {
      return nested;
    }
  }

  const errors = value["errors"];
  if (Array.isArray(errors)) {
    for (const entry of errors) {
      const nested = extractCodexNotificationErrorMessage(entry);
      if (nested) {
        return nested;
      }
    }
  }

  return null;
}

function normalizeCodexNotificationErrorMessage(params: Record<string, unknown>): string {
  return (
    extractCodexNotificationErrorMessage(params) ??
    "Workbook assistant runtime failed. Retry in a moment."
  );
}

type MutableWorkbookAgentSessionSnapshot = {
  -readonly [Key in Exclude<
    keyof WorkbookAgentSessionSnapshot,
    "entries" | "pendingBundle" | "executionRecords" | "workflowRuns"
  >]: WorkbookAgentSessionSnapshot[Key];
} & {
  entries: WorkbookAgentTimelineEntry[];
  pendingBundle: WorkbookAgentCommandBundle | null;
  executionRecords: WorkbookAgentExecutionRecord[];
  workflowRuns: WorkbookAgentWorkflowRun[];
};

interface WorkbookAgentSessionState {
  readonly sessionId: string;
  readonly documentId: string;
  readonly userId: string;
  readonly storageActorUserId: string;
  scope: "private" | "shared";
  threadId: string;
  snapshot: MutableWorkbookAgentSessionSnapshot;
  optimisticUserEntryIdByTurn: Map<string, string>;
  promptByTurn: Map<string, string>;
  turnActorUserIdByTurn: Map<string, string>;
  turnContextByTurn: Map<string, WorkbookAgentUiContext | null>;
  lastAccessedAt: number;
}

interface QueuedWorkflowRun {
  readonly sessionState: WorkbookAgentSessionState;
  readonly documentId: string;
  readonly runId: string;
  readonly workflowTurnId: string;
  readonly workflowTemplate: WorkbookAgentWorkflowRun["workflowTemplate"];
  readonly workflowInput: {
    readonly query?: string;
    readonly sheetName?: string;
    readonly limit?: number;
    readonly name?: string;
  };
  readonly startedByUserId: string;
  readonly runningRun: WorkbookAgentWorkflowRun;
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

function cloneUiContext(context: WorkbookAgentUiContext | null): WorkbookAgentUiContext | null {
  return context
    ? {
        selection: {
          sheetName: context.selection.sheetName,
          address: context.selection.address,
        },
        viewport: {
          ...context.viewport,
        },
      }
    : null;
}

function isMutatingWorkflowTemplate(workflowTemplate: string): boolean {
  return (
    workflowTemplate === "highlightFormulaIssues" ||
    workflowTemplate === "repairFormulaIssues" ||
    workflowTemplate === "highlightCurrentSheetOutliers" ||
    workflowTemplate === "styleCurrentSheetHeaders" ||
    workflowTemplate === "normalizeCurrentSheetHeaders" ||
    workflowTemplate === "normalizeCurrentSheetNumberFormats" ||
    workflowTemplate === "normalizeCurrentSheetWhitespace" ||
    workflowTemplate === "fillCurrentSheetFormulasDown" ||
    workflowTemplate === "createCurrentSheetRollup" ||
    workflowTemplate === "createSheet" ||
    workflowTemplate === "renameCurrentSheet" ||
    workflowTemplate === "hideCurrentRow" ||
    workflowTemplate === "hideCurrentColumn" ||
    workflowTemplate === "unhideCurrentRow" ||
    workflowTemplate === "unhideCurrentColumn"
  );
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
  startWorkflow(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot>;
  cancelWorkflow(input: {
    documentId: string;
    sessionId: string;
    runId: string;
    session: SessionIdentity;
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
  reviewPendingBundle(input: {
    documentId: string;
    sessionId: string;
    bundleId: string;
    session: SessionIdentity;
    body: unknown;
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
  listThreads(input: {
    documentId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSummary[]>;
  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot;
  getSnapshot(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
  }): WorkbookAgentSessionSnapshot;
  subscribe(threadId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void;
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

  async startWorkflow(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async cancelWorkflow(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async interruptTurn(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async applyPendingBundle(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async reviewPendingBundle(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async dismissPendingBundle(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async replayExecutionRecord(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async listThreads(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot {
    return {
      enabled: false,
      generatedAtUnixMs: Date.now(),
      featureFlags: {
        sharedThreadsEnabled: false,
        workflowRunnerEnabled: false,
        autoApplyLowRiskEnabled: false,
        formulaWorkflowFamilyEnabled: false,
        formattingWorkflowFamilyEnabled: false,
        importWorkflowFamilyEnabled: false,
        rollupWorkflowFamilyEnabled: false,
        structuralWorkflowFamilyEnabled: false,
        allowlistedUserCount: 0,
        allowlistedDocumentCount: 0,
      },
      sessions: {
        sessionCount: 0,
        subscriberThreadCount: 0,
        subscriberCount: 0,
        activeTurnCount: 0,
        runningWorkflowCount: 0,
        pendingBundleCount: 0,
        sharedPendingReviewCount: 0,
      },
      pool: {
        slotCount: 0,
        boundThreadCount: 0,
        activeTurnCount: 0,
        queuedTurnCount: 0,
        maxClients: 0,
        maxConcurrentTurnsPerClient: 0,
        maxQueuedTurnsPerClient: 0,
      },
      counters: {
        turnBackpressureCount: 0,
        workflowStartedCount: 0,
        workflowCompletedCount: 0,
        workflowFailedCount: 0,
        workflowCancelledCount: 0,
        sharedReviewApprovedCount: 0,
        sharedReviewRejectedCount: 0,
        sharedRecommendationApprovedCount: 0,
        sharedRecommendationRejectedCount: 0,
      },
    };
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
  maxCodexClients?: number;
  maxConcurrentTurnsPerCodexClient?: number;
  maxQueuedTurnsPerCodexClient?: number;
  maxActiveTurnsPerUser?: number;
  maxActiveTurnsPerDocument?: number;
  featureFlags?: Partial<WorkbookAgentFeatureFlags>;
}

export interface WorkbookAgentObservabilitySnapshot {
  readonly enabled: boolean;
  readonly generatedAtUnixMs: number;
  readonly featureFlags: {
    readonly sharedThreadsEnabled: boolean;
    readonly workflowRunnerEnabled: boolean;
    readonly autoApplyLowRiskEnabled: boolean;
    readonly formulaWorkflowFamilyEnabled: boolean;
    readonly formattingWorkflowFamilyEnabled: boolean;
    readonly importWorkflowFamilyEnabled: boolean;
    readonly rollupWorkflowFamilyEnabled: boolean;
    readonly structuralWorkflowFamilyEnabled: boolean;
    readonly allowlistedUserCount: number;
    readonly allowlistedDocumentCount: number;
  };
  readonly sessions: {
    readonly sessionCount: number;
    readonly subscriberThreadCount: number;
    readonly subscriberCount: number;
    readonly activeTurnCount: number;
    readonly runningWorkflowCount: number;
    readonly pendingBundleCount: number;
    readonly sharedPendingReviewCount: number;
  };
  readonly pool: CodexAppServerClientPoolStats;
  readonly counters: {
    readonly turnBackpressureCount: number;
    readonly workflowStartedCount: number;
    readonly workflowCompletedCount: number;
    readonly workflowFailedCount: number;
    readonly workflowCancelledCount: number;
    readonly sharedReviewApprovedCount: number;
    readonly sharedReviewRejectedCount: number;
    readonly sharedRecommendationApprovedCount: number;
    readonly sharedRecommendationRejectedCount: number;
  };
}

class EnabledWorkbookAgentService implements WorkbookAgentService {
  readonly enabled = true;
  private readonly zeroSyncService: ZeroSyncService;
  private readonly codexClientFactory: (
    options: CodexAppServerClientOptions,
  ) => CodexAppServerTransport;
  private readonly now: () => number;
  private readonly maxSessions: number;
  private readonly maxCodexClients: number;
  private readonly maxConcurrentTurnsPerCodexClient: number;
  private readonly maxQueuedTurnsPerCodexClient: number;
  private readonly maxActiveTurnsPerUser: number;
  private readonly maxActiveTurnsPerDocument: number;
  private readonly featureFlags: WorkbookAgentFeatureFlags;
  private readonly sessions = new Map<string, WorkbookAgentSessionState>();
  private readonly threadToSessionId = new Map<string, string>();
  private readonly subscribers = new Map<string, Set<(event: WorkbookAgentStreamEvent) => void>>();
  private readonly workflowRunTasks = new Map<string, Promise<void>>();
  private readonly workflowAbortControllers = new Map<string, AbortController>();
  private readonly counters = {
    turnBackpressureCount: 0,
    workflowStartedCount: 0,
    workflowCompletedCount: 0,
    workflowFailedCount: 0,
    workflowCancelledCount: 0,
    sharedReviewApprovedCount: 0,
    sharedReviewRejectedCount: 0,
    sharedRecommendationApprovedCount: 0,
    sharedRecommendationRejectedCount: 0,
  };
  private codexClient: CodexAppServerClientPool | null = null;
  private unsubscribeCodex: (() => void) | null = null;

  constructor(options: EnabledWorkbookAgentServiceOptions) {
    this.zeroSyncService = options.zeroSyncService;
    this.codexClientFactory =
      options.codexClientFactory ?? ((clientOptions) => new CodexAppServerClient(clientOptions));
    this.now = options.now ?? (() => Date.now());
    this.maxSessions = options.maxSessions ?? 64;
    this.maxCodexClients = options.maxCodexClients ?? DEFAULT_MAX_CODEX_CLIENTS;
    this.maxConcurrentTurnsPerCodexClient =
      options.maxConcurrentTurnsPerCodexClient ?? DEFAULT_MAX_CODEX_CONCURRENT_TURNS_PER_CLIENT;
    this.maxQueuedTurnsPerCodexClient =
      options.maxQueuedTurnsPerCodexClient ?? DEFAULT_MAX_CODEX_QUEUED_TURNS_PER_CLIENT;
    this.maxActiveTurnsPerUser = options.maxActiveTurnsPerUser ?? DEFAULT_MAX_ACTIVE_TURNS_PER_USER;
    this.maxActiveTurnsPerDocument =
      options.maxActiveTurnsPerDocument ?? DEFAULT_MAX_ACTIVE_TURNS_PER_DOCUMENT;
    this.featureFlags = {
      ...resolveWorkbookAgentFeatureFlags(),
      ...options.featureFlags,
    };
  }

  private async persistSessionState(sessionState: WorkbookAgentSessionState): Promise<void> {
    await this.zeroSyncService.saveWorkbookAgentThreadState({
      documentId: sessionState.documentId,
      threadId: sessionState.threadId,
      actorUserId: sessionState.storageActorUserId,
      scope: sessionState.scope,
      context: sessionState.snapshot.context,
      entries: sessionState.snapshot.entries,
      pendingBundle: sessionState.snapshot.pendingBundle,
      updatedAtUnixMs: this.now(),
    });
  }

  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot {
    const sessions = [...this.sessions.values()];
    const subscriberSets = [...this.subscribers.values()];
    const poolStats = this.codexClient?.getStats() ?? {
      slotCount: 0,
      boundThreadCount: 0,
      activeTurnCount: 0,
      queuedTurnCount: 0,
      maxClients: this.maxCodexClients,
      maxConcurrentTurnsPerClient: this.maxConcurrentTurnsPerCodexClient,
      maxQueuedTurnsPerClient: this.maxQueuedTurnsPerCodexClient,
    };
    return {
      enabled: true,
      generatedAtUnixMs: this.now(),
      featureFlags: {
        sharedThreadsEnabled: this.featureFlags.sharedThreadsEnabled,
        workflowRunnerEnabled: this.featureFlags.workflowRunnerEnabled,
        autoApplyLowRiskEnabled: this.featureFlags.autoApplyLowRiskEnabled,
        formulaWorkflowFamilyEnabled: this.featureFlags.formulaWorkflowFamilyEnabled,
        formattingWorkflowFamilyEnabled: this.featureFlags.formattingWorkflowFamilyEnabled,
        importWorkflowFamilyEnabled: this.featureFlags.importWorkflowFamilyEnabled,
        rollupWorkflowFamilyEnabled: this.featureFlags.rollupWorkflowFamilyEnabled,
        structuralWorkflowFamilyEnabled: this.featureFlags.structuralWorkflowFamilyEnabled,
        allowlistedUserCount: this.featureFlags.allowlistedUserIds.length,
        allowlistedDocumentCount: this.featureFlags.allowlistedDocumentIds.length,
      },
      sessions: {
        sessionCount: sessions.length,
        subscriberThreadCount: this.subscribers.size,
        subscriberCount: subscriberSets.reduce((sum, listeners) => sum + listeners.size, 0),
        activeTurnCount: sessions.filter(
          (sessionState) => sessionState.snapshot.activeTurnId !== null,
        ).length,
        runningWorkflowCount: sessions.reduce(
          (sum, sessionState) =>
            sum +
            sessionState.snapshot.workflowRuns.filter((run) => run.status === "running").length,
          0,
        ),
        pendingBundleCount: sessions.filter(
          (sessionState) => sessionState.snapshot.pendingBundle !== null,
        ).length,
        sharedPendingReviewCount: sessions.filter((sessionState) => {
          const pendingBundle = sessionState.snapshot.pendingBundle;
          return (
            pendingBundle !== null &&
            sessionState.scope === "shared" &&
            normalizeSharedReviewState(pendingBundle, sessionState)?.status === "pending"
          );
        }).length,
      },
      pool: poolStats,
      counters: { ...this.counters },
    };
  }

  private isRolloutAllowed(documentId: string, userId: string): boolean {
    return isWorkbookAgentRolloutAllowed(this.featureFlags, { documentId, userId });
  }

  private assertRolloutAllowed(input: {
    documentId: string;
    userId: string;
    code: string;
    message: string;
  }): void {
    if (this.isRolloutAllowed(input.documentId, input.userId)) {
      return;
    }
    throw createWorkbookAgentServiceError({
      code: input.code,
      message: input.message,
      statusCode: 409,
      retryable: false,
    });
  }

  async createSession(input: {
    documentId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const parsed = createSessionBodySchema.parse(input.body);
    if (parsed.scope === "shared" && !this.featureFlags.sharedThreadsEnabled) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SHARED_THREADS_DISABLED",
        message: "Shared workbook assistant threads are currently disabled.",
        statusCode: 409,
        retryable: false,
      });
    }
    if (parsed.scope === "shared") {
      this.assertRolloutAllowed({
        documentId: input.documentId,
        userId: input.session.userID,
        code: "WORKBOOK_AGENT_SHARED_THREADS_ROLLOUT_BLOCKED",
        message: "Shared workbook assistant threads are still limited to the rollout allowlist.",
      });
    }
    const sessionId = parsed.sessionId ?? crypto.randomUUID();
    if (parsed.threadId !== undefined) {
      const sharedSession = this.tryGetSessionByThreadId(parsed.threadId);
      if (sharedSession) {
        const accessibleSession = this.requireOwnedSession(
          sharedSession,
          input.documentId,
          input.session.userID,
        );
        if (parsed.context) {
          accessibleSession.snapshot.context = parsed.context;
          await this.persistSessionState(accessibleSession);
          this.emitSnapshot(accessibleSession.threadId);
        }
        this.touch(accessibleSession);
        return cloneSnapshot(accessibleSession.snapshot);
      }
    }
    const existing = this.sessions.get(sessionId);
    if (existing) {
      const sessionState = this.requireOwnedSession(
        existing,
        input.documentId,
        input.session.userID,
      );
      if (parsed.context) {
        sessionState.snapshot.context = parsed.context;
        await this.persistSessionState(sessionState);
        this.emitSnapshot(existing.threadId);
      }
      this.touch(sessionState);
      return cloneSnapshot(sessionState.snapshot);
    }

    let thread: Awaited<ReturnType<CodexAppServerTransport["threadStart"]>> | null = null;
    let sessionBootstrapError: unknown = null;
    try {
      const codexClient = await this.getCodexClient();
      thread =
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
    } catch (error) {
      if (parsed.threadId === undefined) {
        throw error;
      }
      sessionBootstrapError = error;
    }
    const threadId = thread?.id ?? parsed.threadId;
    if (!threadId) {
      throw sessionBootstrapError instanceof Error
        ? sessionBootstrapError
        : new Error("Workbook agent thread bootstrap failed");
    }
    const durableThreadState = await this.zeroSyncService.loadWorkbookAgentThreadState(
      input.documentId,
      input.session.userID,
      threadId,
    );
    if (durableThreadState?.scope === "shared" && !this.featureFlags.sharedThreadsEnabled) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SHARED_THREADS_DISABLED",
        message: "Shared workbook assistant threads are currently disabled.",
        statusCode: 409,
        retryable: false,
      });
    }
    if (durableThreadState?.scope === "shared") {
      this.assertRolloutAllowed({
        documentId: input.documentId,
        userId: input.session.userID,
        code: "WORKBOOK_AGENT_SHARED_THREADS_ROLLOUT_BLOCKED",
        message: "Shared workbook assistant threads are still limited to the rollout allowlist.",
      });
    }
    if (!thread && !durableThreadState) {
      throw sessionBootstrapError instanceof Error
        ? sessionBootstrapError
        : new Error("Workbook agent thread bootstrap failed");
    }
    const resolvedScope = durableThreadState?.scope ?? parsed.scope ?? "private";
    const executionRecords = await this.zeroSyncService.listWorkbookAgentThreadRuns(
      input.documentId,
      input.session.userID,
      threadId,
    );
    const workflowRuns = await this.zeroSyncService.listWorkbookThreadWorkflowRuns(
      input.documentId,
      input.session.userID,
      threadId,
    );
    const codexEntries = thread ? buildEntriesFromThread(thread) : [];
    const bootstrapErrorMessage =
      thread || !sessionBootstrapError
        ? null
        : sessionBootstrapError instanceof Error
          ? sessionBootstrapError.message
          : "Workbook assistant live session is unavailable. Loaded durable thread history only.";

    const snapshot: MutableWorkbookAgentSessionSnapshot = {
      sessionId,
      documentId: input.documentId,
      threadId,
      scope: resolvedScope,
      status: !thread
        ? "failed"
        : thread.turns.some((turn) => turn.status === "failed")
          ? "failed"
          : thread.turns.some((turn) => turn.status === "inProgress")
            ? "inProgress"
            : "idle",
      activeTurnId: thread?.turns.findLast((turn) => turn.status === "inProgress")?.id ?? null,
      lastError:
        thread?.turns.findLast((turn) => turn.error?.message)?.error?.message ??
        bootstrapErrorMessage,
      context: parsed.context ?? durableThreadState?.context ?? null,
      entries: mergeTimelineEntries(codexEntries, durableThreadState?.entries ?? []),
      pendingBundle: durableThreadState?.pendingBundle ?? null,
      executionRecords,
      workflowRuns,
    };
    const sessionState: WorkbookAgentSessionState = {
      sessionId,
      documentId: input.documentId,
      userId: input.session.userID,
      storageActorUserId: durableThreadState?.actorUserId ?? input.session.userID,
      scope: resolvedScope,
      threadId,
      snapshot,
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn: new Map(),
      turnContextByTurn: new Map(),
      lastAccessedAt: this.now(),
    };
    this.sessions.set(sessionId, sessionState);
    this.threadToSessionId.set(threadId, sessionId);
    this.evictIfNeeded();
    await this.persistSessionState(sessionState);
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
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
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
    if (sessionState.snapshot.activeTurnId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_TURN_ALREADY_RUNNING",
        message: "Finish or interrupt the current assistant turn before starting another one.",
        statusCode: 409,
        retryable: false,
      });
    }
    this.assertTurnQuota(input.documentId, input.session.userID);
    if (parsed.context) {
      sessionState.snapshot.context = parsed.context;
    }
    const turnContext = cloneUiContext(parsed.context ?? sessionState.snapshot.context);
    const codexClient = await this.getCodexClient();
    let turn;
    try {
      turn = await codexClient.turnStart({
        threadId: sessionState.threadId,
        prompt: parsed.prompt,
      });
    } catch (error) {
      if (isCodexAppServerPoolBackpressureError(error)) {
        this.counters.turnBackpressureCount += 1;
        throw createWorkbookAgentServiceError({
          code: "WORKBOOK_AGENT_TURN_BACKPRESSURE",
          message: error.message,
          statusCode: 429,
          retryable: true,
        });
      }
      throw error;
    }
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
      citations: [],
    });
    sessionState.optimisticUserEntryIdByTurn.set(turn.id, optimisticEntryId);
    sessionState.promptByTurn.set(turn.id, parsed.prompt);
    sessionState.turnActorUserIdByTurn.set(turn.id, input.session.userID);
    sessionState.turnContextByTurn.set(turn.id, turnContext);
    sessionState.snapshot.activeTurnId = turn.id;
    sessionState.snapshot.status = "inProgress";
    sessionState.snapshot.lastError = null;
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return cloneSnapshot(sessionState.snapshot);
  }

  async startWorkflow(input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const parsed = startWorkflowBodySchema.parse(input.body);
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    if (!this.featureFlags.workflowRunnerEnabled) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_WORKFLOW_RUNNER_DISABLED",
        message: "Workbook assistant workflows are currently disabled.",
        statusCode: 409,
        retryable: false,
      });
    }
    this.assertRolloutAllowed({
      documentId: input.documentId,
      userId: input.session.userID,
      code: "WORKBOOK_AGENT_WORKFLOW_RUNNER_ROLLOUT_BLOCKED",
      message: "Workbook assistant workflows are still limited to the rollout allowlist.",
    });
    if (parsed.context) {
      sessionState.snapshot.context = parsed.context;
    }
    const runningWorkflow = sessionState.snapshot.workflowRuns.find(
      (run) => run.status === "running",
    );
    if (runningWorkflow) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_WORKFLOW_ALREADY_RUNNING",
        message: `Finish or cancel the running workflow before starting another one: ${runningWorkflow.title}`,
        statusCode: 409,
        retryable: false,
      });
    }
    const runId = crypto.randomUUID();
    const now = this.now();
    const workflowTemplate = parsed.workflowTemplate;
    this.assertWorkflowFamilyEnabled(workflowTemplate);
    if (sessionState.snapshot.pendingBundle && isMutatingWorkflowTemplate(workflowTemplate)) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_PENDING_BUNDLE_EXISTS",
        message:
          "Apply or dismiss the staged preview bundle before starting another mutating workflow.",
        statusCode: 409,
        retryable: false,
      });
    }
    const workflowTurnId = sessionState.snapshot.activeTurnId ?? createWorkflowTurnId(runId);
    const workflowInput = {
      ...("query" in parsed ? { query: parsed.query } : {}),
      ...("sheetName" in parsed && parsed.sheetName ? { sheetName: parsed.sheetName } : {}),
      ...("limit" in parsed && parsed.limit !== undefined ? { limit: parsed.limit } : {}),
      ...("name" in parsed ? { name: parsed.name } : {}),
    };
    const workflowDescription = describeWorkbookAgentWorkflowTemplate(
      workflowTemplate,
      workflowInput,
    );
    const runningRun = createWorkflowRunRecord({
      runId,
      threadId: sessionState.threadId,
      startedByUserId: input.session.userID,
      workflowTemplate,
      title: workflowDescription.title,
      summary: workflowDescription.runningSummary,
      status: "running",
      now,
      steps: createRunningWorkflowSteps(workflowTemplate, now, workflowInput),
    });
    sessionState.snapshot.workflowRuns = upsertWorkflowRun(
      sessionState.snapshot.workflowRuns,
      runningRun,
    );
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createSystemEntry(
        `system-workflow-start:${runId}`,
        workflowTurnId,
        `Started workflow: ${runningRun.title}`,
      ),
    );
    this.touch(sessionState);
    this.counters.workflowStartedCount += 1;
    await this.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, runningRun);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    this.queueWorkflowRun({
      sessionState,
      documentId: input.documentId,
      runId,
      workflowTurnId,
      workflowTemplate,
      workflowInput,
      startedByUserId: input.session.userID,
      runningRun,
    });
    return cloneSnapshot(sessionState.snapshot);
  }

  private queueWorkflowRun(input: QueuedWorkflowRun): void {
    const existingTask =
      this.workflowRunTasks.get(input.sessionState.threadId) ?? Promise.resolve();
    const nextTask = existingTask
      .catch(() => undefined)
      .then(() => this.executeQueuedWorkflowRun(input));
    this.workflowRunTasks.set(input.sessionState.threadId, nextTask);
    void nextTask
      .catch((error: unknown) => {
        console.error(error);
      })
      .finally(() => {
        if (this.workflowRunTasks.get(input.sessionState.threadId) === nextTask) {
          this.workflowRunTasks.delete(input.sessionState.threadId);
        }
      });
  }

  private async executeQueuedWorkflowRun(input: QueuedWorkflowRun): Promise<void> {
    const currentRun = input.sessionState.snapshot.workflowRuns.find(
      (run) => run.runId === input.runId,
    );
    if (!currentRun || currentRun.status !== "running") {
      return;
    }
    const abortController = new AbortController();
    this.workflowAbortControllers.set(input.runId, abortController);

    try {
      const result = await executeWorkbookAgentWorkflow({
        documentId: input.documentId,
        zeroSyncService: this.zeroSyncService,
        workflowTemplate: input.workflowTemplate,
        context: input.sessionState.snapshot.context,
        workflowInput: input.workflowInput,
        signal: abortController.signal,
      });
      throwIfWorkflowCancelled(abortController.signal);

      const latestRun = input.sessionState.snapshot.workflowRuns.find(
        (run) => run.runId === input.runId,
      );
      if (latestRun?.status === "cancelled") {
        return;
      }

      const completedAtUnixMs = this.now();
      const completedRun: WorkbookAgentWorkflowRun = {
        ...input.runningRun,
        title: result.title,
        summary: result.summary,
        status: "completed",
        updatedAtUnixMs: completedAtUnixMs,
        completedAtUnixMs,
        artifact: result.artifact,
        steps: completeWorkflowSteps(
          input.workflowTemplate,
          result.steps,
          completedAtUnixMs,
          input.workflowInput,
        ),
      };
      input.sessionState.snapshot.workflowRuns = upsertWorkflowRun(
        input.sessionState.snapshot.workflowRuns,
        completedRun,
      );
      if (result.commands && result.commands.length > 0) {
        const baseRevision = await this.zeroSyncService.getWorkbookHeadRevision(input.documentId);
        const workflowBundle = attachSharedReviewState(
          createWorkbookAgentCommandBundle({
            documentId: input.documentId,
            threadId: input.sessionState.threadId,
            turnId: input.workflowTurnId,
            goalText: result.goalText ?? result.title,
            baseRevision,
            context: toContextRef(input.sessionState.snapshot.context),
            commands: result.commands,
            now: completedAtUnixMs,
          }),
          input.sessionState,
        );
        input.sessionState.snapshot.pendingBundle = workflowBundle;
        input.sessionState.snapshot.entries = upsertEntry(
          input.sessionState.snapshot.entries,
          createSystemEntry(
            `system-preview:${workflowBundle.id}`,
            input.workflowTurnId,
            describeWorkbookAgentBundle(workflowBundle),
            createBundleRangeCitations(workflowBundle),
          ),
        );
      }
      input.sessionState.snapshot.entries = upsertEntry(
        input.sessionState.snapshot.entries,
        createSystemEntry(
          `system-workflow-complete:${input.runId}`,
          input.workflowTurnId,
          `Completed workflow: ${result.title}`,
          result.citations,
        ),
      );
      this.touch(input.sessionState);
      this.counters.workflowCompletedCount += 1;
      await this.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, completedRun);
      await this.persistSessionState(input.sessionState);
      this.emitSnapshot(input.sessionState.threadId);
      return;
    } catch (error) {
      if (isWorkflowAbortError(error)) {
        return;
      }
      const latestRun = input.sessionState.snapshot.workflowRuns.find(
        (run) => run.runId === input.runId,
      );
      if (latestRun?.status === "cancelled") {
        return;
      }
      const failedAtUnixMs = this.now();
      const errorMessage = error instanceof Error ? error.message : String(error);
      const failedRun: WorkbookAgentWorkflowRun = {
        ...input.runningRun,
        status: "failed",
        summary: `Workflow failed: ${input.runningRun.title}`,
        updatedAtUnixMs: failedAtUnixMs,
        completedAtUnixMs: failedAtUnixMs,
        errorMessage,
        steps: failWorkflowSteps(
          input.workflowTemplate,
          input.runningRun.steps,
          errorMessage,
          failedAtUnixMs,
          input.workflowInput,
        ),
        artifact: null,
      };
      input.sessionState.snapshot.workflowRuns = upsertWorkflowRun(
        input.sessionState.snapshot.workflowRuns,
        failedRun,
      );
      input.sessionState.snapshot.entries = upsertEntry(
        input.sessionState.snapshot.entries,
        createSystemEntry(
          `system-workflow-failed:${input.runId}`,
          input.workflowTurnId,
          failedRun.errorMessage ?? `Workflow failed: ${input.runningRun.title}`,
        ),
      );
      this.touch(input.sessionState);
      this.counters.workflowFailedCount += 1;
      await this.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, failedRun);
      await this.persistSessionState(input.sessionState);
      this.emitSnapshot(input.sessionState.threadId);
      return;
    } finally {
      this.workflowAbortControllers.delete(input.runId);
    }
  }

  async cancelWorkflow(input: {
    documentId: string;
    sessionId: string;
    runId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.sessionId,
      input.session.userID,
    );
    const runningWorkflow = sessionState.snapshot.workflowRuns.find(
      (run) => run.runId === input.runId,
    );
    if (!runningWorkflow) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_WORKFLOW_NOT_FOUND",
        message: "Workbook agent workflow run not found",
        statusCode: 404,
        retryable: false,
      });
    }
    if (runningWorkflow.status !== "running") {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_WORKFLOW_NOT_RUNNING",
        message: "Workbook agent workflow is not currently running",
        statusCode: 409,
        retryable: false,
      });
    }
    const now = this.now();
    const cancelledRun: WorkbookAgentWorkflowRun = {
      ...runningWorkflow,
      status: "cancelled",
      summary: `Cancelled workflow: ${runningWorkflow.title}`,
      updatedAtUnixMs: now,
      completedAtUnixMs: now,
      errorMessage: `Cancelled by ${input.session.userID}.`,
      steps: cancelWorkflowSteps(runningWorkflow.steps, now),
      artifact: null,
    };
    sessionState.snapshot.workflowRuns = upsertWorkflowRun(
      sessionState.snapshot.workflowRuns,
      cancelledRun,
    );
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createSystemEntry(
        `system-workflow-cancel:${input.runId}:${now}`,
        sessionState.snapshot.activeTurnId,
        `Cancelled workflow: ${runningWorkflow.title}`,
      ),
    );
    this.workflowAbortControllers.get(input.runId)?.abort();
    this.touch(sessionState);
    this.counters.workflowCancelledCount += 1;
    await this.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, cancelledRun);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
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
    if (input.appliedBy === "auto" && !this.featureFlags.autoApplyLowRiskEnabled) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_AUTO_APPLY_DISABLED",
        message: "Workbook agent auto-apply is currently disabled.",
        statusCode: 409,
        retryable: false,
      });
    }
    if (input.appliedBy === "auto") {
      this.assertRolloutAllowed({
        documentId: input.documentId,
        userId: input.session.userID,
        code: "WORKBOOK_AGENT_AUTO_APPLY_ROLLOUT_BLOCKED",
        message: "Workbook agent auto-apply is still limited to the rollout allowlist.",
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
    if (
      sessionState.scope === "shared" &&
      pendingBundle.riskClass !== "low" &&
      sessionState.storageActorUserId !== input.session.userID
    ) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SHARED_APPROVAL_REQUIRED",
        message: "Shared medium/high-risk workbook bundles must be applied by the thread owner.",
        statusCode: 409,
        retryable: false,
      });
    }
    const sharedReview = normalizeSharedReviewState(pendingBundle, sessionState);
    if (
      sessionState.scope === "shared" &&
      pendingBundle.riskClass !== "low" &&
      sharedReview?.status !== "approved"
    ) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SHARED_REVIEW_REQUIRED",
        message:
          "Shared medium/high-risk workbook bundles must be approved by the thread owner before apply.",
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
        : attachSharedReviewState(
            createWorkbookAgentCommandBundle({
              documentId: selection.remainingBundle.documentId,
              threadId: selection.remainingBundle.threadId,
              turnId: selection.remainingBundle.turnId,
              goalText: selection.remainingBundle.goalText,
              baseRevision: result.revision,
              context: selection.remainingBundle.context,
              commands: selection.remainingBundle.commands,
              now: this.now(),
            }),
            sessionState,
          );
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createSystemEntry(
        `system-apply:${executionRecord.id}`,
        pendingBundle.turnId,
        `${input.appliedBy === "auto" ? "Auto-applied" : "Applied"} ${
          selection.acceptedScope === "partial" ? "selected " : ""
        }preview bundle at revision r${String(result.revision)}: ${selection.acceptedBundle.summary}`,
        appendRevisionCitation(
          createBundleRangeCitations(selection.acceptedBundle),
          result.revision,
        ),
      ),
    );
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return cloneSnapshot(sessionState.snapshot);
  }

  async reviewPendingBundle(input: {
    documentId: string;
    sessionId: string;
    bundleId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentSessionSnapshot> {
    const parsed = reviewPendingBundleBodySchema.parse(input.body);
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
    if (!needsSharedOwnerReview(sessionState, pendingBundle)) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SHARED_REVIEW_NOT_REQUIRED",
        message: "Shared review is only required for medium/high-risk bundles.",
        statusCode: 409,
        retryable: false,
      });
    }
    const now = this.now();
    const sharedReview =
      normalizeSharedReviewState(pendingBundle, sessionState) ??
      createPendingSharedReviewState(sessionState.storageActorUserId);
    const isOwnerReviewer = sessionState.storageActorUserId === input.session.userID;
    const nextSharedReview: WorkbookAgentSharedReviewState = isOwnerReviewer
      ? {
          ...sharedReview,
          status: parsed.decision,
          decidedByUserId: input.session.userID,
          decidedAtUnixMs: now,
        }
      : {
          ...sharedReview,
          recommendations: [
            ...sharedReview.recommendations.filter(
              (recommendation) => recommendation.userId !== input.session.userID,
            ),
            {
              userId: input.session.userID,
              decision: parsed.decision,
              decidedAtUnixMs: now,
            },
          ].toSorted((left, right) => left.userId.localeCompare(right.userId)),
        };
    const reviewedBundle: WorkbookAgentCommandBundle = {
      ...pendingBundle,
      sharedReview: nextSharedReview,
    };
    if (isOwnerReviewer) {
      if (parsed.decision === "approved") {
        this.counters.sharedReviewApprovedCount += 1;
      } else {
        this.counters.sharedReviewRejectedCount += 1;
      }
    } else if (parsed.decision === "approved") {
      this.counters.sharedRecommendationApprovedCount += 1;
    } else {
      this.counters.sharedRecommendationRejectedCount += 1;
    }
    sessionState.snapshot.pendingBundle = reviewedBundle;
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createSystemEntry(
        `system-review:${reviewedBundle.id}:${now}`,
        reviewedBundle.turnId,
        isOwnerReviewer
          ? `${parsed.decision === "approved" ? "Approved" : "Rejected"} shared preview bundle: ${reviewedBundle.summary}`
          : `${input.session.userID} recommended ${parsed.decision === "approved" ? "approval" : "rejection"} for shared preview bundle: ${reviewedBundle.summary}`,
        createBundleRangeCitations(reviewedBundle),
      ),
    );
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
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
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_BUNDLE_NOT_FOUND",
        message: "Workbook agent preview bundle not found",
        statusCode: 404,
        retryable: false,
      });
    }
    sessionState.snapshot.pendingBundle = null;
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createSystemEntry(
        `system-dismiss:${pendingBundle.id}:${this.now()}`,
        pendingBundle.turnId,
        `Dismissed preview bundle: ${pendingBundle.summary}`,
        createBundleRangeCitations(pendingBundle),
      ),
    );
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
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
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_RUN_NOT_FOUND",
        message: "Workbook agent execution record not found",
        statusCode: 404,
        retryable: false,
      });
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
        createBundleRangeCitations(replayedBundle),
      ),
    );
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return cloneSnapshot(sessionState.snapshot);
  }

  async listThreads(input: {
    documentId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSummary[]> {
    return await this.zeroSyncService.listWorkbookAgentThreadSummaries(
      input.documentId,
      input.session.userID,
    );
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

  subscribe(threadId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void {
    const listeners = this.subscribers.get(threadId) ?? new Set();
    listeners.add(listener);
    this.subscribers.set(threadId, listeners);
    return () => {
      const current = this.subscribers.get(threadId);
      if (!current) {
        return;
      }
      current.delete(listener);
      if (current.size === 0) {
        this.subscribers.delete(threadId);
      }
    };
  }

  async close(): Promise<void> {
    this.unsubscribeCodex?.();
    this.unsubscribeCodex = null;
    this.workflowAbortControllers.forEach((controller) => {
      controller.abort();
    });
    this.workflowAbortControllers.clear();
    this.workflowRunTasks.clear();
    await this.codexClient?.close();
    this.codexClient = null;
    this.sessions.clear();
    this.threadToSessionId.clear();
    this.subscribers.clear();
  }

  private async getCodexClient(): Promise<CodexAppServerTransport> {
    if (!this.codexClient) {
      this.codexClient = new CodexAppServerClientPool({
        codexClientFactory: this.codexClientFactory,
        maxClients: this.maxCodexClients,
        maxConcurrentTurnsPerClient: this.maxConcurrentTurnsPerCodexClient,
        maxQueuedTurnsPerClient: this.maxQueuedTurnsPerCodexClient,
        clientOptions: {
          command: process.env["BILIG_CODEX_BIN"]?.trim() || "codex",
          args: [...CODEX_APP_SERVER_ARGS],
          cwd: process.cwd(),
          env: process.env,
          onLog: (message) => {
            if (message.length > 0) {
              console.error(message);
            }
          },
          handleDynamicToolCall: (request) => {
            const sessionState = this.getSessionByThreadId(request.threadId);
            const requestActorUserId = this.resolveTurnActorUserId(sessionState, request.turnId);
            const requestContext = this.resolveTurnContext(sessionState, request.turnId);
            return handleWorkbookAgentToolCall(
              {
                documentId: sessionState.documentId,
                session: {
                  userID: requestActorUserId,
                  roles: ["editor"],
                },
                uiContext: requestContext,
                zeroSyncService: this.zeroSyncService,
                stageCommand: async (command) => {
                  const baseRevision = await this.zeroSyncService.getWorkbookHeadRevision(
                    sessionState.documentId,
                  );
                  const bundle = attachSharedReviewState(
                    appendWorkbookAgentCommandToBundle({
                      previousBundle: sessionState.snapshot.pendingBundle,
                      documentId: sessionState.documentId,
                      threadId: sessionState.threadId,
                      turnId: request.turnId,
                      goalText:
                        sessionState.promptByTurn.get(request.turnId) ??
                        "Update workbook from assistant request",
                      baseRevision,
                      context: toContextRef(requestContext),
                      command,
                      now: this.now(),
                    }),
                    sessionState,
                  );
                  sessionState.snapshot.pendingBundle = bundle;
                  sessionState.snapshot.entries = upsertEntry(
                    sessionState.snapshot.entries,
                    createSystemEntry(
                      `system-preview:${bundle.id}`,
                      request.turnId,
                      describeWorkbookAgentBundle(bundle),
                      createBundleRangeCitations(bundle),
                    ),
                  );
                  await this.persistSessionState(sessionState);
                  this.emitSnapshot(sessionState.threadId);
                  return bundle;
                },
                startWorkflow: async (workflowRequest: WorkbookAgentStartWorkflowRequest) => {
                  const previousRunIds = new Set(
                    sessionState.snapshot.workflowRuns.map((run) => run.runId),
                  );
                  const nextSnapshot = await this.startWorkflow({
                    documentId: sessionState.documentId,
                    sessionId: sessionState.sessionId,
                    session: {
                      userID: requestActorUserId,
                      roles: ["editor"],
                    },
                    body: {
                      ...workflowRequest,
                      ...(requestContext ? { context: requestContext } : {}),
                    },
                  });
                  const nextRun =
                    nextSnapshot.workflowRuns.find((run) => !previousRunIds.has(run.runId)) ??
                    nextSnapshot.workflowRuns.find(
                      (run) => run.workflowTemplate === workflowRequest.workflowTemplate,
                    );
                  if (!nextRun) {
                    throw new Error(
                      `Workflow run not found after starting ${workflowRequest.workflowTemplate}`,
                    );
                  }
                  return nextRun;
                },
              },
              request,
            );
          },
        },
      });
      await this.codexClient.ensureReady();
      this.unsubscribeCodex = this.codexClient.subscribe((notification) => {
        void this.handleCodexNotification(notification).catch((error: unknown) => {
          console.error(error);
        });
      });
    }
    return this.codexClient;
  }

  private async handleCodexNotification(notification: CodexServerNotification): Promise<void> {
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
        this.emitSnapshot(sessionState.threadId);
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
        sessionState.promptByTurn.delete(notification.params.turn.id);
        sessionState.turnActorUserIdByTurn.delete(notification.params.turn.id);
        sessionState.turnContextByTurn.delete(notification.params.turn.id);
        await this.persistSessionState(sessionState);
        this.emitSnapshot(sessionState.threadId);
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
        await this.persistSessionState(sessionState);
        this.emitSnapshot(sessionState.threadId);
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
            citations: [],
          } satisfies WorkbookAgentTimelineEntry);
        sessionState.snapshot.entries = upsertEntry(sessionState.snapshot.entries, {
          ...existing,
          text: `${existing.text ?? ""}${notification.params.delta}`,
        });
        this.emit(sessionState.threadId, {
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
            citations: [],
          } satisfies WorkbookAgentTimelineEntry);
        sessionState.snapshot.entries = upsertEntry(sessionState.snapshot.entries, {
          ...existing,
          text: `${existing.text ?? ""}${notification.params.delta}`,
        });
        this.emit(sessionState.threadId, {
          type: "planDelta",
          itemId: notification.params.itemId,
          delta: notification.params.delta,
        });
        return;
      }
      case "error": {
        const message = normalizeCodexNotificationErrorMessage(notification.params);
        await Promise.all(
          [...this.sessions.values()].map(async (sessionState) => {
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
            await this.persistSessionState(sessionState);
            this.emitSnapshot(sessionState.threadId);
          }),
        );
        return;
      }
    }
  }

  private resolveTurnActorUserId(sessionState: WorkbookAgentSessionState, turnId: string): string {
    return sessionState.turnActorUserIdByTurn.get(turnId) ?? sessionState.userId;
  }

  private resolveTurnContext(
    sessionState: WorkbookAgentSessionState,
    turnId: string,
  ): WorkbookAgentUiContext | null {
    return cloneUiContext(
      sessionState.turnContextByTurn.get(turnId) ?? sessionState.snapshot.context,
    );
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
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SESSION_NOT_FOUND",
        message: "Workbook agent session not found",
        statusCode: 404,
        retryable: true,
      });
    }
    return this.requireOwnedSession(sessionState, documentId, userId);
  }

  private requireOwnedSession(
    sessionState: WorkbookAgentSessionState,
    documentId: string,
    userId: string,
  ): WorkbookAgentSessionState {
    if (sessionState.documentId !== documentId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SESSION_NOT_FOUND",
        message: "Workbook agent session not found",
        statusCode: 404,
        retryable: false,
      });
    }
    if (sessionState.scope !== "shared" && sessionState.userId !== userId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SESSION_NOT_FOUND",
        message: "Workbook agent session not found",
        statusCode: 404,
        retryable: false,
      });
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

  private emitSnapshot(threadId: string): void {
    const sessionState = this.tryGetSessionByThreadId(threadId);
    if (!sessionState) {
      return;
    }
    this.emit(threadId, {
      type: "snapshot",
      snapshot: cloneSnapshot(sessionState.snapshot),
    });
  }

  private emit(threadId: string, event: WorkbookAgentStreamEvent): void {
    const listeners = this.subscribers.get(threadId);
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

  private resolveActiveTurnActorUserId(sessionState: WorkbookAgentSessionState): string | null {
    const activeTurnId = sessionState.snapshot.activeTurnId;
    if (!activeTurnId) {
      return null;
    }
    return sessionState.turnActorUserIdByTurn.get(activeTurnId) ?? sessionState.userId;
  }

  private assertTurnQuota(documentId: string, actorUserId: string): void {
    const activeSessions = [...this.sessions.values()].filter(
      (sessionState) =>
        sessionState.snapshot.activeTurnId !== null &&
        sessionState.snapshot.status === "inProgress",
    );
    const activeTurnsForUser = activeSessions.filter(
      (sessionState) => this.resolveActiveTurnActorUserId(sessionState) === actorUserId,
    ).length;
    if (activeTurnsForUser >= this.maxActiveTurnsPerUser) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_USER_TURN_QUOTA_EXCEEDED",
        message:
          "Workbook assistant is already running too many turns for this user. Retry once an in-flight turn finishes.",
        statusCode: 429,
        retryable: true,
      });
    }
    const activeTurnsForDocument = activeSessions.filter(
      (sessionState) => sessionState.documentId === documentId,
    ).length;
    if (activeTurnsForDocument >= this.maxActiveTurnsPerDocument) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_DOCUMENT_TURN_QUOTA_EXCEEDED",
        message:
          "Workbook assistant is already running too many turns for this document. Retry once an in-flight turn finishes.",
        statusCode: 429,
        retryable: true,
      });
    }
  }

  private assertWorkflowFamilyEnabled(
    workflowTemplate: WorkbookAgentWorkflowRun["workflowTemplate"],
  ): void {
    const workflowFamily = getWorkbookAgentWorkflowFamily(workflowTemplate);
    const familyEnabled =
      workflowFamily === "report"
        ? true
        : workflowFamily === "formula"
          ? this.featureFlags.formulaWorkflowFamilyEnabled
          : workflowFamily === "formatting"
            ? this.featureFlags.formattingWorkflowFamilyEnabled
            : workflowFamily === "import"
              ? this.featureFlags.importWorkflowFamilyEnabled
              : workflowFamily === "rollup"
                ? this.featureFlags.rollupWorkflowFamilyEnabled
                : this.featureFlags.structuralWorkflowFamilyEnabled;
    if (familyEnabled) {
      return;
    }
    throw createWorkbookAgentServiceError({
      code: "WORKBOOK_AGENT_WORKFLOW_FAMILY_DISABLED",
      message: `Workbook assistant ${workflowFamily} workflows are currently disabled.`,
      statusCode: 409,
      retryable: false,
    });
  }

  private evictIfNeeded(): void {
    if (this.sessions.size <= this.maxSessions) {
      return;
    }
    const candidates = [...this.sessions.values()]
      .filter((sessionState) => {
        const listeners = this.subscribers.get(sessionState.threadId);
        return sessionState.snapshot.status === "idle" && (!listeners || listeners.size === 0);
      })
      .toSorted((left, right) => left.lastAccessedAt - right.lastAccessedAt);
    while (this.sessions.size > this.maxSessions && candidates.length > 0) {
      const evicted = candidates.shift();
      if (!evicted) {
        return;
      }
      this.codexClient?.releaseThread(evicted.threadId);
      this.sessions.delete(evicted.sessionId);
      this.threadToSessionId.delete(evicted.threadId);
      this.subscribers.delete(evicted.threadId);
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
