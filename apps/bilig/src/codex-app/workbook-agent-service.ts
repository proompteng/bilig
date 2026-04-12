import type {
  WorkbookAgentThreadSnapshot,
  WorkbookAgentStreamEvent,
  WorkbookAgentThreadSummary,
  WorkbookAgentWorkflowRun,
} from "@bilig/contracts";
import type {
  WorkbookAgentAppliedBy,
  WorkbookAgentCommandBundle,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewSummary,
  WorkbookAgentReviewQueueItem,
  WorkbookAgentSharedReviewState,
} from "@bilig/agent-api";
import {
  buildWorkbookAgentPreview,
  buildWorkbookAgentExecutionRecord,
  createWorkbookAgentCommandBundle,
  decodeWorkbookAgentPreviewSummary,
  describeWorkbookAgentBundle,
  isWorkbookAgentBundleAutoApplyEligible,
  requiresWorkbookAgentOwnerReview,
  resolveWorkbookAgentBundleExecutionPolicyInput,
  splitWorkbookAgentCommandBundle,
  toWorkbookAgentCommandBundle,
  toWorkbookAgentReviewQueueItem,
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
import { workbookAgentDynamicToolSpecs } from "./workbook-agent-tools.js";
import {
  buildEntriesFromThread,
  createSessionBodySchema,
  createSystemEntry,
  createWorkbookAgentBaseInstructions,
  createWorkbookAgentDeveloperInstructions,
  reviewPendingBundleBodySchema,
  startWorkflowBodySchema,
  startTurnBodySchema,
  updateContextBodySchema,
} from "./workbook-agent-session-model.js";
import {
  appendRevisionCitation,
  attachSharedReviewState,
  createBundleRangeCitations,
  createPendingSharedReviewState,
  createWorkflowTurnId,
  needsSharedOwnerReview,
  normalizeSharedReviewState,
} from "./workbook-agent-bundle-state.js";
import { routeWorkbookAgentCodexNotification } from "./workbook-agent-codex-notification-router.js";
import { createWorkbookAgentDynamicToolHandler } from "./workbook-agent-dynamic-tool-handler.js";
import { WorkbookAgentThreadRepository } from "./workbook-agent-thread-repository.js";
import {
  buildSnapshot,
  cloneUiContext,
  isMutatingWorkflowTemplate,
  mergeTimelineEntries,
  normalizeExecutionPolicy,
  type WorkbookAgentThreadState,
  toContextRef,
  upsertEntry,
} from "./workbook-agent-service-shared.js";
import { WorkbookAgentWorkflowRuntime } from "./workbook-agent-workflow-runtime.js";

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

export interface WorkbookAgentService {
  readonly enabled: boolean;
  createSession(input: {
    documentId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot>;
  updateContext(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot>;
  startTurn(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot>;
  startWorkflow(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot>;
  cancelWorkflow(input: {
    documentId: string;
    threadId: string;
    runId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSnapshot>;
  interruptTurn(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSnapshot>;
  applyReviewItem(input: {
    documentId: string;
    threadId: string;
    reviewItemId: string;
    session: SessionIdentity;
    appliedBy: WorkbookAgentAppliedBy;
    commandIndexes?: readonly number[] | null;
    preview: unknown;
  }): Promise<WorkbookAgentThreadSnapshot>;
  reviewReviewItem(input: {
    documentId: string;
    threadId: string;
    reviewItemId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot>;
  dismissReviewItem(input: {
    documentId: string;
    threadId: string;
    reviewItemId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSnapshot>;
  replayExecutionRecord(input: {
    documentId: string;
    threadId: string;
    recordId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSnapshot>;
  listThreads(input: {
    documentId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSummary[]>;
  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot;
  getSnapshot(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
  }): WorkbookAgentThreadSnapshot;
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

  async applyReviewItem(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async reviewReviewItem(): Promise<never> {
    throw new Error("Workbook agent service is not configured");
  }

  async dismissReviewItem(): Promise<never> {
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
  private readonly sessions = new Map<string, WorkbookAgentThreadState>();
  private readonly subscribers = new Map<string, Set<(event: WorkbookAgentStreamEvent) => void>>();
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
  private readonly threadRepository: WorkbookAgentThreadRepository;
  private readonly workflowRuntime: WorkbookAgentWorkflowRuntime;

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
    this.threadRepository = new WorkbookAgentThreadRepository(this.zeroSyncService);
    this.workflowRuntime = new WorkbookAgentWorkflowRuntime({
      zeroSyncService: this.zeroSyncService,
      now: this.now,
      touch: (sessionState) => this.touch(sessionState),
      persistSessionState: (sessionState) => this.persistSessionState(sessionState),
      emitSnapshot: (threadId) => this.emitSnapshot(threadId),
      shouldApplyBundleImmediately: (sessionState, bundle) =>
        this.shouldApplyToolBundleImmediately(sessionState, bundle),
      stageReviewBundle: (sessionState, turnId, bundle) =>
        this.stageReviewBundle(sessionState, turnId, attachSharedReviewState(bundle, sessionState)),
      applyCommandBundleAutomatically: (args) => this.applyToolBundleAutomatically(args),
      incrementCounter: (counter) => {
        this.counters[counter] += 1;
      },
    });
  }

  private async persistSessionState(sessionState: WorkbookAgentThreadState): Promise<void> {
    await this.threadRepository.saveThreadState({
      documentId: sessionState.documentId,
      threadId: sessionState.threadId,
      actorUserId: sessionState.storageActorUserId,
      scope: sessionState.scope,
      executionPolicy: sessionState.executionPolicy,
      context: sessionState.durable.context,
      entries: sessionState.durable.entries,
      reviewQueueItems: sessionState.durable.reviewQueueItems,
      updatedAtUnixMs: this.now(),
    });
  }

  private getCurrentReviewItem(
    sessionState: WorkbookAgentThreadState,
  ): WorkbookAgentReviewQueueItem | null {
    return sessionState.durable.reviewQueueItems[0] ?? null;
  }

  private replaceCurrentReviewItem(
    sessionState: WorkbookAgentThreadState,
    reviewItem: WorkbookAgentReviewQueueItem | null,
  ): void {
    sessionState.durable.reviewQueueItems = reviewItem ? [reviewItem] : [];
  }

  private stageReviewBundle(
    sessionState: WorkbookAgentThreadState,
    turnId: string,
    bundle: WorkbookAgentCommandBundle,
  ): void {
    this.replaceCurrentReviewItem(
      sessionState,
      toWorkbookAgentReviewQueueItem({
        bundle,
        reviewMode: needsSharedOwnerReview(sessionState, bundle) ? "ownerReview" : "manual",
        sharedReview: bundle.sharedReview ?? null,
      }),
    );
    sessionState.durable.entries = upsertEntry(
      sessionState.durable.entries,
      createSystemEntry(
        `system-preview:${bundle.id}`,
        turnId,
        describeWorkbookAgentBundle(bundle),
        createBundleRangeCitations(bundle),
      ),
    );
  }

  private queuePrivateTurnBundle(
    sessionState: WorkbookAgentThreadState,
    turnId: string,
    bundle: WorkbookAgentCommandBundle,
  ): void {
    sessionState.live.stagedPrivateBundleByTurn.set(turnId, bundle);
  }

  private shouldApplyToolBundleImmediately(
    sessionState: WorkbookAgentThreadState,
    bundle: WorkbookAgentCommandBundle,
  ): boolean {
    if (
      sessionState.executionPolicy === "autoApplySafe" &&
      !this.featureFlags.autoApplyLowRiskEnabled
    ) {
      return false;
    }
    if (!this.isRolloutAllowed(sessionState.documentId, sessionState.storageActorUserId)) {
      return false;
    }
    return isWorkbookAgentBundleAutoApplyEligible(
      resolveWorkbookAgentBundleExecutionPolicyInput({
        scope: sessionState.scope,
        executionPolicy: sessionState.executionPolicy,
        bundle,
      }),
    );
  }

  private async buildAuthoritativePreview(documentId: string, bundle: WorkbookAgentCommandBundle) {
    return this.zeroSyncService.inspectWorkbook(documentId, async (runtime) =>
      buildWorkbookAgentPreview({
        snapshot: runtime.engine.exportSnapshot(),
        replicaId: `server:${runtime.documentId}:agent-preview`,
        bundle,
      }),
    );
  }

  private async applyPendingBundleForSessionState(input: {
    sessionState: WorkbookAgentThreadState;
    pendingBundle: WorkbookAgentCommandBundle;
    actorUserId: string;
    appliedBy: WorkbookAgentAppliedBy;
    commandIndexes?: readonly number[] | null | undefined;
    preview: WorkbookAgentPreviewSummary;
  }): Promise<WorkbookAgentExecutionRecord> {
    const selection = splitWorkbookAgentCommandBundle({
      bundle: input.pendingBundle,
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
        message: "Automatic apply runs one complete change set per turn.",
        statusCode: 409,
        retryable: false,
      });
    }
    if (input.appliedBy === "auto") {
      if (
        !isWorkbookAgentBundleAutoApplyEligible(
          resolveWorkbookAgentBundleExecutionPolicyInput({
            scope: input.sessionState.scope,
            executionPolicy: input.sessionState.executionPolicy,
            bundle: selection.acceptedBundle,
          }),
        )
      ) {
        throw createWorkbookAgentServiceError({
          code: "WORKBOOK_AGENT_MANUAL_APPROVAL_REQUIRED",
          message: "This session routes workbook edits through the review queue.",
          statusCode: 409,
          retryable: false,
        });
      }
      if (
        input.sessionState.executionPolicy === "autoApplySafe" &&
        !this.featureFlags.autoApplyLowRiskEnabled
      ) {
        throw createWorkbookAgentServiceError({
          code: "WORKBOOK_AGENT_AUTO_APPLY_DISABLED",
          message: "Automatic safe-apply is paused for this environment.",
          statusCode: 409,
          retryable: false,
        });
      }
      this.assertRolloutAllowed({
        documentId: input.sessionState.documentId,
        userId: input.actorUserId,
        code: "WORKBOOK_AGENT_AUTO_APPLY_ROLLOUT_BLOCKED",
        message: "Automatic apply is limited to the rollout allowlist for this environment.",
      });
    }
    if (
      requiresWorkbookAgentOwnerReview({
        scope: input.sessionState.scope,
        riskClass: input.pendingBundle.riskClass,
      }) &&
      input.sessionState.storageActorUserId !== input.actorUserId
    ) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SHARED_APPROVAL_REQUIRED",
        message: "Shared medium/high-risk workbook bundles must be applied by the thread owner.",
        statusCode: 409,
        retryable: false,
      });
    }
    const sharedReview = normalizeSharedReviewState(input.pendingBundle, input.sessionState);
    if (
      requiresWorkbookAgentOwnerReview({
        scope: input.sessionState.scope,
        riskClass: input.pendingBundle.riskClass,
      }) &&
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
      input.sessionState.documentId,
      selection.acceptedBundle,
      input.preview,
      {
        userID: input.actorUserId,
        roles: ["editor"],
      },
    );
    const executionRecord = buildWorkbookAgentExecutionRecord({
      bundle: selection.acceptedBundle,
      actorUserId: input.actorUserId,
      planText: this.collectPlanTextForTurn(input.sessionState, input.pendingBundle.turnId),
      preview: result.preview,
      appliedRevision: result.revision,
      appliedBy: input.appliedBy,
      acceptedScope: selection.acceptedScope,
      now: this.now(),
    });
    await this.zeroSyncService.appendWorkbookAgentRun(executionRecord);
    input.sessionState.durable.executionRecords = [
      executionRecord,
      ...input.sessionState.durable.executionRecords.filter(
        (record) => record.id !== executionRecord.id,
      ),
    ];
    this.replaceCurrentReviewItem(
      input.sessionState,
      selection.remainingBundle === null
        ? null
        : toWorkbookAgentReviewQueueItem({
            bundle: attachSharedReviewState(
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
              input.sessionState,
            ),
            reviewMode:
              input.sessionState.scope === "shared" && selection.remainingBundle.riskClass !== "low"
                ? "ownerReview"
                : "manual",
          }),
    );
    input.sessionState.durable.entries = upsertEntry(
      input.sessionState.durable.entries,
      createSystemEntry(
        `system-apply:${executionRecord.id}`,
        input.pendingBundle.turnId,
        `${input.appliedBy === "auto" ? "Applied automatically" : "Applied"} ${
          selection.acceptedScope === "partial" ? "selected " : ""
        }workbook change set at revision r${String(result.revision)}: ${selection.acceptedBundle.summary}`,
        appendRevisionCitation(
          createBundleRangeCitations(selection.acceptedBundle),
          result.revision,
        ),
      ),
    );
    this.touch(input.sessionState);
    return executionRecord;
  }

  private async applyToolBundleAutomatically(input: {
    sessionState: WorkbookAgentThreadState;
    actorUserId: string;
    bundle: WorkbookAgentCommandBundle;
  }): Promise<WorkbookAgentExecutionRecord | null> {
    this.replaceCurrentReviewItem(
      input.sessionState,
      toWorkbookAgentReviewQueueItem({
        bundle: input.bundle,
        reviewMode: "manual",
      }),
    );
    const preview = await this.buildAuthoritativePreview(
      input.sessionState.documentId,
      input.bundle,
    );
    const executionRecord = await this.applyPendingBundleForSessionState({
      sessionState: input.sessionState,
      pendingBundle: input.bundle,
      actorUserId: input.actorUserId,
      appliedBy: "auto",
      preview,
    });
    await this.persistSessionState(input.sessionState);
    this.emitSnapshot(input.sessionState.threadId);
    return executionRecord.bundleId === input.bundle.id ? executionRecord : null;
  }

  private async finalizePrivateTurnBundle(input: {
    sessionState: WorkbookAgentThreadState;
    turnId: string;
    turnStatus: "completed" | "failed";
  }): Promise<void> {
    const queuedBundle = input.sessionState.live.stagedPrivateBundleByTurn.get(input.turnId);
    if (!queuedBundle) {
      return;
    }
    input.sessionState.live.stagedPrivateBundleByTurn.delete(input.turnId);
    if (input.turnStatus !== "completed") {
      return;
    }
    const actorUserId = this.resolveTurnActorUserId(input.sessionState, input.turnId);
    try {
      const preview = await this.buildAuthoritativePreview(
        input.sessionState.documentId,
        queuedBundle,
      );
      await this.applyPendingBundleForSessionState({
        sessionState: input.sessionState,
        pendingBundle: queuedBundle,
        actorUserId,
        appliedBy: "auto",
        preview,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.stageReviewBundle(
        input.sessionState,
        input.turnId,
        attachSharedReviewState(queuedBundle, input.sessionState),
      );
      input.sessionState.live.lastError = message;
      input.sessionState.durable.entries = upsertEntry(
        input.sessionState.durable.entries,
        createSystemEntry(
          `system-review-fallback:${queuedBundle.id}:${this.now()}`,
          input.turnId,
          `Prepared workbook review item after turn apply failed: ${queuedBundle.summary}`,
          createBundleRangeCitations(queuedBundle),
        ),
      );
      input.sessionState.live.status = "failed";
    }
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
        activeTurnCount: sessions.filter((sessionState) => sessionState.live.activeTurnId !== null)
          .length,
        runningWorkflowCount: sessions.reduce(
          (sum, sessionState) =>
            sum +
            sessionState.durable.workflowRuns.filter((run) => run.status === "running").length,
          0,
        ),
        pendingBundleCount: sessions.filter(
          (sessionState) => sessionState.durable.reviewQueueItems.length > 0,
        ).length,
        sharedPendingReviewCount: sessions.filter((sessionState) => {
          const pendingReviewItem = this.getCurrentReviewItem(sessionState);
          return (
            pendingReviewItem !== null &&
            pendingReviewItem.reviewMode === "ownerReview" &&
            pendingReviewItem.status === "pending"
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
  }): Promise<WorkbookAgentThreadSnapshot> {
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
    if (parsed.threadId !== undefined) {
      const sharedSession = this.tryGetSessionByThreadId(parsed.threadId);
      if (sharedSession) {
        const accessibleSession = this.requireOwnedSession(
          sharedSession,
          input.documentId,
          input.session.userID,
        );
        if (parsed.context) {
          accessibleSession.durable.context = parsed.context;
        }
        if (parsed.executionPolicy) {
          accessibleSession.executionPolicy = normalizeExecutionPolicy({
            scope: accessibleSession.scope,
            requestedPolicy: parsed.executionPolicy,
          });
        }
        if (parsed.context || parsed.executionPolicy) {
          await this.persistSessionState(accessibleSession);
          this.emitSnapshot(accessibleSession.threadId);
        }
        this.touch(accessibleSession);
        return buildSnapshot(accessibleSession);
      }
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
    const durableThreadSession = await this.threadRepository.loadThreadState({
      documentId: input.documentId,
      actorUserId: input.session.userID,
      threadId,
    });
    const durableThreadState = durableThreadSession.threadState;
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
    const resolvedExecutionPolicy = normalizeExecutionPolicy({
      scope: resolvedScope,
      requestedPolicy: parsed.executionPolicy ?? durableThreadState?.executionPolicy ?? null,
    });
    const executionRecords = durableThreadSession.executionRecords;
    const workflowRuns = durableThreadSession.workflowRuns;
    const codexEntries = thread ? buildEntriesFromThread(thread) : [];
    const bootstrapErrorMessage =
      thread || !sessionBootstrapError
        ? null
        : sessionBootstrapError instanceof Error
          ? sessionBootstrapError.message
          : "Workbook assistant live session is unavailable. Loaded durable thread history only.";

    const sessionState: WorkbookAgentThreadState = {
      documentId: input.documentId,
      userId: input.session.userID,
      storageActorUserId: durableThreadState?.actorUserId ?? input.session.userID,
      scope: resolvedScope,
      executionPolicy: resolvedExecutionPolicy,
      threadId,
      durable: {
        context: parsed.context ?? durableThreadState?.context ?? null,
        entries: mergeTimelineEntries(codexEntries, durableThreadState?.entries ?? []),
        reviewQueueItems: [...(durableThreadState?.reviewQueueItems ?? [])],
        executionRecords,
        workflowRuns,
      },
      live: {
        activeTurnId: thread?.turns.findLast((turn) => turn.status === "inProgress")?.id ?? null,
        status: !thread
          ? "failed"
          : thread.turns.some((turn) => turn.status === "failed")
            ? "failed"
            : thread.turns.some((turn) => turn.status === "inProgress")
              ? "inProgress"
              : "idle",
        lastError:
          thread?.turns.findLast((turn) => turn.error?.message)?.error?.message ??
          bootstrapErrorMessage,
        stagedPrivateBundleByTurn: new Map(),
        optimisticUserEntryIdByTurn: new Map(),
        promptByTurn: new Map(),
        turnActorUserIdByTurn: new Map(),
        turnContextByTurn: new Map(),
        lastAccessedAt: this.now(),
      },
    };
    const canApplyBootstrapBundle =
      resolvedScope === "private" &&
      sessionState.durable.reviewQueueItems.length > 0 &&
      this.isRolloutAllowed(input.documentId, input.session.userID) &&
      (sessionState.executionPolicy === "autoApplyAll" ||
        (sessionState.executionPolicy === "autoApplySafe" &&
          this.featureFlags.autoApplyLowRiskEnabled &&
          (this.getCurrentReviewItem(sessionState)?.riskClass ?? "high") === "low"));
    const bootstrapReviewItem = this.getCurrentReviewItem(sessionState);
    if (canApplyBootstrapBundle && bootstrapReviewItem) {
      const queuedBundle = toWorkbookAgentCommandBundle(bootstrapReviewItem);
      const currentRevision = await this.zeroSyncService.getWorkbookHeadRevision(input.documentId);
      const migratedBundle =
        queuedBundle.baseRevision === currentRevision
          ? queuedBundle
          : createWorkbookAgentCommandBundle({
              bundleId: queuedBundle.id,
              documentId: queuedBundle.documentId,
              threadId: queuedBundle.threadId,
              turnId: queuedBundle.turnId,
              goalText: queuedBundle.goalText,
              baseRevision: currentRevision,
              context: queuedBundle.context,
              commands: queuedBundle.commands,
              now: queuedBundle.createdAtUnixMs,
              sharedReview: queuedBundle.sharedReview ?? null,
            });
      this.replaceCurrentReviewItem(
        sessionState,
        toWorkbookAgentReviewQueueItem({
          bundle: migratedBundle,
          reviewMode: bootstrapReviewItem.reviewMode,
          sharedReview:
            bootstrapReviewItem.reviewMode === "ownerReview"
              ? {
                  ownerUserId: bootstrapReviewItem.ownerUserId ?? input.session.userID,
                  status: bootstrapReviewItem.status,
                  decidedByUserId: bootstrapReviewItem.decidedByUserId,
                  decidedAtUnixMs: bootstrapReviewItem.decidedAtUnixMs,
                  recommendations: [...bootstrapReviewItem.recommendations],
                }
              : null,
        }),
      );
      const preview = await this.buildAuthoritativePreview(input.documentId, migratedBundle);
      await this.applyPendingBundleForSessionState({
        sessionState,
        pendingBundle: migratedBundle,
        actorUserId: input.session.userID,
        appliedBy: "auto",
        preview,
      });
    }
    this.sessions.set(threadId, sessionState);
    this.evictIfNeeded();
    await this.persistSessionState(sessionState);
    return buildSnapshot(sessionState);
  }

  async updateContext(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = updateContextBodySchema.parse(input.body);
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    sessionState.durable.context = parsed.context;
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return buildSnapshot(sessionState);
  }

  async startTurn(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = startTurnBodySchema.parse(input.body);
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    if (sessionState.live.activeTurnId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_TURN_ALREADY_RUNNING",
        message: "Finish or interrupt the current assistant turn before starting another one.",
        statusCode: 409,
        retryable: false,
      });
    }
    this.assertTurnQuota(input.documentId, input.session.userID);
    if (parsed.context) {
      sessionState.durable.context = parsed.context;
    }
    const turnContext = cloneUiContext(parsed.context ?? sessionState.durable.context);
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
    sessionState.durable.entries = upsertEntry(sessionState.durable.entries, {
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
    sessionState.live.optimisticUserEntryIdByTurn.set(turn.id, optimisticEntryId);
    sessionState.live.promptByTurn.set(turn.id, parsed.prompt);
    sessionState.live.turnActorUserIdByTurn.set(turn.id, input.session.userID);
    sessionState.live.turnContextByTurn.set(turn.id, turnContext);
    sessionState.live.activeTurnId = turn.id;
    sessionState.live.status = "inProgress";
    sessionState.live.lastError = null;
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return buildSnapshot(sessionState);
  }

  async startWorkflow(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = startWorkflowBodySchema.parse(input.body);
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
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
      sessionState.durable.context = parsed.context;
    }
    const runningWorkflow = sessionState.durable.workflowRuns.find(
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
    const workflowTemplate = parsed.workflowTemplate;
    this.assertWorkflowFamilyEnabled(workflowTemplate);
    if (
      sessionState.durable.reviewQueueItems.length > 0 &&
      isMutatingWorkflowTemplate(workflowTemplate)
    ) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_PENDING_BUNDLE_EXISTS",
        message:
          "Finish the current workbook review item before starting another mutating workflow.",
        statusCode: 409,
        retryable: false,
      });
    }
    const workflowInput = {
      ...("query" in parsed ? { query: parsed.query } : {}),
      ...("sheetName" in parsed && parsed.sheetName ? { sheetName: parsed.sheetName } : {}),
      ...("limit" in parsed && parsed.limit !== undefined ? { limit: parsed.limit } : {}),
      ...("name" in parsed ? { name: parsed.name } : {}),
    };
    const workflowTurnId =
      sessionState.live.activeTurnId ?? createWorkflowTurnId(crypto.randomUUID());
    await this.workflowRuntime.startWorkflow({
      sessionState,
      documentId: input.documentId,
      actorUserId: input.session.userID,
      workflowTemplate,
      workflowInput,
      workflowTurnId,
    });
    return buildSnapshot(sessionState);
  }

  async cancelWorkflow(input: {
    documentId: string;
    threadId: string;
    runId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    const runningWorkflow = sessionState.durable.workflowRuns.find(
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
    await this.workflowRuntime.cancelRunningWorkflow({
      sessionState,
      documentId: input.documentId,
      runId: input.runId,
      runningWorkflow,
      actorUserId: input.session.userID,
    });
    return buildSnapshot(sessionState);
  }

  async interruptTurn(input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    const codexClient = await this.getCodexClient();
    await codexClient.turnInterrupt(sessionState.threadId);
    return buildSnapshot(sessionState);
  }

  async applyReviewItem(input: {
    documentId: string;
    threadId: string;
    reviewItemId: string;
    session: SessionIdentity;
    appliedBy: WorkbookAgentAppliedBy;
    commandIndexes?: readonly number[] | null;
    preview: unknown;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    const reviewItem = this.getCurrentReviewItem(sessionState);
    if (!reviewItem || reviewItem.id !== input.reviewItemId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_BUNDLE_NOT_FOUND",
        message: "Workbook agent change set was not found.",
        statusCode: 404,
        retryable: false,
      });
    }
    const preview = decodeWorkbookAgentPreviewSummary(input.preview);
    if (!preview) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_PREVIEW_REQUIRED",
        message: "Workbook preview details are required before applying this change set.",
        statusCode: 400,
        retryable: false,
      });
    }
    await this.applyPendingBundleForSessionState({
      sessionState,
      pendingBundle: toWorkbookAgentCommandBundle(reviewItem),
      actorUserId: input.session.userID,
      appliedBy: input.appliedBy,
      commandIndexes: input.commandIndexes,
      preview,
    });
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return buildSnapshot(sessionState);
  }

  async reviewReviewItem(input: {
    documentId: string;
    threadId: string;
    reviewItemId: string;
    session: SessionIdentity;
    body: unknown;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = reviewPendingBundleBodySchema.parse(input.body);
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    const reviewItem = this.getCurrentReviewItem(sessionState);
    if (!reviewItem || reviewItem.id !== input.reviewItemId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_BUNDLE_NOT_FOUND",
        message: "Workbook review item was not found.",
        statusCode: 404,
        retryable: false,
      });
    }
    if (reviewItem.reviewMode !== "ownerReview") {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_SHARED_REVIEW_NOT_REQUIRED",
        message: "Shared review is only required for medium/high-risk bundles.",
        statusCode: 409,
        retryable: false,
      });
    }
    const now = this.now();
    const sharedReview =
      normalizeSharedReviewState(toWorkbookAgentCommandBundle(reviewItem), sessionState) ??
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
    const reviewedBundle = {
      ...toWorkbookAgentCommandBundle(reviewItem),
      sharedReview: nextSharedReview,
    } satisfies WorkbookAgentCommandBundle;
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
    this.replaceCurrentReviewItem(
      sessionState,
      toWorkbookAgentReviewQueueItem({
        bundle: reviewedBundle,
        reviewMode: "ownerReview",
        sharedReview: nextSharedReview,
      }),
    );
    sessionState.durable.entries = upsertEntry(
      sessionState.durable.entries,
      createSystemEntry(
        `system-review:${reviewedBundle.id}:${now}`,
        reviewedBundle.turnId,
        isOwnerReviewer
          ? `${parsed.decision === "approved" ? "Approved" : "Returned"} shared review item: ${reviewedBundle.summary}`
          : `${input.session.userID} shared a ${parsed.decision === "approved" ? "ready-to-apply" : "return-for-edit"} review recommendation: ${reviewedBundle.summary}`,
        createBundleRangeCitations(reviewedBundle),
      ),
    );
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return buildSnapshot(sessionState);
  }

  async dismissReviewItem(input: {
    documentId: string;
    threadId: string;
    reviewItemId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    const reviewItem = this.getCurrentReviewItem(sessionState);
    if (!reviewItem || reviewItem.id !== input.reviewItemId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_BUNDLE_NOT_FOUND",
        message: "Workbook review item was not found.",
        statusCode: 404,
        retryable: false,
      });
    }
    this.replaceCurrentReviewItem(sessionState, null);
    sessionState.durable.entries = upsertEntry(
      sessionState.durable.entries,
      createSystemEntry(
        `system-dismiss:${reviewItem.id}:${this.now()}`,
        reviewItem.turnId,
        `Cleared workbook review item: ${reviewItem.summary}`,
        createBundleRangeCitations(toWorkbookAgentCommandBundle(reviewItem)),
      ),
    );
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return buildSnapshot(sessionState);
  }

  async replayExecutionRecord(input: {
    documentId: string;
    threadId: string;
    recordId: string;
    session: SessionIdentity;
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    const record = sessionState.durable.executionRecords.find(
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
      context: toContextRef(sessionState.durable.context) ?? record.context,
      commands: record.commands,
      now: this.now(),
    });
    if (this.shouldApplyToolBundleImmediately(sessionState, replayedBundle)) {
      this.replaceCurrentReviewItem(
        sessionState,
        toWorkbookAgentReviewQueueItem({
          bundle: replayedBundle,
          reviewMode: "manual",
        }),
      );
      const preview = await this.buildAuthoritativePreview(input.documentId, replayedBundle);
      await this.applyPendingBundleForSessionState({
        sessionState,
        pendingBundle: replayedBundle,
        actorUserId: input.session.userID,
        appliedBy: "auto",
        preview,
      });
      await this.persistSessionState(sessionState);
      this.emitSnapshot(sessionState.threadId);
      return buildSnapshot(sessionState);
    }
    this.replaceCurrentReviewItem(
      sessionState,
      toWorkbookAgentReviewQueueItem({
        bundle: replayedBundle,
        reviewMode: needsSharedOwnerReview(sessionState, replayedBundle) ? "ownerReview" : "manual",
        sharedReview: replayedBundle.sharedReview ?? null,
      }),
    );
    sessionState.durable.entries = upsertEntry(
      sessionState.durable.entries,
      createSystemEntry(
        `system-replay:${record.id}:${String(this.now())}`,
        replayedBundle.turnId,
        `Prepared workbook review item from a prior execution: ${replayedBundle.summary}`,
        createBundleRangeCitations(replayedBundle),
      ),
    );
    this.touch(sessionState);
    await this.persistSessionState(sessionState);
    this.emitSnapshot(sessionState.threadId);
    return buildSnapshot(sessionState);
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
    threadId: string;
    session: SessionIdentity;
  }): WorkbookAgentThreadSnapshot {
    const sessionState = this.getOwnedSession(
      input.documentId,
      input.threadId,
      input.session.userID,
    );
    this.touch(sessionState);
    return buildSnapshot(sessionState);
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
    this.workflowRuntime.close();
    await this.codexClient?.close();
    this.codexClient = null;
    this.sessions.clear();
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
          handleDynamicToolCall: createWorkbookAgentDynamicToolHandler({
            zeroSyncService: this.zeroSyncService,
            now: this.now,
            getSessionByThreadId: (threadId) => this.getSessionByThreadId(threadId),
            resolveTurnActorUserId: (sessionState, turnId) =>
              this.resolveTurnActorUserId(sessionState, turnId),
            resolveTurnContext: (sessionState, turnId) =>
              this.resolveTurnContext(sessionState, turnId),
            stageReviewBundle: (sessionState, turnId, bundle) =>
              this.stageReviewBundle(
                sessionState,
                turnId,
                attachSharedReviewState(bundle, sessionState),
              ),
            queuePrivateTurnBundle: (sessionState, turnId, bundle) =>
              this.queuePrivateTurnBundle(sessionState, turnId, bundle),
            shouldApplyToolBundleImmediately: (sessionState, bundle) =>
              this.shouldApplyToolBundleImmediately(sessionState, bundle),
            applyToolBundleAutomatically: (args) => this.applyToolBundleAutomatically(args),
            persistSessionState: (sessionState) => this.persistSessionState(sessionState),
            emitSnapshot: (threadId) => this.emitSnapshot(threadId),
            startWorkflow: (request) => this.startWorkflow(request),
          }),
        },
      });
      await this.codexClient.ensureReady();
      this.unsubscribeCodex = this.codexClient.subscribe((notification) => {
        void routeWorkbookAgentCodexNotification({
          notification,
          listSessions: () => [...this.sessions.values()],
          tryGetSessionByThreadId: (threadId) => this.tryGetSessionByThreadId(threadId),
          finalizeCompletedTurn: async (sessionState, turnId, turnStatus) =>
            await this.finalizePrivateTurnBundle({
              sessionState,
              turnId,
              turnStatus,
            }),
          persistSessionState: (sessionState) => this.persistSessionState(sessionState),
          emitSnapshot: (threadId) => this.emitSnapshot(threadId),
          emit: (threadId, event) => this.emit(threadId, event),
          now: this.now,
        }).catch((error: unknown) => {
          console.error(error);
        });
      });
    }
    return this.codexClient;
  }

  private resolveTurnActorUserId(sessionState: WorkbookAgentThreadState, turnId: string): string {
    return sessionState.live.turnActorUserIdByTurn.get(turnId) ?? sessionState.userId;
  }

  private resolveTurnContext(sessionState: WorkbookAgentThreadState, turnId: string) {
    return cloneUiContext(
      sessionState.live.turnContextByTurn.get(turnId) ?? sessionState.durable.context,
    );
  }

  private collectPlanTextForTurn(
    sessionState: WorkbookAgentThreadState,
    turnId: string,
  ): string | null {
    const planText = sessionState.durable.entries
      .filter((entry) => entry.turnId === turnId && entry.kind === "plan" && entry.text)
      .map((entry) => entry.text?.trim() ?? "")
      .filter((text) => text.length > 0)
      .join("\n\n");
    return planText.length > 0 ? planText : null;
  }

  private getOwnedSession(
    documentId: string,
    threadId: string,
    userId: string,
  ): WorkbookAgentThreadState {
    const sessionState = this.sessions.get(threadId);
    if (!sessionState) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_THREAD_NOT_FOUND",
        message: "Workbook agent thread not found",
        statusCode: 404,
        retryable: true,
      });
    }
    return this.requireOwnedSession(sessionState, documentId, userId);
  }

  private requireOwnedSession(
    sessionState: WorkbookAgentThreadState,
    documentId: string,
    userId: string,
  ): WorkbookAgentThreadState {
    if (sessionState.documentId !== documentId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_THREAD_NOT_FOUND",
        message: "Workbook agent thread not found",
        statusCode: 404,
        retryable: false,
      });
    }
    if (sessionState.scope !== "shared" && sessionState.userId !== userId) {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_THREAD_NOT_FOUND",
        message: "Workbook agent thread not found",
        statusCode: 404,
        retryable: false,
      });
    }
    return sessionState;
  }

  private getSessionByThreadId(threadId: string): WorkbookAgentThreadState {
    const sessionState = this.tryGetSessionByThreadId(threadId);
    if (!sessionState) {
      throw new Error(`Workbook agent thread not found for thread ${threadId}`);
    }
    return sessionState;
  }

  private tryGetSessionByThreadId(threadId: string): WorkbookAgentThreadState | null {
    return this.sessions.get(threadId) ?? null;
  }

  private emitSnapshot(threadId: string): void {
    const sessionState = this.tryGetSessionByThreadId(threadId);
    if (!sessionState) {
      return;
    }
    this.emit(threadId, {
      type: "snapshot",
      snapshot: buildSnapshot(sessionState),
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

  private touch(sessionState: WorkbookAgentThreadState): void {
    sessionState.live.lastAccessedAt = this.now();
  }

  private resolveActiveTurnActorUserId(sessionState: WorkbookAgentThreadState): string | null {
    const activeTurnId = sessionState.live.activeTurnId;
    if (!activeTurnId) {
      return null;
    }
    return sessionState.live.turnActorUserIdByTurn.get(activeTurnId) ?? sessionState.userId;
  }

  private assertTurnQuota(documentId: string, actorUserId: string): void {
    const activeSessions = [...this.sessions.values()].filter(
      (sessionState) =>
        sessionState.live.activeTurnId !== null && sessionState.live.status === "inProgress",
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
        return sessionState.live.status === "idle" && (!listeners || listeners.size === 0);
      })
      .toSorted((left, right) => left.live.lastAccessedAt - right.live.lastAccessedAt);
    while (this.sessions.size > this.maxSessions && candidates.length > 0) {
      const evicted = candidates.shift();
      if (!evicted) {
        return;
      }
      this.codexClient?.releaseThread(evicted.threadId);
      this.sessions.delete(evicted.threadId);
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
