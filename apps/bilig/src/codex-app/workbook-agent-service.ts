import type {
  WorkbookAgentThreadSnapshot,
  WorkbookAgentStreamEvent,
  WorkbookAgentThreadSummary,
  WorkbookAgentWorkflowRun,
} from '@bilig/contracts'
import type {
  WorkbookAgentAppliedBy,
  WorkbookAgentCommandBundle,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewSummary,
} from '@bilig/agent-api'
import {
  createWorkbookAgentCommandBundle,
  decodeWorkbookAgentPreviewSummary,
  isWorkbookAgentBundleAutoApplyEligible,
  resolveWorkbookAgentBundleExecutionPolicyInput,
  toWorkbookAgentCommandBundle,
} from '@bilig/agent-api'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import type { CodexAppServerTransport } from './codex-app-server-client.js'
import { isCodexAppServerPoolBackpressureError } from './codex-app-server-pool.js'
import { DisabledWorkbookAgentService } from './workbook-agent-disabled-service.js'
import {
  getWorkbookAgentWorkflowFamily,
  isWorkbookAgentRolloutAllowed,
  isWorkbookAgentWorkflowFamilyEnabled,
  resolveWorkbookAgentFeatureFlags,
  type WorkbookAgentFeatureFlags,
} from './workbook-agent-feature-flags.js'
import {
  createSessionBodySchema,
  createSystemEntry,
  reviewReviewItemBodySchema,
  startWorkflowBodySchema,
  startTurnBodySchema,
  updateContextBodySchema,
} from './workbook-agent-session-model.js'
import { attachSharedReviewState, createBundleRangeCitations, createWorkflowTurnId } from './workbook-agent-bundle-state.js'
import {
  clearLegacyPrivateBootstrapReviewItem,
  createWorkbookAgentBootstrappedSessionState,
  planWorkbookAgentBootstrapReviewRecovery,
  rebaseWorkbookAgentBootstrapReviewItem,
} from './workbook-agent-service-bootstrap.js'
import {
  WorkbookAgentCodexRuntime,
  createWorkbookAgentThreadResumeInput,
  createWorkbookAgentThreadStartInput,
} from './workbook-agent-codex-runtime.js'
import {
  createWorkbookAgentReviewQueueItem,
  createWorkbookAgentDismissReviewEntry,
  getCurrentWorkbookAgentReviewItem,
  replaceCurrentWorkbookAgentReviewItem,
  requireWorkbookAgentReviewItem,
  stageWorkbookAgentReviewBundle,
  transitionWorkbookAgentSharedReview,
} from './workbook-agent-review-transitions.js'
import { WorkbookAgentThreadRepository } from './workbook-agent-thread-repository.js'
import {
  buildSnapshot,
  cloneUiContext,
  isMutatingWorkflowTemplate,
  normalizeExecutionPolicy,
  type WorkbookAgentThreadState,
  toContextRef,
  upsertEntry,
} from './workbook-agent-service-shared.js'
import { updateWorkbookAgentDurableUiContextFromUser } from './workbook-agent-service-context.js'
import {
  assertWorkbookAgentSessionAccessPolicy,
  assertWorkbookAgentSharedThreadAccess,
  filterWorkbookAgentThreadSummariesByAccessPolicy,
} from './workbook-agent-service-access-policy.js'
import { assertWorkbookAgentTurnQuota } from './workbook-agent-service-session-policy.js'
import { WorkbookAgentWorkflowRuntime } from './workbook-agent-workflow-runtime.js'
import {
  WorkbookAgentSessionRegistry,
  type WorkbookAgentObservabilityCounterName,
  type WorkbookAgentObservabilitySnapshot,
} from './workbook-agent-session-registry.js'
import {
  resolveWorkbookAgentServiceLimits,
  type EnabledWorkbookAgentServiceOptions,
  type WorkbookAgentService,
} from './workbook-agent-service-options.js'
import {
  applyWorkbookAgentCommandBundleForSessionState,
  applyWorkbookAgentToolBundleAutomatically,
  buildWorkbookAgentAuthoritativePreview,
  finalizeWorkbookAgentPrivateTurnBundle,
} from './workbook-agent-service-application.js'

export type { EnabledWorkbookAgentServiceOptions, WorkbookAgentService } from './workbook-agent-service-options.js'

class EnabledWorkbookAgentService implements WorkbookAgentService {
  readonly enabled = true
  private readonly zeroSyncService: ZeroSyncService
  private readonly now: () => number
  private readonly maxSessions: number
  private readonly maxCodexClients: number
  private readonly maxConcurrentTurnsPerCodexClient: number
  private readonly maxQueuedTurnsPerCodexClient: number
  private readonly maxActiveTurnsPerUser: number
  private readonly maxActiveTurnsPerDocument: number
  private readonly featureFlags: WorkbookAgentFeatureFlags
  private readonly sessionRegistry: WorkbookAgentSessionRegistry
  private readonly codexRuntime: WorkbookAgentCodexRuntime
  private readonly threadRepository: WorkbookAgentThreadRepository
  private readonly workflowRuntime: WorkbookAgentWorkflowRuntime

  constructor(options: EnabledWorkbookAgentServiceOptions) {
    const limits = resolveWorkbookAgentServiceLimits(options)
    this.zeroSyncService = options.zeroSyncService
    this.now = options.now ?? (() => Date.now())
    this.maxSessions = limits.maxSessions
    this.maxCodexClients = limits.maxCodexClients
    this.maxConcurrentTurnsPerCodexClient = limits.maxConcurrentTurnsPerCodexClient
    this.maxQueuedTurnsPerCodexClient = limits.maxQueuedTurnsPerCodexClient
    this.maxActiveTurnsPerUser = limits.maxActiveTurnsPerUser
    this.maxActiveTurnsPerDocument = limits.maxActiveTurnsPerDocument
    this.featureFlags = {
      ...resolveWorkbookAgentFeatureFlags(),
      ...options.featureFlags,
    }
    this.sessionRegistry = new WorkbookAgentSessionRegistry({
      maxSessions: this.maxSessions,
      now: this.now,
    })
    this.codexRuntime = new WorkbookAgentCodexRuntime({
      zeroSyncService: this.zeroSyncService,
      now: this.now,
      maxCodexClients: this.maxCodexClients,
      maxConcurrentTurnsPerCodexClient: this.maxConcurrentTurnsPerCodexClient,
      maxQueuedTurnsPerCodexClient: this.maxQueuedTurnsPerCodexClient,
      getSessionByThreadId: (threadId) => this.getSessionByThreadId(threadId),
      tryGetSessionByThreadId: (threadId) => this.tryGetSessionByThreadId(threadId),
      listSessions: () => this.sessionRegistry.listSessions(),
      resolveTurnActorUserId: (sessionState, turnId) => this.resolveTurnActorUserId(sessionState, turnId),
      resolveTurnContext: (sessionState, turnId) => this.resolveTurnContext(sessionState, turnId),
      stageReviewBundle: (sessionState, turnId, bundle) =>
        stageWorkbookAgentReviewBundle({
          sessionState,
          turnId,
          bundle: attachSharedReviewState(bundle, sessionState),
        }),
      shouldApplyToolBundleImmediately: (sessionState, bundle) => this.shouldApplyToolBundleImmediately(sessionState, bundle),
      applyToolBundleAutomatically: (args) => this.applyToolBundleAutomatically(args),
      persistSessionState: (sessionState) => this.persistSessionState(sessionState),
      emitSnapshot: (threadId) => this.sessionRegistry.emitSnapshot(threadId),
      emit: (threadId, event) => this.sessionRegistry.emit(threadId, event),
      finalizeCompletedTurn: async (sessionState, turnId, turnStatus) =>
        await this.finalizePrivateTurnBundle({
          sessionState,
          turnId,
          turnStatus,
        }),
      startWorkflow: async (request) => await this.startWorkflow(request),
      ...(options.codexClientFactory ? { codexClientFactory: options.codexClientFactory } : {}),
    })
    this.threadRepository = new WorkbookAgentThreadRepository(this.zeroSyncService)
    this.workflowRuntime = new WorkbookAgentWorkflowRuntime({
      zeroSyncService: this.zeroSyncService,
      now: this.now,
      touch: (sessionState) => this.sessionRegistry.touch(sessionState),
      persistSessionState: (sessionState) => this.persistSessionState(sessionState),
      emitSnapshot: (threadId) => this.sessionRegistry.emitSnapshot(threadId),
      shouldApplyBundleImmediately: (sessionState, bundle) => this.shouldApplyToolBundleImmediately(sessionState, bundle),
      stageReviewBundle: (sessionState, turnId, bundle) =>
        stageWorkbookAgentReviewBundle({
          sessionState,
          turnId,
          bundle: attachSharedReviewState(bundle, sessionState),
        }),
      applyCommandBundleAutomatically: (args) => this.applyToolBundleAutomatically(args),
      incrementCounter: (counter) => {
        this.sessionRegistry.incrementCounter(counter)
      },
    })
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
    })
  }

  private shouldApplyToolBundleImmediately(sessionState: WorkbookAgentThreadState, bundle: WorkbookAgentCommandBundle): boolean {
    if (sessionState.executionPolicy === 'autoApplySafe' && !this.featureFlags.autoApplyLowRiskEnabled) {
      return false
    }
    if (!this.isRolloutAllowed(sessionState.documentId, sessionState.storageActorUserId)) {
      return false
    }
    return isWorkbookAgentBundleAutoApplyEligible(
      resolveWorkbookAgentBundleExecutionPolicyInput({
        scope: sessionState.scope,
        executionPolicy: sessionState.executionPolicy,
        bundle,
      }),
    )
  }

  private createBundleApplicationContext() {
    return {
      zeroSyncService: this.zeroSyncService,
      now: this.now,
      autoApplyLowRiskEnabled: this.featureFlags.autoApplyLowRiskEnabled,
      isRolloutAllowed: (documentId: string, userId: string) => this.isRolloutAllowed(documentId, userId),
      touchSession: (sessionState: WorkbookAgentThreadState) => this.sessionRegistry.touch(sessionState),
    }
  }

  private async buildAuthoritativePreview(documentId: string, bundle: WorkbookAgentCommandBundle): Promise<WorkbookAgentPreviewSummary> {
    return await buildWorkbookAgentAuthoritativePreview({
      zeroSyncService: this.zeroSyncService,
      documentId,
      bundle,
    })
  }

  private async applyCommandBundleForSessionState(input: {
    sessionState: WorkbookAgentThreadState
    commandBundle: WorkbookAgentCommandBundle
    actorUserId: string
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null | undefined
    preview: WorkbookAgentPreviewSummary
  }): Promise<WorkbookAgentExecutionRecord> {
    return await applyWorkbookAgentCommandBundleForSessionState(this.createBundleApplicationContext(), input)
  }

  private async applyToolBundleAutomatically(input: {
    sessionState: WorkbookAgentThreadState
    actorUserId: string
    bundle: WorkbookAgentCommandBundle
  }): Promise<WorkbookAgentExecutionRecord | null> {
    return await applyWorkbookAgentToolBundleAutomatically(
      {
        ...this.createBundleApplicationContext(),
        persistSessionState: async (sessionState) => await this.persistSessionState(sessionState),
        emitSnapshot: (threadId) => this.sessionRegistry.emitSnapshot(threadId),
      },
      input,
    )
  }

  private async finalizePrivateTurnBundle(input: {
    sessionState: WorkbookAgentThreadState
    turnId: string
    turnStatus: 'completed' | 'failed'
  }): Promise<void> {
    await finalizeWorkbookAgentPrivateTurnBundle(
      {
        ...this.createBundleApplicationContext(),
        resolveTurnActorUserId: (sessionState, turnId) => this.resolveTurnActorUserId(sessionState, turnId),
      },
      input,
    )
  }

  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot {
    const runtimePoolStats = this.codexRuntime.getStats()
    const poolStats = runtimePoolStats ?? {
      slotCount: 0,
      boundThreadCount: 0,
      activeTurnCount: 0,
      queuedTurnCount: 0,
      maxClients: this.maxCodexClients,
      maxConcurrentTurnsPerClient: this.maxConcurrentTurnsPerCodexClient,
      maxQueuedTurnsPerClient: this.maxQueuedTurnsPerCodexClient,
    }
    return this.sessionRegistry.getObservabilitySnapshot({
      featureFlags: this.featureFlags,
      poolStats,
    })
  }

  private isRolloutAllowed(documentId: string, userId: string): boolean {
    return isWorkbookAgentRolloutAllowed(this.featureFlags, { documentId, userId })
  }

  private assertRolloutAllowed(input: { documentId: string; userId: string; code: string; message: string }): void {
    if (this.isRolloutAllowed(input.documentId, input.userId)) {
      return
    }
    throw createWorkbookAgentServiceError({
      code: input.code,
      message: input.message,
      statusCode: 409,
      retryable: false,
    })
  }

  async createSession(input: { documentId: string; session: SessionIdentity; body: unknown }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = createSessionBodySchema.parse(input.body)
    if (parsed.scope === 'shared') {
      assertWorkbookAgentSharedThreadAccess({
        featureFlags: this.featureFlags,
        documentId: input.documentId,
        userId: input.session.userID,
        disabledCode: 'WORKBOOK_AGENT_SHARED_THREADS_DISABLED',
        rolloutBlockedCode: 'WORKBOOK_AGENT_SHARED_THREADS_ROLLOUT_BLOCKED',
      })
    }
    if (parsed.threadId !== undefined) {
      const sharedSession = this.tryGetSessionByThreadId(parsed.threadId)
      if (sharedSession) {
        const accessibleSession = this.requireOwnedSession(sharedSession, input.documentId, input.session.userID)
        if (parsed.context) {
          updateWorkbookAgentDurableUiContextFromUser({
            sessionState: accessibleSession,
            context: parsed.context,
            userId: input.session.userID,
          })
        }
        if (parsed.executionPolicy) {
          accessibleSession.executionPolicy = normalizeExecutionPolicy({
            scope: accessibleSession.scope,
            requestedPolicy: parsed.executionPolicy,
          })
        }
        if (parsed.context || parsed.executionPolicy) {
          await this.persistSessionState(accessibleSession)
          this.sessionRegistry.emitSnapshot(accessibleSession.threadId)
        }
        this.sessionRegistry.touch(accessibleSession)
        return buildSnapshot(accessibleSession)
      }
    }

    let thread: Awaited<ReturnType<CodexAppServerTransport['threadStart']>> | null = null
    let sessionBootstrapError: unknown = null
    try {
      const codexClient = await this.codexRuntime.getClient()
      thread =
        parsed.threadId === undefined
          ? await codexClient.threadStart(createWorkbookAgentThreadStartInput())
          : await codexClient.threadResume(createWorkbookAgentThreadResumeInput(parsed.threadId))
    } catch (error) {
      if (parsed.threadId === undefined) {
        throw error
      }
      sessionBootstrapError = error
    }
    const threadId = thread?.id ?? parsed.threadId
    if (!threadId) {
      throw sessionBootstrapError instanceof Error ? sessionBootstrapError : new Error('Workbook agent thread bootstrap failed')
    }
    const durableThreadSession = await this.threadRepository.loadThreadState({
      documentId: input.documentId,
      actorUserId: input.session.userID,
      threadId,
    })
    const durableThreadState = durableThreadSession.threadState
    if (durableThreadState?.scope === 'shared') {
      assertWorkbookAgentSharedThreadAccess({
        featureFlags: this.featureFlags,
        documentId: input.documentId,
        userId: input.session.userID,
        disabledCode: 'WORKBOOK_AGENT_SHARED_THREADS_DISABLED',
        rolloutBlockedCode: 'WORKBOOK_AGENT_SHARED_THREADS_ROLLOUT_BLOCKED',
      })
    }
    if (!thread && !durableThreadState) {
      throw sessionBootstrapError instanceof Error ? sessionBootstrapError : new Error('Workbook agent thread bootstrap failed')
    }
    const sessionState = createWorkbookAgentBootstrappedSessionState({
      documentId: input.documentId,
      userId: input.session.userID,
      threadId,
      ...(parsed.scope === undefined ? {} : { requestedScope: parsed.scope }),
      ...(parsed.executionPolicy === undefined ? {} : { requestedExecutionPolicy: parsed.executionPolicy }),
      ...(parsed.context === undefined ? {} : { requestedContext: parsed.context }),
      durableThreadSession,
      liveThread: thread,
      sessionBootstrapError,
      now: this.now(),
    })
    const bootstrapRecovery = planWorkbookAgentBootstrapReviewRecovery({
      sessionState,
      rolloutAllowed: this.isRolloutAllowed(input.documentId, input.session.userID),
      autoApplyLowRiskEnabled: this.featureFlags.autoApplyLowRiskEnabled,
    })
    if (bootstrapRecovery.kind === 'autoApply') {
      const migratedReviewItem = rebaseWorkbookAgentBootstrapReviewItem({
        reviewItem: bootstrapRecovery.reviewItem,
        currentRevision: await this.zeroSyncService.getWorkbookHeadRevision(input.documentId),
        fallbackOwnerUserId: input.session.userID,
      })
      const migratedBundle = toWorkbookAgentCommandBundle(migratedReviewItem)
      replaceCurrentWorkbookAgentReviewItem(sessionState, migratedReviewItem)
      const preview = await this.buildAuthoritativePreview(input.documentId, migratedBundle)
      await this.applyCommandBundleForSessionState({
        sessionState,
        commandBundle: migratedBundle,
        actorUserId: input.session.userID,
        appliedBy: 'auto',
        preview,
      })
    } else if (bootstrapRecovery.kind === 'clearLegacy') {
      clearLegacyPrivateBootstrapReviewItem({
        sessionState,
        reviewItem: bootstrapRecovery.reviewItem,
        now: this.now(),
      })
    }
    this.sessionRegistry.storeSession(sessionState, (evictedThreadId) => {
      this.codexRuntime.releaseThread(evictedThreadId)
    })
    await this.persistSessionState(sessionState)
    return buildSnapshot(sessionState)
  }

  async updateContext(input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = updateContextBodySchema.parse(input.body)
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    updateWorkbookAgentDurableUiContextFromUser({
      sessionState,
      context: parsed.context,
      userId: input.session.userID,
    })
    this.sessionRegistry.touch(sessionState)
    await this.persistSessionState(sessionState)
    this.sessionRegistry.emitSnapshot(sessionState.threadId)
    return buildSnapshot(sessionState)
  }

  async startTurn(input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = startTurnBodySchema.parse(input.body)
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    if (sessionState.live.activeTurnId) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_TURN_ALREADY_RUNNING',
        message: 'Finish or interrupt the current assistant turn before starting another one.',
        statusCode: 409,
        retryable: false,
      })
    }
    this.assertTurnQuota(input.documentId, input.session.userID)
    if (parsed.context) {
      sessionState.durable.context = parsed.context
    }
    const turnContext = cloneUiContext(sessionState.durable.context)
    const codexClient = await this.codexRuntime.getClient()
    let turn
    try {
      turn = await codexClient.turnStart({
        threadId: sessionState.threadId,
        prompt: parsed.prompt,
      })
    } catch (error) {
      if (isCodexAppServerPoolBackpressureError(error)) {
        this.sessionRegistry.incrementCounter('turnBackpressureCount')
        throw createWorkbookAgentServiceError({
          code: 'WORKBOOK_AGENT_TURN_BACKPRESSURE',
          message: error.message,
          statusCode: 429,
          retryable: true,
        })
      }
      throw error
    }
    const optimisticEntryId = `optimistic-user:${turn.id}`
    sessionState.durable.entries = upsertEntry(sessionState.durable.entries, {
      id: optimisticEntryId,
      kind: 'user',
      turnId: turn.id,
      text: parsed.prompt,
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [],
    })
    sessionState.live.optimisticUserEntryIdByTurn.set(turn.id, optimisticEntryId)
    sessionState.live.promptByTurn.set(turn.id, parsed.prompt)
    sessionState.live.turnActorUserIdByTurn.set(turn.id, input.session.userID)
    sessionState.live.turnContextByTurn.set(turn.id, turnContext)
    sessionState.live.activeTurnId = turn.id
    sessionState.live.status = 'inProgress'
    sessionState.live.lastError = null
    this.sessionRegistry.touch(sessionState)
    await this.persistSessionState(sessionState)
    this.sessionRegistry.emitSnapshot(sessionState.threadId)
    return buildSnapshot(sessionState)
  }

  async startWorkflow(input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = startWorkflowBodySchema.parse(input.body)
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    if (!this.featureFlags.workflowRunnerEnabled) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_WORKFLOW_RUNNER_DISABLED',
        message: 'Workbook assistant workflows are currently disabled.',
        statusCode: 409,
        retryable: false,
      })
    }
    this.assertRolloutAllowed({
      documentId: input.documentId,
      userId: input.session.userID,
      code: 'WORKBOOK_AGENT_WORKFLOW_RUNNER_ROLLOUT_BLOCKED',
      message: 'Workbook assistant workflows are still limited to the rollout allowlist.',
    })
    if (parsed.context) {
      updateWorkbookAgentDurableUiContextFromUser({
        sessionState,
        context: parsed.context,
        userId: input.session.userID,
      })
    }
    const runningWorkflow = sessionState.durable.workflowRuns.find((run) => run.status === 'running')
    if (runningWorkflow) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_WORKFLOW_ALREADY_RUNNING',
        message: `Finish or cancel the running workflow before starting another one: ${runningWorkflow.title}`,
        statusCode: 409,
        retryable: false,
      })
    }
    const workflowTemplate = parsed.workflowTemplate
    this.assertWorkflowFamilyEnabled(workflowTemplate)
    if (sessionState.durable.reviewQueueItems.length > 0 && isMutatingWorkflowTemplate(workflowTemplate)) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_REVIEW_ITEM_EXISTS',
        message: 'Finish the current workbook review item before starting another mutating workflow.',
        statusCode: 409,
        retryable: false,
      })
    }
    const workflowInput = {
      ...('query' in parsed ? { query: parsed.query } : {}),
      ...('sheetName' in parsed && parsed.sheetName ? { sheetName: parsed.sheetName } : {}),
      ...('limit' in parsed && parsed.limit !== undefined ? { limit: parsed.limit } : {}),
      ...('name' in parsed ? { name: parsed.name } : {}),
    }
    const workflowTurnId = sessionState.live.activeTurnId ?? createWorkflowTurnId(crypto.randomUUID())
    await this.workflowRuntime.startWorkflow({
      sessionState,
      documentId: input.documentId,
      actorUserId: input.session.userID,
      workflowTemplate,
      workflowInput,
      workflowTurnId,
    })
    return buildSnapshot(sessionState)
  }

  async cancelWorkflow(input: {
    documentId: string
    threadId: string
    runId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    const runningWorkflow = sessionState.durable.workflowRuns.find((run) => run.runId === input.runId)
    if (!runningWorkflow) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_WORKFLOW_NOT_FOUND',
        message: 'Workbook agent workflow run not found',
        statusCode: 404,
        retryable: false,
      })
    }
    if (runningWorkflow.status !== 'running') {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_WORKFLOW_NOT_RUNNING',
        message: 'Workbook agent workflow is not currently running',
        statusCode: 409,
        retryable: false,
      })
    }
    await this.workflowRuntime.cancelRunningWorkflow({
      sessionState,
      documentId: input.documentId,
      runId: input.runId,
      runningWorkflow,
      actorUserId: input.session.userID,
    })
    return buildSnapshot(sessionState)
  }

  async interruptTurn(input: { documentId: string; threadId: string; session: SessionIdentity }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    const codexClient = await this.codexRuntime.getClient()
    await codexClient.turnInterrupt(sessionState.threadId)
    return buildSnapshot(sessionState)
  }

  async applyReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null
    preview: unknown
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    const reviewItem = requireWorkbookAgentReviewItem({
      reviewItem: getCurrentWorkbookAgentReviewItem(sessionState),
      reviewItemId: input.reviewItemId,
      notFoundMessage: 'Workbook agent change set was not found.',
    })
    const preview = decodeWorkbookAgentPreviewSummary(input.preview)
    if (!preview) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_PREVIEW_REQUIRED',
        message: 'Workbook preview details are required before applying this change set.',
        statusCode: 400,
        retryable: false,
      })
    }
    await this.applyCommandBundleForSessionState({
      sessionState,
      commandBundle: toWorkbookAgentCommandBundle(reviewItem),
      actorUserId: input.session.userID,
      appliedBy: input.appliedBy,
      commandIndexes: input.commandIndexes,
      preview,
    })
    await this.persistSessionState(sessionState)
    this.sessionRegistry.emitSnapshot(sessionState.threadId)
    return buildSnapshot(sessionState)
  }

  async reviewReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = reviewReviewItemBodySchema.parse(input.body)
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    const reviewItem = requireWorkbookAgentReviewItem({
      reviewItem: getCurrentWorkbookAgentReviewItem(sessionState),
      reviewItemId: input.reviewItemId,
      notFoundMessage: 'Workbook review item was not found.',
    })
    const now = this.now()
    const transition = transitionWorkbookAgentSharedReview({
      sessionState,
      reviewItem,
      decision: parsed.decision,
      reviewerUserId: input.session.userID,
      now,
    })
    this.sessionRegistry.incrementCounter(transition.counter as WorkbookAgentObservabilityCounterName)
    replaceCurrentWorkbookAgentReviewItem(sessionState, transition.nextReviewItem)
    sessionState.durable.entries = upsertEntry(
      sessionState.durable.entries,
      createSystemEntry(
        `system-review:${transition.reviewedBundle.id}:${now}`,
        transition.reviewedBundle.turnId,
        transition.entryText,
        createBundleRangeCitations(transition.reviewedBundle),
      ),
    )
    this.sessionRegistry.touch(sessionState)
    await this.persistSessionState(sessionState)
    this.sessionRegistry.emitSnapshot(sessionState.threadId)
    return buildSnapshot(sessionState)
  }

  async dismissReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    const reviewItem = requireWorkbookAgentReviewItem({
      reviewItem: getCurrentWorkbookAgentReviewItem(sessionState),
      reviewItemId: input.reviewItemId,
      notFoundMessage: 'Workbook review item was not found.',
    })
    replaceCurrentWorkbookAgentReviewItem(sessionState, null)
    sessionState.durable.entries = upsertEntry(
      sessionState.durable.entries,
      createWorkbookAgentDismissReviewEntry({ reviewItem, now: this.now() }),
    )
    this.sessionRegistry.touch(sessionState)
    await this.persistSessionState(sessionState)
    this.sessionRegistry.emitSnapshot(sessionState.threadId)
    return buildSnapshot(sessionState)
  }

  async replayExecutionRecord(input: {
    documentId: string
    threadId: string
    recordId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    const record = sessionState.durable.executionRecords.find((entry) => entry.id === input.recordId)
    if (!record) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_RUN_NOT_FOUND',
        message: 'Workbook agent execution record not found',
        statusCode: 404,
        retryable: false,
      })
    }
    const baseRevision = await this.zeroSyncService.getWorkbookHeadRevision(input.documentId)
    const replayedBundle = createWorkbookAgentCommandBundle({
      documentId: input.documentId,
      threadId: sessionState.threadId,
      turnId: `replay:${record.id}:${String(this.now())}`,
      goalText: record.goalText,
      baseRevision,
      context: toContextRef(sessionState.durable.context) ?? record.context,
      commands: record.commands,
      now: this.now(),
    })
    if (this.shouldApplyToolBundleImmediately(sessionState, replayedBundle)) {
      const preview = await this.buildAuthoritativePreview(input.documentId, replayedBundle)
      await this.applyCommandBundleForSessionState({
        sessionState,
        commandBundle: replayedBundle,
        actorUserId: input.session.userID,
        appliedBy: 'auto',
        preview,
      })
      await this.persistSessionState(sessionState)
      this.sessionRegistry.emitSnapshot(sessionState.threadId)
      return buildSnapshot(sessionState)
    }
    if (sessionState.scope === 'private') {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_PRIVATE_EXECUTION_BLOCKED',
        message:
          'Private workbook threads execute replayed changes directly and do not queue review items under the current execution policy.',
        statusCode: 409,
        retryable: false,
      })
    }
    replaceCurrentWorkbookAgentReviewItem(sessionState, createWorkbookAgentReviewQueueItem({ sessionState, bundle: replayedBundle }))
    sessionState.durable.entries = upsertEntry(
      sessionState.durable.entries,
      createSystemEntry(
        `system-replay:${record.id}:${String(this.now())}`,
        replayedBundle.turnId,
        `Prepared workbook review item from a prior execution: ${replayedBundle.summary}`,
        createBundleRangeCitations(replayedBundle),
      ),
    )
    this.sessionRegistry.touch(sessionState)
    await this.persistSessionState(sessionState)
    this.sessionRegistry.emitSnapshot(sessionState.threadId)
    return buildSnapshot(sessionState)
  }

  async listThreads(input: { documentId: string; session: SessionIdentity }): Promise<WorkbookAgentThreadSummary[]> {
    const summaries = await this.zeroSyncService.listWorkbookAgentThreadSummaries(input.documentId, input.session.userID)
    return filterWorkbookAgentThreadSummariesByAccessPolicy({
      featureFlags: this.featureFlags,
      documentId: input.documentId,
      summaries,
      userId: input.session.userID,
    })
  }

  getSnapshot(input: { documentId: string; threadId: string; session: SessionIdentity }): WorkbookAgentThreadSnapshot {
    const sessionState = this.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    this.sessionRegistry.touch(sessionState)
    return buildSnapshot(sessionState)
  }

  subscribe(threadId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void {
    return this.sessionRegistry.subscribe(threadId, listener)
  }

  async close(): Promise<void> {
    this.workflowRuntime.close()
    await this.codexRuntime.close()
    this.sessionRegistry.clear()
  }

  private resolveTurnActorUserId(sessionState: WorkbookAgentThreadState, turnId: string): string {
    return sessionState.live.turnActorUserIdByTurn.get(turnId) ?? sessionState.userId
  }

  private resolveTurnContext(sessionState: WorkbookAgentThreadState, turnId: string) {
    return cloneUiContext(sessionState.live.turnContextByTurn.get(turnId) ?? sessionState.durable.context)
  }

  private getOwnedSession(documentId: string, threadId: string, userId: string): WorkbookAgentThreadState {
    const sessionState = this.sessionRegistry.tryGetSession(threadId)
    if (!sessionState) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_THREAD_NOT_FOUND',
        message: 'Workbook agent thread not found',
        statusCode: 404,
        retryable: true,
      })
    }
    return this.requireOwnedSession(sessionState, documentId, userId)
  }

  private requireOwnedSession(sessionState: WorkbookAgentThreadState, documentId: string, userId: string): WorkbookAgentThreadState {
    if (sessionState.documentId !== documentId) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_THREAD_NOT_FOUND',
        message: 'Workbook agent thread not found',
        statusCode: 404,
        retryable: false,
      })
    }
    if (sessionState.scope !== 'shared' && sessionState.userId !== userId) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_THREAD_NOT_FOUND',
        message: 'Workbook agent thread not found',
        statusCode: 404,
        retryable: false,
      })
    }
    assertWorkbookAgentSessionAccessPolicy({
      featureFlags: this.featureFlags,
      sessionState,
      documentId,
      userId,
    })
    return sessionState
  }

  private getSessionByThreadId(threadId: string): WorkbookAgentThreadState {
    const sessionState = this.tryGetSessionByThreadId(threadId)
    if (!sessionState) {
      throw new Error(`Workbook agent thread not found for thread ${threadId}`)
    }
    return sessionState
  }

  private tryGetSessionByThreadId(threadId: string): WorkbookAgentThreadState | null {
    return this.sessionRegistry.tryGetSession(threadId)
  }

  private assertTurnQuota(documentId: string, actorUserId: string): void {
    assertWorkbookAgentTurnQuota({
      sessions: this.sessionRegistry.listSessions(),
      documentId,
      actorUserId,
      maxActiveTurnsPerUser: this.maxActiveTurnsPerUser,
      maxActiveTurnsPerDocument: this.maxActiveTurnsPerDocument,
    })
  }

  private assertWorkflowFamilyEnabled(workflowTemplate: WorkbookAgentWorkflowRun['workflowTemplate']): void {
    if (isWorkbookAgentWorkflowFamilyEnabled(this.featureFlags, workflowTemplate)) {
      return
    }
    const workflowFamily = getWorkbookAgentWorkflowFamily(workflowTemplate)
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_WORKFLOW_FAMILY_DISABLED',
      message: `Workbook assistant ${workflowFamily} workflows are currently disabled.`,
      statusCode: 409,
      retryable: false,
    })
  }
}

export function createWorkbookAgentService(
  zeroSyncService: ZeroSyncService,
  options: Omit<EnabledWorkbookAgentServiceOptions, 'zeroSyncService'> = {},
): WorkbookAgentService {
  if (!zeroSyncService.enabled) {
    return new DisabledWorkbookAgentService()
  }
  return new EnabledWorkbookAgentService({
    zeroSyncService,
    ...options,
  })
}
