import type {
  WorkbookAgentThreadSnapshot,
  WorkbookAgentStreamEvent,
  WorkbookAgentThreadSummary,
  WorkbookAgentWorkflowRun,
} from '@bilig/contracts'
import type { WorkbookAgentAppliedBy, WorkbookAgentCommandBundle, WorkbookAgentExecutionRecord } from '@bilig/agent-api'
import {
  canCancelWorkbookAgentWorkflowRun,
  canInterruptWorkbookAgentTurn,
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
  findRecoveredStaleBootstrapWorkflowRuns,
  planWorkbookAgentBootstrapReviewRecovery,
  rebaseWorkbookAgentBootstrapReviewItem,
} from './workbook-agent-service-bootstrap.js'
import {
  WorkbookAgentCodexRuntime,
  createWorkbookAgentThreadResumeInput,
  createWorkbookAgentThreadStartInput,
} from './workbook-agent-codex-runtime.js'
import {
  assertNoWorkbookAgentReviewItem,
  assertWorkbookAgentReviewDismissAllowed,
  createWorkbookAgentDismissReviewEntry,
  getCurrentWorkbookAgentReviewItem,
  replaceCurrentWorkbookAgentReviewItem,
  requireWorkbookAgentReviewItem,
  stageWorkbookAgentReviewBundle,
  transitionWorkbookAgentSharedReview,
} from './workbook-agent-review-transitions.js'
import { WorkbookAgentThreadRepository } from './workbook-agent-thread-repository.js'
import { WorkbookAgentSessionAuthority } from './workbook-agent-session-authority.js'
import {
  createWorkbookAgentBundleApplicationContext,
  createWorkbookAgentReviewActionContext,
} from './workbook-agent-service-action-contexts.js'
import {
  buildSnapshot,
  cloneUiContext,
  isMutatingWorkflowTemplate,
  normalizeExecutionPolicy,
  type WorkbookAgentThreadState,
  upsertEntry,
} from './workbook-agent-service-shared.js'
import { updateWorkbookAgentDurableUiContextFromUser } from './workbook-agent-durable-context-sync.js'
import {
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
  finalizeWorkbookAgentPrivateTurnBundle,
} from './workbook-agent-service-application.js'
import { applyWorkbookAgentReviewItem, replayWorkbookAgentExecutionRecord } from './workbook-agent-service-review-actions.js'
import { startWorkbookAgentTurn } from './workbook-agent-turn-lifecycle.js'

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
  private readonly sessionAuthority: WorkbookAgentSessionAuthority
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
    this.threadRepository = new WorkbookAgentThreadRepository(this.zeroSyncService)
    this.sessionAuthority = new WorkbookAgentSessionAuthority({
      featureFlags: () => this.featureFlags,
      sessionRegistry: this.sessionRegistry,
      threadRepository: this.threadRepository,
    })
    this.codexRuntime = new WorkbookAgentCodexRuntime({
      zeroSyncService: this.zeroSyncService,
      now: this.now,
      maxCodexClients: this.maxCodexClients,
      maxConcurrentTurnsPerCodexClient: this.maxConcurrentTurnsPerCodexClient,
      maxQueuedTurnsPerCodexClient: this.maxQueuedTurnsPerCodexClient,
      getSessionByThreadId: (threadId) => this.sessionAuthority.getSessionByThreadId(threadId),
      tryGetSessionByThreadId: (threadId) => this.sessionAuthority.tryGetSessionByThreadId(threadId),
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
    this.workflowRuntime = new WorkbookAgentWorkflowRuntime({
      zeroSyncService: this.zeroSyncService,
      now: this.now,
      ...(options.workflowShutdownDrainTimeoutMs === undefined ? {} : { shutdownDrainTimeoutMs: options.workflowShutdownDrainTimeoutMs }),
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
    return createWorkbookAgentBundleApplicationContext({
      zeroSyncService: this.zeroSyncService,
      now: this.now,
      featureFlags: this.featureFlags,
      sessionRegistry: this.sessionRegistry,
      isRolloutAllowed: (documentId: string, userId: string) => this.isRolloutAllowed(documentId, userId),
    })
  }

  private async applyCommandBundleForSessionState(input: {
    sessionState: WorkbookAgentThreadState
    commandBundle: WorkbookAgentCommandBundle
    actorUserId: string
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null | undefined
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

  private createReviewActionContext() {
    return createWorkbookAgentReviewActionContext({
      zeroSyncService: this.zeroSyncService,
      now: this.now,
      sessionRegistry: this.sessionRegistry,
      applyCommandBundleForSessionState: async (input: {
        sessionState: WorkbookAgentThreadState
        commandBundle: WorkbookAgentCommandBundle
        actorUserId: string
        appliedBy: WorkbookAgentAppliedBy
        commandIndexes?: readonly number[] | null | undefined
      }) => await this.applyCommandBundleForSessionState(input),
      shouldApplyToolBundleImmediately: (sessionState: WorkbookAgentThreadState, bundle: WorkbookAgentCommandBundle) =>
        this.shouldApplyToolBundleImmediately(sessionState, bundle),
      persistSessionState: async (sessionState: WorkbookAgentThreadState) => await this.persistSessionState(sessionState),
    })
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
        const accessibleSession = this.sessionAuthority.requireOwnedSession(sharedSession, input.documentId, input.session.userID)
        await this.sessionAuthority.authorizeSharedSessionForUser(accessibleSession, input.documentId, input.session.userID)
        let contextChanged = false
        if (parsed.context) {
          contextChanged = updateWorkbookAgentDurableUiContextFromUser({
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
        if (contextChanged || parsed.executionPolicy) {
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
    const recoveredBootstrapWorkflowRuns = findRecoveredStaleBootstrapWorkflowRuns({
      previousWorkflowRuns: durableThreadSession.workflowRuns,
      nextWorkflowRuns: sessionState.durable.workflowRuns,
    })
    sessionState.live.authorizedUserIds.add(input.session.userID)
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
      await this.applyCommandBundleForSessionState({
        sessionState,
        commandBundle: migratedBundle,
        actorUserId: input.session.userID,
        appliedBy: 'auto',
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
    await Promise.all(
      recoveredBootstrapWorkflowRuns.map(async (run) => await this.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, run)),
    )
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
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
    const contextChanged = updateWorkbookAgentDurableUiContextFromUser({
      sessionState,
      context: parsed.context,
      userId: input.session.userID,
    })
    this.sessionRegistry.touch(sessionState)
    if (contextChanged) {
      await this.persistSessionState(sessionState)
      this.sessionRegistry.emitSnapshot(sessionState.threadId)
    }
    return buildSnapshot(sessionState)
  }

  async startTurn(input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot> {
    const parsed = startTurnBodySchema.parse(input.body)
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
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
    startWorkbookAgentTurn(sessionState, {
      turnId: turn.id,
      prompt: parsed.prompt,
      actorUserId: input.session.userID,
      context: turnContext,
      optimisticEntryId,
    })
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
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
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
    if (isMutatingWorkflowTemplate(workflowTemplate)) {
      assertNoWorkbookAgentReviewItem({
        sessionState,
        message: 'Finish the current workbook review item before starting another mutating workflow.',
      })
    }
    if (parsed.context) {
      updateWorkbookAgentDurableUiContextFromUser({
        sessionState,
        context: parsed.context,
        userId: input.session.userID,
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
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
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
    if (
      !canCancelWorkbookAgentWorkflowRun({
        scope: sessionState.scope,
        ownerUserId: sessionState.storageActorUserId,
        actorUserId: input.session.userID,
        startedByUserId: runningWorkflow.startedByUserId,
      })
    ) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_WORKFLOW_CANCEL_FORBIDDEN',
        message: 'Only the workflow starter or shared thread owner can cancel this workflow.',
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
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
    const activeTurnId = sessionState.live.activeTurnId
    if (!activeTurnId || sessionState.live.status !== 'inProgress') {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_TURN_NOT_RUNNING',
        message: 'Workbook agent turn is not currently running',
        statusCode: 409,
        retryable: false,
      })
    }
    const turnActorUserId = this.resolveTurnActorUserId(sessionState, activeTurnId)
    if (
      !canInterruptWorkbookAgentTurn({
        scope: sessionState.scope,
        ownerUserId: sessionState.storageActorUserId,
        actorUserId: input.session.userID,
        turnActorUserId,
      })
    ) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_TURN_INTERRUPT_FORBIDDEN',
        message: 'Only the active turn author or shared thread owner can stop this turn.',
        statusCode: 409,
        retryable: false,
      })
    }
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
  }): Promise<WorkbookAgentThreadSnapshot> {
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
    await applyWorkbookAgentReviewItem({
      context: this.createReviewActionContext(),
      sessionState,
      reviewItemId: input.reviewItemId,
      actorUserId: input.session.userID,
      appliedBy: input.appliedBy,
      commandIndexes: input.commandIndexes,
    })
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
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
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
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
    const reviewItem = requireWorkbookAgentReviewItem({
      reviewItem: getCurrentWorkbookAgentReviewItem(sessionState),
      reviewItemId: input.reviewItemId,
      notFoundMessage: 'Workbook review item was not found.',
    })
    assertWorkbookAgentReviewDismissAllowed({
      sessionState,
      reviewItem,
      actorUserId: input.session.userID,
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
    const sessionState = await this.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
    await replayWorkbookAgentExecutionRecord({
      context: this.createReviewActionContext(),
      sessionState,
      documentId: input.documentId,
      recordId: input.recordId,
      actorUserId: input.session.userID,
    })
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
    const sessionState = this.sessionAuthority.getOwnedSession(input.documentId, input.threadId, input.session.userID)
    this.sessionAuthority.assertSharedSessionAlreadyAuthorized(sessionState, input.session.userID)
    this.sessionRegistry.touch(sessionState)
    return buildSnapshot(sessionState)
  }

  subscribe(threadId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void {
    return this.sessionRegistry.subscribe(threadId, listener)
  }

  async close(): Promise<void> {
    await this.workflowRuntime.close()
    await this.codexRuntime.close()
    this.sessionRegistry.clear()
  }

  private resolveTurnActorUserId(sessionState: WorkbookAgentThreadState, turnId: string): string {
    return sessionState.live.turnActorUserIdByTurn.get(turnId) ?? sessionState.userId
  }

  private resolveTurnContext(sessionState: WorkbookAgentThreadState, turnId: string) {
    return cloneUiContext(sessionState.live.turnContextByTurn.get(turnId) ?? sessionState.durable.context)
  }

  private async getAuthorizedSession(documentId: string, threadId: string, userId: string): Promise<WorkbookAgentThreadState> {
    return await this.sessionAuthority.getAuthorizedSession(documentId, threadId, userId)
  }

  private tryGetSessionByThreadId(threadId: string): WorkbookAgentThreadState | null {
    return this.sessionAuthority.tryGetSessionByThreadId(threadId)
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
