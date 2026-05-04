import type {
  WorkbookAgentThreadSnapshot,
  WorkbookAgentStreamEvent,
  WorkbookAgentThreadSummary,
  WorkbookAgentWorkflowRun,
} from '@bilig/contracts'
import type {
  WorkbookAgentAppliedBy,
  WorkbookAgentCommandBundle,
  WorkbookAgentCommand,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewSummary,
} from '@bilig/agent-api'
import {
  buildWorkbookAgentPreview,
  buildWorkbookAgentExecutionRecord,
  createWorkbookAgentCommandBundle,
  decodeWorkbookAgentPreviewSummary,
  isWorkbookAgentBundleAutoApplyEligible,
  requiresWorkbookAgentOwnerReview,
  resolveWorkbookAgentBundleExecutionPolicyInput,
  splitWorkbookAgentCommandBundle,
  toWorkbookAgentCommandBundle,
} from '@bilig/agent-api'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { isCodexAppServerPoolBackpressureError } from './codex-app-server-pool.js'
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
import {
  appendRevisionCitation,
  attachSharedReviewState,
  createBundleRangeCitations,
  createWorkflowTurnId,
  normalizeSharedReviewState,
} from './workbook-agent-bundle-state.js'
import {
  clearLegacyPrivateBootstrapReviewItem,
  createWorkbookAgentBootstrappedSessionState,
  planWorkbookAgentBootstrapReviewRecovery,
  rebaseWorkbookAgentBootstrapReviewItem,
} from './workbook-agent-service-bootstrap.js'
import {
  WorkbookAgentCodexRuntime,
  DEFAULT_MAX_CODEX_CLIENTS,
  DEFAULT_MAX_CODEX_CONCURRENT_TURNS_PER_CLIENT,
  DEFAULT_MAX_CODEX_QUEUED_TURNS_PER_CLIENT,
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
import { applyWorkbookAgentStructuralContextHints, updateWorkbookAgentDurableUiContextFromUser } from './workbook-agent-service-context.js'
import { assertWorkbookAgentTurnQuota } from './workbook-agent-service-session-policy.js'
import { WorkbookAgentWorkflowRuntime } from './workbook-agent-workflow-runtime.js'
import { normalizeWorkbookAgentUiContext } from './workbook-agent-inspection.js'
import {
  WorkbookAgentSessionRegistry,
  createDisabledWorkbookAgentObservabilitySnapshot,
  type WorkbookAgentObservabilityCounterName,
  type WorkbookAgentObservabilitySnapshot,
} from './workbook-agent-session-registry.js'

function parsePositiveIntegerEnv(value: string | undefined, fallback: number): number {
  if (!value) {
    return fallback
  }
  const parsed = Number.parseInt(value, 10)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback
}

const DEFAULT_MAX_ACTIVE_TURNS_PER_USER = parsePositiveIntegerEnv(process.env['BILIG_CODEX_MAX_ACTIVE_TURNS_PER_USER'], 4)
const DEFAULT_MAX_ACTIVE_TURNS_PER_DOCUMENT = parsePositiveIntegerEnv(process.env['BILIG_CODEX_MAX_ACTIVE_TURNS_PER_DOCUMENT'], 16)

export interface WorkbookAgentService {
  readonly enabled: boolean
  createSession(input: { documentId: string; session: SessionIdentity; body: unknown }): Promise<WorkbookAgentThreadSnapshot>
  updateContext(input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot>
  startTurn(input: { documentId: string; threadId: string; session: SessionIdentity; body: unknown }): Promise<WorkbookAgentThreadSnapshot>
  startWorkflow(input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot>
  cancelWorkflow(input: {
    documentId: string
    threadId: string
    runId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot>
  interruptTurn(input: { documentId: string; threadId: string; session: SessionIdentity }): Promise<WorkbookAgentThreadSnapshot>
  applyReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null
    preview: unknown
  }): Promise<WorkbookAgentThreadSnapshot>
  reviewReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot>
  dismissReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot>
  replayExecutionRecord(input: {
    documentId: string
    threadId: string
    recordId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot>
  listThreads(input: { documentId: string; session: SessionIdentity }): Promise<WorkbookAgentThreadSummary[]>
  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot
  getSnapshot(input: { documentId: string; threadId: string; session: SessionIdentity }): WorkbookAgentThreadSnapshot
  subscribe(threadId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void
  close(): Promise<void>
}

class DisabledWorkbookAgentService implements WorkbookAgentService {
  readonly enabled = false

  async createSession(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async updateContext(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async startTurn(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async startWorkflow(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async cancelWorkflow(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async interruptTurn(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async applyReviewItem(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async reviewReviewItem(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async dismissReviewItem(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async replayExecutionRecord(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async listThreads(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot {
    return createDisabledWorkbookAgentObservabilitySnapshot(Date.now())
  }

  getSnapshot(): never {
    throw new Error('Workbook agent service is not configured')
  }

  subscribe(): () => void {
    return () => {}
  }

  async close(): Promise<void> {}
}

export interface EnabledWorkbookAgentServiceOptions {
  zeroSyncService: ZeroSyncService
  codexClientFactory?: (options: CodexAppServerClientOptions) => CodexAppServerTransport
  now?: () => number
  maxSessions?: number
  maxCodexClients?: number
  maxConcurrentTurnsPerCodexClient?: number
  maxQueuedTurnsPerCodexClient?: number
  maxActiveTurnsPerUser?: number
  maxActiveTurnsPerDocument?: number
  featureFlags?: Partial<WorkbookAgentFeatureFlags>
}

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
    this.zeroSyncService = options.zeroSyncService
    this.now = options.now ?? (() => Date.now())
    this.maxSessions = options.maxSessions ?? 64
    this.maxCodexClients = options.maxCodexClients ?? DEFAULT_MAX_CODEX_CLIENTS
    this.maxConcurrentTurnsPerCodexClient = options.maxConcurrentTurnsPerCodexClient ?? DEFAULT_MAX_CODEX_CONCURRENT_TURNS_PER_CLIENT
    this.maxQueuedTurnsPerCodexClient = options.maxQueuedTurnsPerCodexClient ?? DEFAULT_MAX_CODEX_QUEUED_TURNS_PER_CLIENT
    this.maxActiveTurnsPerUser = options.maxActiveTurnsPerUser ?? DEFAULT_MAX_ACTIVE_TURNS_PER_USER
    this.maxActiveTurnsPerDocument = options.maxActiveTurnsPerDocument ?? DEFAULT_MAX_ACTIVE_TURNS_PER_DOCUMENT
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

  private async buildAuthoritativePreview(documentId: string, bundle: WorkbookAgentCommandBundle) {
    return this.zeroSyncService.inspectWorkbook(documentId, async (runtime) =>
      buildWorkbookAgentPreview({
        snapshot: runtime.engine.exportSnapshot(),
        replicaId: `server:${runtime.documentId}:agent-preview`,
        bundle,
      }),
    )
  }

  private async refreshAppliedWorkbookContext(input: {
    sessionState: WorkbookAgentThreadState
    turnId: string
    commands: readonly WorkbookAgentCommand[]
  }): Promise<void> {
    const hintedContext = applyWorkbookAgentStructuralContextHints(
      this.resolveTurnContext(input.sessionState, input.turnId),
      input.commands,
    )
    const normalizedContext = await this.zeroSyncService.inspectWorkbook(input.sessionState.documentId, (runtime) =>
      normalizeWorkbookAgentUiContext(runtime, hintedContext),
    )
    input.sessionState.durable.context = cloneUiContext(normalizedContext)
    input.sessionState.live.turnContextByTurn.set(input.turnId, cloneUiContext(normalizedContext))
  }

  private async applyCommandBundleForSessionState(input: {
    sessionState: WorkbookAgentThreadState
    commandBundle: WorkbookAgentCommandBundle
    actorUserId: string
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null | undefined
    preview: WorkbookAgentPreviewSummary
  }): Promise<WorkbookAgentExecutionRecord> {
    const selection = splitWorkbookAgentCommandBundle({
      bundle: input.commandBundle,
      acceptedCommandIndexes: input.commandIndexes,
    })
    if (!selection.acceptedBundle || !selection.acceptedScope) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_COMMAND_SELECTION_REQUIRED',
        message: 'Select at least one staged workbook change before apply',
        statusCode: 400,
        retryable: false,
      })
    }
    if (input.appliedBy === 'auto' && selection.acceptedScope !== 'full') {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_MANUAL_APPROVAL_REQUIRED',
        message: 'Automatic apply runs one complete change set per turn.',
        statusCode: 409,
        retryable: false,
      })
    }
    if (input.appliedBy === 'auto') {
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
          code: 'WORKBOOK_AGENT_MANUAL_APPROVAL_REQUIRED',
          message: 'This session routes workbook edits through the review queue.',
          statusCode: 409,
          retryable: false,
        })
      }
      if (input.sessionState.executionPolicy === 'autoApplySafe' && !this.featureFlags.autoApplyLowRiskEnabled) {
        throw createWorkbookAgentServiceError({
          code: 'WORKBOOK_AGENT_AUTO_APPLY_DISABLED',
          message: 'Automatic safe-apply is paused for this environment.',
          statusCode: 409,
          retryable: false,
        })
      }
      this.assertRolloutAllowed({
        documentId: input.sessionState.documentId,
        userId: input.actorUserId,
        code: 'WORKBOOK_AGENT_AUTO_APPLY_ROLLOUT_BLOCKED',
        message: 'Automatic apply is limited to the rollout allowlist for this environment.',
      })
    }
    if (
      requiresWorkbookAgentOwnerReview({
        scope: input.sessionState.scope,
        riskClass: input.commandBundle.riskClass,
      }) &&
      input.sessionState.storageActorUserId !== input.actorUserId
    ) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_SHARED_APPROVAL_REQUIRED',
        message: 'Shared medium/high-risk workbook bundles must be applied by the thread owner.',
        statusCode: 409,
        retryable: false,
      })
    }
    const sharedReview = normalizeSharedReviewState(input.commandBundle, input.sessionState)
    if (
      requiresWorkbookAgentOwnerReview({
        scope: input.sessionState.scope,
        riskClass: input.commandBundle.riskClass,
      }) &&
      sharedReview?.status !== 'approved'
    ) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_SHARED_REVIEW_REQUIRED',
        message: 'Shared medium/high-risk workbook bundles must be approved by the thread owner before apply.',
        statusCode: 409,
        retryable: false,
      })
    }
    const result = await this.zeroSyncService.applyAgentCommandBundle(
      input.sessionState.documentId,
      selection.acceptedBundle,
      input.preview,
      {
        userID: input.actorUserId,
        roles: ['editor'],
      },
    )
    const executionRecord = buildWorkbookAgentExecutionRecord({
      bundle: selection.acceptedBundle,
      actorUserId: input.actorUserId,
      planText: this.collectPlanTextForTurn(input.sessionState, input.commandBundle.turnId),
      preview: result.preview,
      appliedRevision: result.revision,
      appliedBy: input.appliedBy,
      acceptedScope: selection.acceptedScope,
      now: this.now(),
    })
    await this.zeroSyncService.appendWorkbookAgentRun(executionRecord)
    input.sessionState.durable.executionRecords = [
      executionRecord,
      ...input.sessionState.durable.executionRecords.filter((record) => record.id !== executionRecord.id),
    ]
    await this.refreshAppliedWorkbookContext({
      sessionState: input.sessionState,
      turnId: input.commandBundle.turnId,
      commands: selection.acceptedBundle.commands,
    })
    replaceCurrentWorkbookAgentReviewItem(
      input.sessionState,
      selection.remainingBundle === null
        ? null
        : createWorkbookAgentReviewQueueItem({
            sessionState: input.sessionState,
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
          }),
    )
    input.sessionState.durable.entries = upsertEntry(
      input.sessionState.durable.entries,
      createSystemEntry(
        `system-apply:${executionRecord.id}`,
        input.commandBundle.turnId,
        `${input.appliedBy === 'auto' ? 'Applied automatically' : 'Applied'} ${
          selection.acceptedScope === 'partial' ? 'selected ' : ''
        }workbook change set at revision r${String(result.revision)}: ${selection.acceptedBundle.summary}`,
        appendRevisionCitation(createBundleRangeCitations(selection.acceptedBundle), result.revision),
      ),
    )
    this.sessionRegistry.touch(input.sessionState)
    return executionRecord
  }

  private async applyToolBundleAutomatically(input: {
    sessionState: WorkbookAgentThreadState
    actorUserId: string
    bundle: WorkbookAgentCommandBundle
  }): Promise<WorkbookAgentExecutionRecord | null> {
    const preview = await this.buildAuthoritativePreview(input.sessionState.documentId, input.bundle)
    const executionRecord = await this.applyCommandBundleForSessionState({
      sessionState: input.sessionState,
      commandBundle: input.bundle,
      actorUserId: input.actorUserId,
      appliedBy: 'auto',
      preview,
    })
    await this.persistSessionState(input.sessionState)
    this.sessionRegistry.emitSnapshot(input.sessionState.threadId)
    return executionRecord.bundleId === input.bundle.id ? executionRecord : null
  }

  private async finalizePrivateTurnBundle(input: {
    sessionState: WorkbookAgentThreadState
    turnId: string
    turnStatus: 'completed' | 'failed'
  }): Promise<void> {
    const queuedBundle = input.sessionState.live.stagedPrivateBundleByTurn.get(input.turnId)
    if (!queuedBundle) {
      return
    }
    input.sessionState.live.stagedPrivateBundleByTurn.delete(input.turnId)
    if (input.turnStatus !== 'completed') {
      return
    }
    const actorUserId = this.resolveTurnActorUserId(input.sessionState, input.turnId)
    try {
      const preview = await this.buildAuthoritativePreview(input.sessionState.documentId, queuedBundle)
      await this.applyCommandBundleForSessionState({
        sessionState: input.sessionState,
        commandBundle: queuedBundle,
        actorUserId,
        appliedBy: 'auto',
        preview,
      })
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      input.sessionState.live.lastError = message
      input.sessionState.durable.entries = upsertEntry(
        input.sessionState.durable.entries,
        createSystemEntry(
          `system-auto-apply-failed:${queuedBundle.id}:${this.now()}`,
          input.turnId,
          `Automatic workbook apply failed: ${queuedBundle.summary}. ${message}`,
          createBundleRangeCitations(queuedBundle),
        ),
      )
      input.sessionState.live.status = 'failed'
    }
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
    if (parsed.scope === 'shared' && !this.featureFlags.sharedThreadsEnabled) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_SHARED_THREADS_DISABLED',
        message: 'Shared workbook assistant threads are currently disabled.',
        statusCode: 409,
        retryable: false,
      })
    }
    if (parsed.scope === 'shared') {
      this.assertRolloutAllowed({
        documentId: input.documentId,
        userId: input.session.userID,
        code: 'WORKBOOK_AGENT_SHARED_THREADS_ROLLOUT_BLOCKED',
        message: 'Shared workbook assistant threads are still limited to the rollout allowlist.',
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
    if (durableThreadState?.scope === 'shared' && !this.featureFlags.sharedThreadsEnabled) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_SHARED_THREADS_DISABLED',
        message: 'Shared workbook assistant threads are currently disabled.',
        statusCode: 409,
        retryable: false,
      })
    }
    if (durableThreadState?.scope === 'shared') {
      this.assertRolloutAllowed({
        documentId: input.documentId,
        userId: input.session.userID,
        code: 'WORKBOOK_AGENT_SHARED_THREADS_ROLLOUT_BLOCKED',
        message: 'Shared workbook assistant threads are still limited to the rollout allowlist.',
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
    return await this.zeroSyncService.listWorkbookAgentThreadSummaries(input.documentId, input.session.userID)
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

  private collectPlanTextForTurn(sessionState: WorkbookAgentThreadState, turnId: string): string | null {
    const planText = sessionState.durable.entries
      .filter((entry) => entry.turnId === turnId && entry.kind === 'plan' && entry.text)
      .map((entry) => entry.text?.trim() ?? '')
      .filter((text) => text.length > 0)
      .join('\n\n')
    return planText.length > 0 ? planText : null
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
