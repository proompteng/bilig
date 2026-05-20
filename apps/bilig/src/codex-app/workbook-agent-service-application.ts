import type {
  WorkbookAgentAppliedBy,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewSummary,
} from '@bilig/agent-api'
import {
  buildWorkbookAgentExecutionRecord,
  buildWorkbookAgentPreview,
  createWorkbookAgentCommandBundle,
  isWorkbookAgentBundleAutoApplyEligible,
  requiresWorkbookAgentOwnerReview,
  resolveWorkbookAgentBundleExecutionPolicyInput,
  splitWorkbookAgentCommandBundle,
} from '@bilig/agent-api'
import type { ZeroSyncService } from '../zero/service.js'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import {
  appendRevisionCitation,
  attachSharedReviewState,
  createBundleRangeCitations,
  normalizeSharedReviewState,
} from './workbook-agent-bundle-state.js'
import { normalizeWorkbookAgentUiContext } from './workbook-agent-inspection.js'
import {
  createWorkbookAgentReviewQueueItem,
  getCurrentWorkbookAgentReviewItem,
  replaceCurrentWorkbookAgentReviewItem,
} from './workbook-agent-review-transitions.js'
import { createSystemEntry } from './workbook-agent-session-model.js'
import { applyWorkbookAgentStructuralContextHints } from './workbook-agent-structural-context-hints.js'
import { cloneUiContext, type WorkbookAgentThreadState, upsertEntry } from './workbook-agent-service-shared.js'

export interface WorkbookAgentBundleApplicationContext {
  zeroSyncService: ZeroSyncService
  now: () => number
  autoApplyLowRiskEnabled: boolean
  isRolloutAllowed: (documentId: string, userId: string) => boolean
  touchSession: (sessionState: WorkbookAgentThreadState) => void
}

export type WorkbookAgentApplyAuthorityGuard = () => void

async function buildWorkbookAgentAuthoritativePreview(input: {
  zeroSyncService: ZeroSyncService
  documentId: string
  bundle: WorkbookAgentCommandBundle
}): Promise<WorkbookAgentPreviewSummary> {
  return input.zeroSyncService.inspectWorkbook(input.documentId, async (runtime) =>
    buildWorkbookAgentPreview({
      snapshot: runtime.engine.exportSnapshot(),
      replicaId: `server:${runtime.documentId}:agent-preview`,
      bundle: input.bundle,
    }),
  )
}

async function refreshAppliedWorkbookContext(input: {
  zeroSyncService: ZeroSyncService
  sessionState: WorkbookAgentThreadState
  turnId: string
  commands: readonly WorkbookAgentCommand[]
}): Promise<void> {
  const hintedContext = applyWorkbookAgentStructuralContextHints(resolveTurnContext(input.sessionState, input.turnId), input.commands)
  const normalizedContext = await input.zeroSyncService.inspectWorkbook(input.sessionState.documentId, (runtime) =>
    normalizeWorkbookAgentUiContext(runtime, hintedContext),
  )
  input.sessionState.durable.context = cloneUiContext(normalizedContext)
  input.sessionState.live.turnContextByTurn.set(input.turnId, cloneUiContext(normalizedContext))
}

function resolveTurnContext(sessionState: WorkbookAgentThreadState, turnId: string) {
  return cloneUiContext(sessionState.live.turnContextByTurn.get(turnId) ?? sessionState.durable.context)
}

function collectPlanTextForTurn(sessionState: WorkbookAgentThreadState, turnId: string): string | null {
  const planText = sessionState.durable.entries
    .filter((entry) => entry.turnId === turnId && entry.kind === 'plan' && entry.text)
    .map((entry) => entry.text?.trim() ?? '')
    .filter((text) => text.length > 0)
    .join('\n\n')
  return planText.length > 0 ? planText : null
}

function assertAutoApplyRolloutAllowed(input: {
  context: WorkbookAgentBundleApplicationContext
  documentId: string
  userId: string
}): void {
  if (input.context.isRolloutAllowed(input.documentId, input.userId)) {
    return
  }
  throw createWorkbookAgentServiceError({
    code: 'WORKBOOK_AGENT_AUTO_APPLY_ROLLOUT_BLOCKED',
    message: 'Automatic apply is limited to the rollout allowlist for this environment.',
    statusCode: 409,
    retryable: false,
  })
}

function replaceAppliedReviewItemIfCurrent(input: {
  sessionState: WorkbookAgentThreadState
  appliedBundleId: string
  nextReviewItem: ReturnType<typeof createWorkbookAgentReviewQueueItem> | null
}): void {
  const currentReviewItem = getCurrentWorkbookAgentReviewItem(input.sessionState)
  if (!currentReviewItem || currentReviewItem.id !== input.appliedBundleId) {
    return
  }
  replaceCurrentWorkbookAgentReviewItem(input.sessionState, input.nextReviewItem)
}

function assertManualReviewItemStillCurrent(input: {
  sessionState: WorkbookAgentThreadState
  commandBundle: WorkbookAgentCommandBundle
}): void {
  const currentReviewItem = getCurrentWorkbookAgentReviewItem(input.sessionState)
  if (currentReviewItem?.id === input.commandBundle.id) {
    return
  }
  throw createWorkbookAgentServiceError({
    code: 'WORKBOOK_AGENT_REVIEW_ITEM_NOT_FOUND',
    message: 'Workbook agent change set was not found.',
    statusCode: 404,
    retryable: false,
  })
}

function assertWorkbookAgentApplyStillAuthorized(input: {
  sessionState: WorkbookAgentThreadState
  commandBundle: WorkbookAgentCommandBundle
  appliedBy: WorkbookAgentAppliedBy
  assertApplyStillAuthorized?: WorkbookAgentApplyAuthorityGuard | null | undefined
}): void {
  input.assertApplyStillAuthorized?.()
  if (input.appliedBy === 'user') {
    assertManualReviewItemStillCurrent({
      sessionState: input.sessionState,
      commandBundle: input.commandBundle,
    })
  }
}

export async function applyWorkbookAgentCommandBundleForSessionState(
  context: WorkbookAgentBundleApplicationContext,
  input: {
    sessionState: WorkbookAgentThreadState
    commandBundle: WorkbookAgentCommandBundle
    actorUserId: string
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null | undefined
    assertApplyStillAuthorized?: WorkbookAgentApplyAuthorityGuard | null | undefined
  },
): Promise<WorkbookAgentExecutionRecord> {
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
    if (input.sessionState.executionPolicy === 'autoApplySafe' && !context.autoApplyLowRiskEnabled) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_AUTO_APPLY_DISABLED',
        message: 'Automatic safe-apply is paused for this environment.',
        statusCode: 409,
        retryable: false,
      })
    }
    assertAutoApplyRolloutAllowed({
      context,
      documentId: input.sessionState.documentId,
      userId: input.actorUserId,
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
  assertWorkbookAgentApplyStillAuthorized({
    sessionState: input.sessionState,
    commandBundle: input.commandBundle,
    appliedBy: input.appliedBy,
    assertApplyStillAuthorized: input.assertApplyStillAuthorized,
  })
  const preview = await buildWorkbookAgentAuthoritativePreview({
    zeroSyncService: context.zeroSyncService,
    documentId: input.sessionState.documentId,
    bundle: selection.acceptedBundle,
  })
  assertWorkbookAgentApplyStillAuthorized({
    sessionState: input.sessionState,
    commandBundle: input.commandBundle,
    appliedBy: input.appliedBy,
    assertApplyStillAuthorized: input.assertApplyStillAuthorized,
  })
  const result = await context.zeroSyncService.applyAgentCommandBundle(input.sessionState.documentId, selection.acceptedBundle, preview, {
    userID: input.actorUserId,
    roles: ['editor'],
  })
  const executionRecord = buildWorkbookAgentExecutionRecord({
    bundle: selection.acceptedBundle,
    actorUserId: input.actorUserId,
    planText: collectPlanTextForTurn(input.sessionState, input.commandBundle.turnId),
    preview: result.preview,
    appliedRevision: result.revision,
    appliedBy: input.appliedBy,
    acceptedScope: selection.acceptedScope,
    now: context.now(),
  })
  await context.zeroSyncService.appendWorkbookAgentRun(executionRecord)
  input.sessionState.durable.executionRecords = [
    executionRecord,
    ...input.sessionState.durable.executionRecords.filter((record) => record.id !== executionRecord.id),
  ]
  await refreshAppliedWorkbookContext({
    zeroSyncService: context.zeroSyncService,
    sessionState: input.sessionState,
    turnId: input.commandBundle.turnId,
    commands: selection.acceptedBundle.commands,
  })
  const nextReviewItem =
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
              now: context.now(),
            }),
            input.sessionState,
          ),
        })
  replaceAppliedReviewItemIfCurrent({
    sessionState: input.sessionState,
    appliedBundleId: input.commandBundle.id,
    nextReviewItem,
  })
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
  context.touchSession(input.sessionState)
  return executionRecord
}

export async function applyWorkbookAgentToolBundleAutomatically(
  context: WorkbookAgentBundleApplicationContext & {
    persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>
    emitSnapshot: (threadId: string) => void
  },
  input: {
    sessionState: WorkbookAgentThreadState
    actorUserId: string
    bundle: WorkbookAgentCommandBundle
    assertApplyStillAuthorized?: WorkbookAgentApplyAuthorityGuard | null | undefined
  },
): Promise<WorkbookAgentExecutionRecord | null> {
  const executionRecord = await applyWorkbookAgentCommandBundleForSessionState(context, {
    sessionState: input.sessionState,
    commandBundle: input.bundle,
    actorUserId: input.actorUserId,
    appliedBy: 'auto',
    assertApplyStillAuthorized: input.assertApplyStillAuthorized,
  })
  await context.persistSessionState(input.sessionState)
  context.emitSnapshot(input.sessionState.threadId)
  return executionRecord.bundleId === input.bundle.id ? executionRecord : null
}

export async function finalizeWorkbookAgentPrivateTurnBundle(
  context: WorkbookAgentBundleApplicationContext & {
    resolveTurnActorUserId: (sessionState: WorkbookAgentThreadState, turnId: string) => string
  },
  input: {
    sessionState: WorkbookAgentThreadState
    turnId: string
    turnStatus: 'completed' | 'failed'
  },
): Promise<void> {
  const queuedBundle = input.sessionState.live.stagedPrivateBundleByTurn.get(input.turnId)
  if (!queuedBundle) {
    return
  }
  input.sessionState.live.stagedPrivateBundleByTurn.delete(input.turnId)
  if (input.turnStatus !== 'completed') {
    return
  }
  if (input.sessionState.live.activeTurnId !== input.turnId) {
    return
  }
  const actorUserId = context.resolveTurnActorUserId(input.sessionState, input.turnId)
  try {
    await applyWorkbookAgentCommandBundleForSessionState(context, {
      sessionState: input.sessionState,
      commandBundle: queuedBundle,
      actorUserId,
      appliedBy: 'auto',
      assertApplyStillAuthorized: () => {
        if (input.sessionState.live.activeTurnId === input.turnId) {
          return
        }
        throw createWorkbookAgentServiceError({
          code: 'WORKBOOK_AGENT_STALE_TURN_APPLY_BLOCKED',
          message: 'Automatic workbook apply was skipped because the assistant turn no longer owns the session.',
          statusCode: 409,
          retryable: false,
        })
      },
    })
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    input.sessionState.live.lastError = message
    input.sessionState.durable.entries = upsertEntry(
      input.sessionState.durable.entries,
      createSystemEntry(
        `system-auto-apply-failed:${queuedBundle.id}:${context.now()}`,
        input.turnId,
        `Automatic workbook apply failed: ${queuedBundle.summary}. ${message}`,
        createBundleRangeCitations(queuedBundle),
      ),
    )
    input.sessionState.live.status = 'failed'
  }
}
