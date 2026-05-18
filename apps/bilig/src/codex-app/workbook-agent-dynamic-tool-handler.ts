import {
  appendWorkbookAgentCommandToBundle,
  toWorkbookAgentCommandBundle,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolCallResult,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import type { WorkbookAgentThreadSnapshot } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { handleWorkbookAgentToolCall, type WorkbookAgentStartWorkflowRequest } from './workbook-agent-tools.js'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import { cloneUiContext, type WorkbookAgentThreadState, toContextRef } from './workbook-agent-service-shared.js'
import { inspectWorkbookRange, normalizeWorkbookAgentUiContext } from './workbook-agent-inspection.js'
import { selectWorkbookRenderedReadback } from './workbook-agent-rendered-readback.js'
import {
  hasRenderedContext,
  shouldWaitForRenderedTool,
  waitForWorkbookAgentRenderedContext,
} from './workbook-agent-rendered-context-wait.js'

function firstRenderedVerificationRange(bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>) {
  const targetRange = bundle.affectedRanges.find((range) => range.role === 'target') ?? null
  return targetRange
    ? {
        sheetName: targetRange.sheetName,
        startAddress: targetRange.startAddress,
        endAddress: targetRange.endAddress,
      }
    : null
}

export function createWorkbookAgentDynamicToolHandler(input: {
  zeroSyncService: ZeroSyncService
  now: () => number
  getSessionByThreadId: (threadId: string) => WorkbookAgentThreadState
  resolveTurnActorUserId: (sessionState: WorkbookAgentThreadState, turnId: string) => string
  resolveTurnContext: (sessionState: WorkbookAgentThreadState, turnId: string) => WorkbookAgentThreadState['durable']['context']
  stageReviewBundle: (
    sessionState: WorkbookAgentThreadState,
    turnId: string,
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
  ) => void
  shouldApplyToolBundleImmediately: (
    sessionState: WorkbookAgentThreadState,
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
  ) => boolean
  applyToolBundleAutomatically: (input: {
    sessionState: WorkbookAgentThreadState
    actorUserId: string
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>
  }) => Promise<WorkbookAgentExecutionRecord | null>
  persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>
  emitSnapshot: (threadId: string) => void
  startWorkflow: (input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: WorkbookAgentStartWorkflowRequest & {
      context?: WorkbookAgentThreadState['durable']['context']
    }
  }) => Promise<WorkbookAgentThreadSnapshot>
}): (request: CodexDynamicToolCallRequest) => Promise<CodexDynamicToolCallResult> {
  return async (request: CodexDynamicToolCallRequest) => {
    const sessionState = input.getSessionByThreadId(request.threadId)
    const requestActorUserId = input.resolveTurnActorUserId(sessionState, request.turnId)
    let requestContext: WorkbookAgentThreadState['durable']['context'] = null

    const refreshRequestContext = async (): Promise<WorkbookAgentThreadState['durable']['context']> => {
      const rawRequestContext = input.resolveTurnContext(sessionState, request.turnId)
      const normalizedContext = await input.zeroSyncService.inspectWorkbook(sessionState.documentId, (runtime) =>
        normalizeWorkbookAgentUiContext(runtime, rawRequestContext),
      )
      requestContext = normalizedContext
      if (JSON.stringify(rawRequestContext) !== JSON.stringify(normalizedContext)) {
        sessionState.durable.context = cloneUiContext(normalizedContext)
        sessionState.live.turnContextByTurn.set(request.turnId, cloneUiContext(normalizedContext))
        await input.persistSessionState(sessionState)
        input.emitSnapshot(sessionState.threadId)
      }
      return normalizedContext
    }

    const waitForRenderedContext = async (
      minRevision: number,
      isReady?: (context: WorkbookAgentThreadState['durable']['context']) => Promise<boolean>,
    ): Promise<WorkbookAgentThreadState['durable']['context']> => {
      return await waitForWorkbookAgentRenderedContext({
        minRevision,
        refreshContext: refreshRequestContext,
        ...(isReady ? { isReady } : {}),
      })
    }

    const renderedVerificationRangeMatches = async (
      latestContext: WorkbookAgentThreadState['durable']['context'],
      bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
      minRevision: number,
    ): Promise<boolean> => {
      const targetRange = firstRenderedVerificationRange(bundle)
      if (!targetRange) {
        return true
      }
      return await input.zeroSyncService.inspectWorkbook(sessionState.documentId, (runtime) => {
        const normalizedContext = normalizeWorkbookAgentUiContext(runtime, latestContext)
        const authoritativeReadback = inspectWorkbookRange(runtime, targetRange)
        const authoritativeRows = authoritativeReadback.rows.filter(Array.isArray) as readonly (readonly unknown[])[]
        const renderedReadback = selectWorkbookRenderedReadback({
          renderedContext: normalizedContext?.rendered,
          requestedRange: targetRange,
          authoritativeRows,
          minRevision,
        })
        return renderedReadback.matched === true
      })
    }

    requestContext = await refreshRequestContext()
    if (shouldWaitForRenderedTool(request.tool)) {
      const headRevision = await input.zeroSyncService.getWorkbookHeadRevision(sessionState.documentId)
      requestContext = await waitForRenderedContext(headRevision)
    }

    return handleWorkbookAgentToolCall(
      {
        documentId: sessionState.documentId,
        session: {
          userID: requestActorUserId,
          roles: ['editor'],
        },
        get uiContext() {
          return requestContext
        },
        zeroSyncService: input.zeroSyncService,
        updateUiContext: async (nextContext) => {
          sessionState.durable.context = cloneUiContext(nextContext)
          sessionState.live.turnContextByTurn.set(request.turnId, cloneUiContext(nextContext))
          await input.persistSessionState(sessionState)
          input.emitSnapshot(sessionState.threadId)
        },
        awaitRenderedRevision: async (revision) => {
          requestContext = await waitForRenderedContext(revision)
        },
        stageCommand: async (command) => {
          const currentReviewBundle = sessionState.durable.reviewQueueItems[0]
            ? toWorkbookAgentCommandBundle(sessionState.durable.reviewQueueItems[0])
            : null
          const previousBundle = sessionState.scope === 'private' ? null : currentReviewBundle
          const baseRevision = await input.zeroSyncService.getWorkbookHeadRevision(sessionState.documentId)
          const bundle = appendWorkbookAgentCommandToBundle({
            previousBundle,
            documentId: sessionState.documentId,
            threadId: sessionState.threadId,
            turnId: request.turnId,
            goalText: sessionState.live.promptByTurn.get(request.turnId) ?? 'Update workbook from assistant request',
            baseRevision,
            context: toContextRef(requestContext),
            command,
            now: input.now(),
          })
          if (input.shouldApplyToolBundleImmediately(sessionState, bundle)) {
            const executionRecord = await input.applyToolBundleAutomatically({
              sessionState,
              actorUserId: requestActorUserId,
              bundle,
            })
            if (executionRecord && hasRenderedContext(requestContext)) {
              requestContext = await waitForRenderedContext(
                executionRecord.appliedRevision,
                async (latestContext) => await renderedVerificationRangeMatches(latestContext, bundle, executionRecord.appliedRevision),
              )
            }
            return {
              bundle,
              executionRecord,
            }
          }
          if (sessionState.scope === 'private') {
            throw createWorkbookAgentServiceError({
              code: 'WORKBOOK_AGENT_PRIVATE_EXECUTION_BLOCKED',
              message:
                'Private workbook threads execute changes directly and do not queue review items under the current execution policy.',
              statusCode: 409,
              retryable: false,
            })
          }
          input.stageReviewBundle(sessionState, request.turnId, bundle)
          await input.persistSessionState(sessionState)
          input.emitSnapshot(sessionState.threadId)
          return {
            bundle,
            executionRecord: null,
            disposition: 'reviewQueued',
          }
        },
        startWorkflow: async (workflowRequest: WorkbookAgentStartWorkflowRequest) => {
          const previousRunIds = new Set(sessionState.durable.workflowRuns.map((run) => run.runId))
          const nextSnapshot = await input.startWorkflow({
            documentId: sessionState.documentId,
            threadId: sessionState.threadId,
            session: {
              userID: requestActorUserId,
              roles: ['editor'],
            },
            body: {
              ...workflowRequest,
              ...(requestContext ? { context: requestContext } : {}),
            },
          })
          const nextRun =
            nextSnapshot.workflowRuns.find((run) => !previousRunIds.has(run.runId)) ??
            nextSnapshot.workflowRuns.find((run) => run.workflowTemplate === workflowRequest.workflowTemplate) ??
            null
          if (!nextRun) {
            throw new Error(`Workflow run not found after starting ${workflowRequest.workflowTemplate}`)
          }
          return nextRun
        },
      },
      request,
    )
  }
}
