import type { WorkbookAgentThreadSnapshot } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import { isCodexAppServerPoolBackpressureError } from './codex-app-server-pool.js'
import { startTurnBodySchema } from './workbook-agent-session-model.js'
import { buildSnapshot, cloneUiContext, type WorkbookAgentThreadState, upsertEntry } from './workbook-agent-service-shared.js'
import type { WorkbookAgentCodexRuntime } from './workbook-agent-codex-runtime.js'
import type { WorkbookAgentSessionRegistry } from './workbook-agent-session-registry.js'
import { startWorkbookAgentTurn } from './workbook-agent-turn-lifecycle.js'

interface StartWorkbookAgentServiceTurnContext {
  readonly codexRuntime: WorkbookAgentCodexRuntime
  readonly sessionRegistry: WorkbookAgentSessionRegistry
  getAuthorizedSession(documentId: string, threadId: string, userId: string): Promise<WorkbookAgentThreadState>
  assertTurnQuota(documentId: string, actorUserId: string): void
  persistSessionState(sessionState: WorkbookAgentThreadState): Promise<void>
}

export async function startWorkbookAgentServiceTurn(
  context: StartWorkbookAgentServiceTurnContext,
  input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  },
): Promise<WorkbookAgentThreadSnapshot> {
  const parsed = startTurnBodySchema.parse(input.body)
  const sessionState = await context.getAuthorizedSession(input.documentId, input.threadId, input.session.userID)
  if (sessionState.live.activeTurnId) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_TURN_ALREADY_RUNNING',
      message: 'Finish or interrupt the current assistant turn before starting another one.',
      statusCode: 409,
      retryable: false,
    })
  }
  context.assertTurnQuota(input.documentId, input.session.userID)
  if (parsed.context) {
    sessionState.durable.context = parsed.context
  }
  const turnContext = cloneUiContext(sessionState.durable.context)
  const codexClient = await context.codexRuntime.getClient()
  let turn
  try {
    turn = await codexClient.turnStart({
      threadId: sessionState.threadId,
      prompt: parsed.prompt,
    })
  } catch (error) {
    if (isCodexAppServerPoolBackpressureError(error)) {
      context.sessionRegistry.incrementCounter('turnBackpressureCount')
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
  context.sessionRegistry.touch(sessionState)
  await context.persistSessionState(sessionState)
  context.sessionRegistry.emitSnapshot(sessionState.threadId)
  return buildSnapshot(sessionState)
}
