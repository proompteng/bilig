import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export interface WorkbookAgentStartedTurnInput {
  readonly turnId: string
  readonly prompt: string
  readonly actorUserId: string
  readonly context: WorkbookAgentUiContext | null
  readonly optimisticEntryId: string
}

export interface WorkbookAgentCompletedTurnInput {
  readonly turnId: string
  readonly status: 'completed' | 'failed'
  readonly errorMessage: string | null
}

export function startWorkbookAgentTurn(sessionState: WorkbookAgentThreadState, input: WorkbookAgentStartedTurnInput): void {
  sessionState.live.optimisticUserEntryIdByTurn.set(input.turnId, input.optimisticEntryId)
  sessionState.live.promptByTurn.set(input.turnId, input.prompt)
  sessionState.live.turnActorUserIdByTurn.set(input.turnId, input.actorUserId)
  sessionState.live.turnContextByTurn.set(input.turnId, input.context)
  sessionState.live.activeTurnId = input.turnId
  sessionState.live.status = 'inProgress'
  sessionState.live.lastError = null
}

export function markWorkbookAgentTurnStarted(sessionState: WorkbookAgentThreadState, turnId: string): void {
  const activeTurnId = sessionState.live.activeTurnId
  const knownTurn =
    sessionState.live.promptByTurn.has(turnId) ||
    sessionState.live.turnActorUserIdByTurn.has(turnId) ||
    sessionState.live.turnContextByTurn.has(turnId) ||
    sessionState.live.optimisticUserEntryIdByTurn.has(turnId) ||
    sessionState.durable.entries.some((entry) => entry.turnId === turnId && entry.id.startsWith('optimistic-user:'))
  if (activeTurnId && activeTurnId !== turnId && knownTurn) {
    clearWorkbookAgentTurnState(sessionState, turnId)
    return
  }
  sessionState.live.activeTurnId = turnId
  sessionState.live.status = 'inProgress'
  sessionState.live.lastError = null
}

function clearWorkbookAgentTurnState(sessionState: WorkbookAgentThreadState, turnId: string): void {
  sessionState.live.promptByTurn.delete(turnId)
  sessionState.live.turnActorUserIdByTurn.delete(turnId)
  sessionState.live.turnContextByTurn.delete(turnId)
  sessionState.live.stagedPrivateBundleByTurn.delete(turnId)
  sessionState.live.optimisticUserEntryIdByTurn.delete(turnId)
}

function resolveWorkbookAgentRuntimeErrorTurnId(sessionState: WorkbookAgentThreadState): string | null {
  if (sessionState.live.activeTurnId) {
    return sessionState.live.activeTurnId
  }
  const liveTurnId =
    Array.from(
      new Set([
        ...sessionState.live.promptByTurn.keys(),
        ...sessionState.live.turnActorUserIdByTurn.keys(),
        ...sessionState.live.turnContextByTurn.keys(),
        ...sessionState.live.stagedPrivateBundleByTurn.keys(),
        ...sessionState.live.optimisticUserEntryIdByTurn.keys(),
      ]),
    ).at(-1) ?? null
  if (liveTurnId) {
    return liveTurnId
  }
  return sessionState.durable.entries.findLast((entry) => entry.id.startsWith('optimistic-user:'))?.turnId ?? null
}

export function completeWorkbookAgentTurn(sessionState: WorkbookAgentThreadState, input: WorkbookAgentCompletedTurnInput): void {
  const completedActiveTurn = sessionState.live.activeTurnId === input.turnId
  if (completedActiveTurn) {
    sessionState.live.activeTurnId = null
    if (input.errorMessage) {
      sessionState.live.lastError = input.errorMessage
    }
    sessionState.live.status = input.status === 'failed' || sessionState.live.lastError ? 'failed' : 'idle'
    clearWorkbookAgentTurnState(sessionState, input.turnId)
    return
  }

  clearWorkbookAgentTurnState(sessionState, input.turnId)
  if (sessionState.live.activeTurnId) {
    sessionState.live.status = 'inProgress'
    return
  }

  if (input.errorMessage) {
    sessionState.live.lastError = input.errorMessage
  }
  sessionState.live.status = input.status === 'failed' || sessionState.live.lastError ? 'failed' : 'idle'
}

export function failWorkbookAgentRuntime(sessionState: WorkbookAgentThreadState, message: string): string | null {
  const failedTurnId = resolveWorkbookAgentRuntimeErrorTurnId(sessionState)
  if (failedTurnId) {
    clearWorkbookAgentTurnState(sessionState, failedTurnId)
  }
  sessionState.live.activeTurnId = null
  sessionState.live.stagedPrivateBundleByTurn.clear()
  sessionState.live.lastError = message
  sessionState.live.status = 'failed'
  return failedTurnId
}

export function assertWorkbookAgentToolCallOwnsTurn(sessionState: WorkbookAgentThreadState, turnId: string): void {
  const activeTurnId = sessionState.live.activeTurnId
  if (activeTurnId === turnId) {
    return
  }
  throw createWorkbookAgentServiceError({
    code: 'WORKBOOK_AGENT_STALE_TOOL_CALL',
    message:
      activeTurnId === null
        ? 'Rejecting workbook tool call because the assistant turn is no longer active.'
        : `Rejecting workbook tool call for stale turn ${turnId}; active turn is ${activeTurnId}.`,
    statusCode: 409,
    retryable: false,
  })
}
