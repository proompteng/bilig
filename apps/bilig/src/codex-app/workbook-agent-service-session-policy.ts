import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export function resolveWorkbookAgentActiveTurnActorUserId(sessionState: WorkbookAgentThreadState): string | null {
  const activeTurnId = sessionState.live.activeTurnId
  if (!activeTurnId) {
    return null
  }
  return sessionState.live.turnActorUserIdByTurn.get(activeTurnId) ?? sessionState.userId
}

export function summarizeWorkbookAgentActiveTurnCounts(
  sessions: readonly WorkbookAgentThreadState[],
  input: { actorUserId: string; documentId: string },
): {
  readonly activeTurnsForUser: number
  readonly activeTurnsForDocument: number
} {
  const activeSessions = sessions.filter(
    (sessionState) => sessionState.live.activeTurnId !== null && sessionState.live.status === 'inProgress',
  )
  return {
    activeTurnsForUser: activeSessions.filter(
      (sessionState) => resolveWorkbookAgentActiveTurnActorUserId(sessionState) === input.actorUserId,
    ).length,
    activeTurnsForDocument: activeSessions.filter((sessionState) => sessionState.documentId === input.documentId).length,
  }
}

export function assertWorkbookAgentTurnQuota(input: {
  readonly sessions: readonly WorkbookAgentThreadState[]
  readonly documentId: string
  readonly actorUserId: string
  readonly maxActiveTurnsPerUser: number
  readonly maxActiveTurnsPerDocument: number
}): void {
  const counts = summarizeWorkbookAgentActiveTurnCounts(input.sessions, {
    actorUserId: input.actorUserId,
    documentId: input.documentId,
  })
  if (counts.activeTurnsForUser >= input.maxActiveTurnsPerUser) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_USER_TURN_QUOTA_EXCEEDED',
      message: 'Workbook assistant is already running too many turns for this user. Retry once an in-flight turn finishes.',
      statusCode: 429,
      retryable: true,
    })
  }
  if (counts.activeTurnsForDocument >= input.maxActiveTurnsPerDocument) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_DOCUMENT_TURN_QUOTA_EXCEEDED',
      message: 'Workbook assistant is already running too many turns for this document. Retry once an in-flight turn finishes.',
      statusCode: 429,
      retryable: true,
    })
  }
}

export function chooseWorkbookAgentEvictionCandidates(
  sessions: readonly WorkbookAgentThreadState[],
  subscribers: ReadonlyMap<string, ReadonlySet<unknown>>,
): WorkbookAgentThreadState[] {
  return sessions
    .filter((sessionState) => {
      const listeners = subscribers.get(sessionState.threadId)
      return sessionState.live.status === 'idle' && (!listeners || listeners.size === 0)
    })
    .toSorted((left, right) => left.live.lastAccessedAt - right.live.lastAccessedAt)
}
