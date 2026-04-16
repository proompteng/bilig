import type { AgentEvent } from '@bilig/agent-api'

export interface LocalAgentEventQueueSessionState {
  documentId: string
  eventBacklog: AgentEvent[]
  eventFlushScheduled: boolean
}

export interface LocalAgentEventQueueContext<SessionState extends LocalAgentEventQueueSessionState = LocalAgentEventQueueSessionState> {
  getSession(documentId: string): SessionState | undefined
  listeners: ReadonlySet<(event: AgentEvent) => void>
  schedule?(callback: () => void): void
}

export function removeQueuedSubscriptionEvents<SessionState extends LocalAgentEventQueueSessionState>(
  session: SessionState,
  subscriptionId: string,
): void {
  session.eventBacklog = session.eventBacklog.filter((event) => {
    return event.kind !== 'rangeChanged' || event.subscriptionId !== subscriptionId
  })
}

export function queueLocalAgentEvent<SessionState extends LocalAgentEventQueueSessionState>(
  context: LocalAgentEventQueueContext<SessionState>,
  documentId: string,
  event: AgentEvent,
): void {
  const session = context.getSession(documentId)
  if (!session) {
    return
  }
  session.eventBacklog.push(event)
  if (session.eventFlushScheduled) {
    return
  }
  session.eventFlushScheduled = true
  const schedule =
    context.schedule === undefined
      ? (callback: () => void) => setImmediate(callback)
      : (callback: () => void) => context.schedule?.(callback)
  schedule(() => {
    session.eventFlushScheduled = false
    flushQueuedLocalAgentEvents(context, documentId)
  })
}

export function flushQueuedLocalAgentEvents<SessionState extends LocalAgentEventQueueSessionState>(
  context: LocalAgentEventQueueContext<SessionState>,
  documentId: string,
): void {
  const session = context.getSession(documentId)
  if (!session || session.eventBacklog.length === 0 || context.listeners.size === 0) {
    return
  }
  const pending = session.eventBacklog.splice(0)
  pending.forEach((event) => {
    context.listeners.forEach((listener) => listener(event))
  })
}
