import type { CodexServerNotification } from '@bilig/agent-api'
import type { WorkbookAgentStreamEvent, WorkbookAgentTextEntryKind } from '@bilig/contracts'
import { createSystemEntry, createTextTimelineEntry, mapThreadItemToEntry } from './workbook-agent-session-model.js'
import {
  type WorkbookAgentThreadState,
  normalizeCodexNotificationErrorMessage,
  removeEntry,
  upsertEntry,
} from './workbook-agent-service-shared.js'

export async function routeWorkbookAgentCodexNotification(input: {
  notification: CodexServerNotification
  listSessions: () => readonly WorkbookAgentThreadState[]
  tryGetSessionByThreadId: (threadId: string) => WorkbookAgentThreadState | null
  finalizeCompletedTurn?: (sessionState: WorkbookAgentThreadState, turnId: string, turnStatus: 'completed' | 'failed') => Promise<void>
  persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>
  emitSnapshot: (threadId: string) => void
  emit: (threadId: string, event: WorkbookAgentStreamEvent) => void
  now: () => number
}): Promise<void> {
  async function appendTextDelta(params: {
    threadId: string
    turnId: string
    itemId: string
    delta: string
    entryKind: WorkbookAgentTextEntryKind
  }): Promise<void> {
    const sessionState = input.tryGetSessionByThreadId(params.threadId)
    if (!sessionState) {
      return
    }
    const existing = sessionState.durable.entries.find((entry) => entry.id === params.itemId)
    sessionState.durable.entries = upsertEntry(
      sessionState.durable.entries,
      createTextTimelineEntry({
        id: params.itemId,
        kind: params.entryKind,
        turnId: params.turnId,
        text: `${existing?.text ?? ''}${params.delta}`,
        phase: existing?.phase ?? null,
        citations: existing?.citations ?? [],
      }),
    )
    await input.persistSessionState(sessionState)
    input.emit(sessionState.threadId, {
      type: 'entryTextDelta',
      entryKind: params.entryKind,
      itemId: params.itemId,
      turnId: params.turnId,
      delta: params.delta,
    })
  }

  switch (input.notification.method) {
    case 'thread/started':
      return
    case 'turn/started': {
      const sessionState = input.tryGetSessionByThreadId(input.notification.params.threadId)
      if (!sessionState) {
        return
      }
      sessionState.live.activeTurnId = input.notification.params.turn.id
      sessionState.live.status = 'inProgress'
      sessionState.live.lastError = null
      input.emitSnapshot(sessionState.threadId)
      return
    }
    case 'turn/completed': {
      const sessionState = input.tryGetSessionByThreadId(input.notification.params.threadId)
      if (!sessionState) {
        return
      }
      sessionState.live.activeTurnId = null
      sessionState.live.status = input.notification.params.turn.status === 'failed' || sessionState.live.lastError ? 'failed' : 'idle'
      if (input.notification.params.turn.error?.message) {
        sessionState.live.lastError = input.notification.params.turn.error.message
      }
      await input.finalizeCompletedTurn?.(
        sessionState,
        input.notification.params.turn.id,
        input.notification.params.turn.status === 'failed' ? 'failed' : 'completed',
      )
      sessionState.live.status = input.notification.params.turn.status === 'failed' || sessionState.live.lastError ? 'failed' : 'idle'
      if (input.notification.params.turn.error?.message) {
        sessionState.live.lastError = input.notification.params.turn.error.message
      }
      sessionState.live.promptByTurn.delete(input.notification.params.turn.id)
      sessionState.live.turnActorUserIdByTurn.delete(input.notification.params.turn.id)
      sessionState.live.turnContextByTurn.delete(input.notification.params.turn.id)
      sessionState.live.stagedPrivateBundleByTurn.delete(input.notification.params.turn.id)
      await input.persistSessionState(sessionState)
      input.emitSnapshot(sessionState.threadId)
      return
    }
    case 'item/started':
    case 'item/completed': {
      const sessionState = input.tryGetSessionByThreadId(input.notification.params.threadId)
      if (!sessionState) {
        return
      }
      const optimisticUserEntryId = sessionState.live.optimisticUserEntryIdByTurn.get(input.notification.params.turnId)
      if (input.notification.params.item.type === 'userMessage' && optimisticUserEntryId) {
        sessionState.durable.entries = removeEntry(sessionState.durable.entries, optimisticUserEntryId)
        sessionState.live.optimisticUserEntryIdByTurn.delete(input.notification.params.turnId)
      }
      sessionState.durable.entries = upsertEntry(
        sessionState.durable.entries,
        mapThreadItemToEntry(input.notification.params.item, input.notification.params.turnId),
      )
      await input.persistSessionState(sessionState)
      input.emitSnapshot(sessionState.threadId)
      return
    }
    case 'item/agentMessage/delta': {
      await appendTextDelta({
        ...input.notification.params,
        entryKind: 'assistant',
      })
      return
    }
    case 'item/plan/delta': {
      await appendTextDelta({
        ...input.notification.params,
        entryKind: 'plan',
      })
      return
    }
    case 'item/reasoning/delta': {
      await appendTextDelta({
        ...input.notification.params,
        entryKind: 'reasoning',
      })
      return
    }
    case 'error': {
      const message = normalizeCodexNotificationErrorMessage(input.notification)
      await Promise.all(
        input.listSessions().map(async (sessionState) => {
          sessionState.live.stagedPrivateBundleByTurn.clear()
          sessionState.live.lastError = message
          sessionState.live.status = 'failed'
          sessionState.durable.entries = upsertEntry(
            sessionState.durable.entries,
            createSystemEntry(`system-error:${input.now()}`, sessionState.live.activeTurnId, message),
          )
          await input.persistSessionState(sessionState)
          input.emitSnapshot(sessionState.threadId)
        }),
      )
      return
    }
  }
}
