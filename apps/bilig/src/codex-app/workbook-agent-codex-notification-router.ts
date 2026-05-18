import type { CodexServerNotification } from '@bilig/agent-api'
import type { WorkbookAgentStreamEvent, WorkbookAgentTextEntryKind } from '@bilig/contracts'
import {
  appendCommandExecutionOutput,
  createCommandExecutionOutputEntry,
  createSystemEntry,
  createTextTimelineEntry,
  decodeCommandExecutionOutput,
  mapThreadItemToEntry,
} from './workbook-agent-session-model.js'
import {
  type WorkbookAgentThreadState,
  normalizeCodexNotificationErrorMessage,
  removeEntry,
  upsertEntry,
} from './workbook-agent-service-shared.js'

function clearWorkbookAgentLiveTurnState(sessionState: WorkbookAgentThreadState, turnId: string): void {
  sessionState.live.promptByTurn.delete(turnId)
  sessionState.live.turnActorUserIdByTurn.delete(turnId)
  sessionState.live.turnContextByTurn.delete(turnId)
  sessionState.live.stagedPrivateBundleByTurn.delete(turnId)
  sessionState.live.optimisticUserEntryIdByTurn.delete(turnId)
}

function resolveRuntimeErrorTurnId(sessionState: WorkbookAgentThreadState): string | null {
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

  const notification = input.notification

  switch (notification.method) {
    case 'thread/started':
      return
    case 'turn/started': {
      const sessionState = input.tryGetSessionByThreadId(notification.params.threadId)
      if (!sessionState) {
        return
      }
      sessionState.live.activeTurnId = notification.params.turn.id
      sessionState.live.status = 'inProgress'
      sessionState.live.lastError = null
      input.emitSnapshot(sessionState.threadId)
      return
    }
    case 'turn/completed': {
      const sessionState = input.tryGetSessionByThreadId(notification.params.threadId)
      if (!sessionState) {
        return
      }
      const completedTurnId = notification.params.turn.id
      const completedActiveTurn = sessionState.live.activeTurnId === completedTurnId
      if (completedActiveTurn) {
        sessionState.live.activeTurnId = null
        sessionState.live.status = notification.params.turn.status === 'failed' || sessionState.live.lastError ? 'failed' : 'idle'
      }
      if (notification.params.turn.error?.message) {
        sessionState.live.lastError = notification.params.turn.error.message
      }
      await input.finalizeCompletedTurn?.(
        sessionState,
        completedTurnId,
        notification.params.turn.status === 'failed' ? 'failed' : 'completed',
      )
      if (completedActiveTurn) {
        sessionState.live.status = notification.params.turn.status === 'failed' || sessionState.live.lastError ? 'failed' : 'idle'
      } else if (sessionState.live.activeTurnId) {
        sessionState.live.status = 'inProgress'
      } else {
        sessionState.live.status = notification.params.turn.status === 'failed' || sessionState.live.lastError ? 'failed' : 'idle'
      }
      if (notification.params.turn.error?.message) {
        sessionState.live.lastError = notification.params.turn.error.message
      }
      clearWorkbookAgentLiveTurnState(sessionState, completedTurnId)
      await input.persistSessionState(sessionState)
      input.emitSnapshot(sessionState.threadId)
      return
    }
    case 'item/started':
    case 'item/completed': {
      const params = notification.params
      const sessionState = input.tryGetSessionByThreadId(params.threadId)
      if (!sessionState) {
        return
      }
      const optimisticUserEntryId = sessionState.live.optimisticUserEntryIdByTurn.get(params.turnId)
      if (params.item.type === 'userMessage' && optimisticUserEntryId) {
        sessionState.durable.entries = removeEntry(sessionState.durable.entries, optimisticUserEntryId)
        sessionState.live.optimisticUserEntryIdByTurn.delete(params.turnId)
      }
      const existingEntry = sessionState.durable.entries.find((entry) => entry.id === params.item.id)
      const mappedEntry = mapThreadItemToEntry(params.item, params.turnId)
      sessionState.durable.entries = upsertEntry(
        sessionState.durable.entries,
        params.item.type === 'commandExecution' && mappedEntry.outputText === null && existingEntry?.outputText
          ? {
              ...mappedEntry,
              outputText: existingEntry.outputText,
            }
          : mappedEntry,
      )
      await input.persistSessionState(sessionState)
      input.emitSnapshot(sessionState.threadId)
      return
    }
    case 'item/agentMessage/delta': {
      await appendTextDelta({
        ...notification.params,
        entryKind: 'assistant',
      })
      return
    }
    case 'item/plan/delta': {
      await appendTextDelta({
        ...notification.params,
        entryKind: 'plan',
      })
      return
    }
    case 'item/reasoning/delta':
    case 'item/reasoning/textDelta':
    case 'item/reasoning/summaryTextDelta': {
      await appendTextDelta({
        ...notification.params,
        entryKind: 'reasoning',
      })
      return
    }
    case 'item/reasoning/summaryPartAdded':
      return
    case 'item/commandExecution/outputDelta': {
      const params = notification.params
      const sessionState = input.tryGetSessionByThreadId(params.threadId)
      if (!sessionState) {
        return
      }
      const delta = decodeCommandExecutionOutput(params.delta)
      const existing = sessionState.durable.entries.find((entry) => entry.id === params.itemId)
      sessionState.durable.entries = upsertEntry(
        sessionState.durable.entries,
        existing
          ? appendCommandExecutionOutput(existing, delta)
          : createCommandExecutionOutputEntry({
              id: params.itemId,
              turnId: params.turnId,
              outputText: delta,
            }),
      )
      await input.persistSessionState(sessionState)
      input.emit(sessionState.threadId, {
        type: 'entryToolOutputDelta',
        itemId: params.itemId,
        turnId: params.turnId,
        delta,
      })
      return
    }
    case 'item/commandExecution/terminalInteraction':
      return
    case 'error': {
      const message = normalizeCodexNotificationErrorMessage(input.notification)
      await Promise.all(
        input.listSessions().map(async (sessionState) => {
          const failedTurnId = resolveRuntimeErrorTurnId(sessionState)
          if (failedTurnId) {
            clearWorkbookAgentLiveTurnState(sessionState, failedTurnId)
          }
          sessionState.live.activeTurnId = null
          sessionState.live.stagedPrivateBundleByTurn.clear()
          sessionState.live.lastError = message
          sessionState.live.status = 'failed'
          sessionState.durable.entries = upsertEntry(
            sessionState.durable.entries,
            createSystemEntry(`system-error:${input.now()}`, failedTurnId, message),
          )
          await input.persistSessionState(sessionState)
          input.emitSnapshot(sessionState.threadId)
        }),
      )
      return
    }
  }
}
