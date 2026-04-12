import type { CodexServerNotification } from "@bilig/agent-api";
import type { WorkbookAgentStreamEvent, WorkbookAgentTextEntryKind } from "@bilig/contracts";
import {
  createSystemEntry,
  createTextTimelineEntry,
  mapThreadItemToEntry,
} from "./workbook-agent-session-model.js";
import {
  type WorkbookAgentSessionState,
  normalizeCodexNotificationErrorMessage,
  removeEntry,
  upsertEntry,
} from "./workbook-agent-service-shared.js";

export async function routeWorkbookAgentCodexNotification(input: {
  notification: CodexServerNotification;
  listSessions: () => readonly WorkbookAgentSessionState[];
  tryGetSessionByThreadId: (threadId: string) => WorkbookAgentSessionState | null;
  persistSessionState: (sessionState: WorkbookAgentSessionState) => Promise<void>;
  emitSnapshot: (threadId: string) => void;
  emit: (threadId: string, event: WorkbookAgentStreamEvent) => void;
  now: () => number;
}): Promise<void> {
  function appendTextDelta(params: {
    threadId: string;
    turnId: string;
    itemId: string;
    delta: string;
    entryKind: WorkbookAgentTextEntryKind;
  }): void {
    const sessionState = input.tryGetSessionByThreadId(params.threadId);
    if (!sessionState) {
      return;
    }
    const existing = sessionState.snapshot.entries.find((entry) => entry.id === params.itemId);
    sessionState.snapshot.entries = upsertEntry(
      sessionState.snapshot.entries,
      createTextTimelineEntry({
        id: params.itemId,
        kind: params.entryKind,
        turnId: params.turnId,
        text: `${existing?.text ?? ""}${params.delta}`,
        phase: existing?.phase ?? null,
        citations: existing?.citations ?? [],
      }),
    );
    input.emit(sessionState.threadId, {
      type: "entryTextDelta",
      entryKind: params.entryKind,
      itemId: params.itemId,
      turnId: params.turnId,
      delta: params.delta,
    });
  }

  switch (input.notification.method) {
    case "thread/started":
      return;
    case "turn/started": {
      const sessionState = input.tryGetSessionByThreadId(input.notification.params.threadId);
      if (!sessionState) {
        return;
      }
      sessionState.snapshot.activeTurnId = input.notification.params.turn.id;
      sessionState.snapshot.status = "inProgress";
      sessionState.snapshot.lastError = null;
      input.emitSnapshot(sessionState.threadId);
      return;
    }
    case "turn/completed": {
      const sessionState = input.tryGetSessionByThreadId(input.notification.params.threadId);
      if (!sessionState) {
        return;
      }
      sessionState.snapshot.activeTurnId = null;
      sessionState.snapshot.status =
        input.notification.params.turn.status === "failed" ? "failed" : "idle";
      sessionState.snapshot.lastError = input.notification.params.turn.error?.message ?? null;
      sessionState.promptByTurn.delete(input.notification.params.turn.id);
      sessionState.turnActorUserIdByTurn.delete(input.notification.params.turn.id);
      sessionState.turnContextByTurn.delete(input.notification.params.turn.id);
      await input.persistSessionState(sessionState);
      input.emitSnapshot(sessionState.threadId);
      return;
    }
    case "item/started":
    case "item/completed": {
      const sessionState = input.tryGetSessionByThreadId(input.notification.params.threadId);
      if (!sessionState) {
        return;
      }
      const optimisticUserEntryId = sessionState.optimisticUserEntryIdByTurn.get(
        input.notification.params.turnId,
      );
      if (input.notification.params.item.type === "userMessage" && optimisticUserEntryId) {
        sessionState.snapshot.entries = removeEntry(
          sessionState.snapshot.entries,
          optimisticUserEntryId,
        );
        sessionState.optimisticUserEntryIdByTurn.delete(input.notification.params.turnId);
      }
      sessionState.snapshot.entries = upsertEntry(
        sessionState.snapshot.entries,
        mapThreadItemToEntry(input.notification.params.item, input.notification.params.turnId),
      );
      await input.persistSessionState(sessionState);
      input.emitSnapshot(sessionState.threadId);
      return;
    }
    case "item/agentMessage/delta": {
      appendTextDelta({
        ...input.notification.params,
        entryKind: "assistant",
      });
      return;
    }
    case "item/plan/delta": {
      appendTextDelta({
        ...input.notification.params,
        entryKind: "plan",
      });
      return;
    }
    case "item/reasoning/delta": {
      appendTextDelta({
        ...input.notification.params,
        entryKind: "reasoning",
      });
      return;
    }
    case "error": {
      const message = normalizeCodexNotificationErrorMessage(input.notification);
      await Promise.all(
        input.listSessions().map(async (sessionState) => {
          sessionState.snapshot.lastError = message;
          sessionState.snapshot.status = "failed";
          sessionState.snapshot.entries = upsertEntry(
            sessionState.snapshot.entries,
            createSystemEntry(
              `system-error:${input.now()}`,
              sessionState.snapshot.activeTurnId,
              message,
            ),
          );
          await input.persistSessionState(sessionState);
          input.emitSnapshot(sessionState.threadId);
        }),
      );
      return;
    }
  }
}
