import type { CodexServerNotification } from "@bilig/agent-api";
import type { WorkbookAgentStreamEvent, WorkbookAgentTimelineEntry } from "@bilig/contracts";
import { createSystemEntry, mapThreadItemToEntry } from "./workbook-agent-session-model.js";
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
      const params = input.notification.params;
      const sessionState = input.tryGetSessionByThreadId(params.threadId);
      if (!sessionState) {
        return;
      }
      const existing =
        sessionState.snapshot.entries.find((entry) => entry.id === params.itemId) ??
        ({
          id: params.itemId,
          kind: "assistant",
          turnId: params.turnId,
          text: "",
          phase: null,
          toolName: null,
          toolStatus: null,
          argumentsText: null,
          outputText: null,
          success: null,
          citations: [],
        } satisfies WorkbookAgentTimelineEntry);
      sessionState.snapshot.entries = upsertEntry(sessionState.snapshot.entries, {
        ...existing,
        text: `${existing.text ?? ""}${params.delta}`,
      });
      input.emit(sessionState.threadId, {
        type: "assistantDelta",
        itemId: params.itemId,
        delta: params.delta,
      });
      return;
    }
    case "item/plan/delta": {
      const params = input.notification.params;
      const sessionState = input.tryGetSessionByThreadId(params.threadId);
      if (!sessionState) {
        return;
      }
      const existing =
        sessionState.snapshot.entries.find((entry) => entry.id === params.itemId) ??
        ({
          id: params.itemId,
          kind: "plan",
          turnId: params.turnId,
          text: "",
          phase: null,
          toolName: null,
          toolStatus: null,
          argumentsText: null,
          outputText: null,
          success: null,
          citations: [],
        } satisfies WorkbookAgentTimelineEntry);
      sessionState.snapshot.entries = upsertEntry(sessionState.snapshot.entries, {
        ...existing,
        text: `${existing.text ?? ""}${params.delta}`,
      });
      input.emit(sessionState.threadId, {
        type: "planDelta",
        itemId: params.itemId,
        delta: params.delta,
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
