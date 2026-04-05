import type { SpreadsheetEngine } from "@bilig/core";
import type { EngineEvent, WorkbookSnapshot } from "@bilig/protocol";
import type { AgentEvent, AgentResponse } from "@bilig/agent-api";
import type { WorksheetAgentRequest } from "./agent-routing.js";
import type {
  LocalAgentRangeSubscriptionState,
  LocalAgentSessionContainer,
  LocalAgentSessionState,
} from "./local-agent-session-store.js";
import {
  cellCountForRange,
  collectChangedAddressesForEvent,
  getRangeBounds,
  iterateRange,
} from "./range-subscription-utils.js";

export interface LocalWorksheetSessionState extends LocalAgentSessionContainer {
  documentId: string;
  engine: SpreadsheetEngine;
  agentSessions: Map<string, LocalAgentSessionState>;
  agentSubscriptions: Map<string, LocalAgentRangeSubscriptionState>;
}

export interface LocalWorksheetAgentHandlerContext<
  SessionState extends LocalWorksheetSessionState = LocalWorksheetSessionState,
> {
  largeRangeSubscriptionThreshold: number;
  agentSubscriptionOwners: Map<string, string>;
  getSessionByAgentSessionId(sessionId: string): SessionState;
  getCachedSnapshot(session: SessionState): WorkbookSnapshot;
  importSnapshot(session: SessionState, snapshot: WorkbookSnapshot): void;
  removeAgentSubscription(session: SessionState, sessionId: string, subscriptionId: string): void;
  queueAgentEvent(documentId: string, event: AgentEvent): void;
}

export function handleLocalWorksheetAgentRequest<SessionState extends LocalWorksheetSessionState>(
  context: LocalWorksheetAgentHandlerContext<SessionState>,
  request: WorksheetAgentRequest,
): AgentResponse {
  switch (request.kind) {
    case "readRange": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      return {
        kind: "rangeValues",
        id: request.id,
        values: engine.getRangeValues(request.range),
      };
    }
    case "writeRange": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.setRangeValues(request.range, request.values);
      return { kind: "ok", id: request.id };
    }
    case "setRangeFormulas": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.setRangeFormulas(request.range, request.formulas);
      return { kind: "ok", id: request.id };
    }
    case "setRangeStyle": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.setRangeStyle(request.range, request.patch);
      return { kind: "ok", id: request.id };
    }
    case "clearRangeStyle": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.clearRangeStyle(request.range, request.fields);
      return { kind: "ok", id: request.id };
    }
    case "setRangeNumberFormat": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.setRangeNumberFormat(request.range, request.format);
      return { kind: "ok", id: request.id };
    }
    case "clearRangeNumberFormat": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.clearRangeNumberFormat(request.range);
      return { kind: "ok", id: request.id };
    }
    case "clearRange": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.clearRange(request.range);
      return { kind: "ok", id: request.id };
    }
    case "fillRange": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.fillRange(request.source, request.target);
      return { kind: "ok", id: request.id };
    }
    case "copyRange": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.copyRange(request.source, request.target);
      return { kind: "ok", id: request.id };
    }
    case "moveRange": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.moveRange(request.source, request.target);
      return { kind: "ok", id: request.id };
    }
    case "pasteRange": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      engine.pasteRange(request.source, request.target);
      return { kind: "ok", id: request.id };
    }
    case "getDependents": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      const dependencies = engine.getDependents(request.sheetName, request.address);
      return {
        kind: "dependencies",
        id: request.id,
        addresses: dependencies.directDependents,
      };
    }
    case "getPrecedents": {
      const { engine } = context.getSessionByAgentSessionId(request.sessionId);
      const dependencies = engine.getDependencies(request.sheetName, request.address);
      return {
        kind: "dependencies",
        id: request.id,
        addresses: dependencies.directPrecedents,
      };
    }
    case "exportSnapshot": {
      const session = context.getSessionByAgentSessionId(request.sessionId);
      return { kind: "snapshot", id: request.id, snapshot: context.getCachedSnapshot(session) };
    }
    case "importSnapshot": {
      const session = context.getSessionByAgentSessionId(request.sessionId);
      context.importSnapshot(session, request.snapshot);
      return { kind: "ok", id: request.id };
    }
    case "subscribeRange": {
      const session = context.getSessionByAgentSessionId(request.sessionId);
      const agentSession = session.agentSessions.get(request.sessionId);
      if (!agentSession) {
        throw new Error(`Unknown agent session: ${request.sessionId}`);
      }
      const existingOwner = context.agentSubscriptionOwners.get(request.subscriptionId);
      if (existingOwner) {
        throw new Error(`Subscription id already in use: ${request.subscriptionId}`);
      }

      const bounds = getRangeBounds(request.range);
      const rangeCellCount = cellCountForRange(request.range);
      const watchedAddresses =
        rangeCellCount <= context.largeRangeSubscriptionThreshold
          ? iterateRange(request.range)
          : null;
      const unsubscribe = watchedAddresses
        ? session.engine.subscribeCells(request.range.sheetName, watchedAddresses, () => {
            context.queueAgentEvent(session.documentId, {
              kind: "rangeChanged",
              subscriptionId: request.subscriptionId,
              range: request.range,
              changedAddresses: [...watchedAddresses],
            });
          })
        : session.engine.subscribe((event: EngineEvent) => {
            const changedInRange = collectChangedAddressesForEvent(
              session.engine,
              request.range,
              bounds,
              event,
            );
            if (changedInRange.length === 0) {
              return;
            }
            context.queueAgentEvent(session.documentId, {
              kind: "rangeChanged",
              subscriptionId: request.subscriptionId,
              range: request.range,
              changedAddresses: changedInRange,
            });
          });

      session.agentSubscriptions.set(request.subscriptionId, {
        subscriptionId: request.subscriptionId,
        sessionId: request.sessionId,
        range: request.range,
        bounds,
        watchedAddresses,
        unsubscribe,
      });
      agentSession.subscriptionIds.add(request.subscriptionId);
      context.agentSubscriptionOwners.set(request.subscriptionId, request.sessionId);
      return {
        kind: "ok",
        id: request.id,
        value: { subscriptionId: request.subscriptionId },
      };
    }
    case "unsubscribe": {
      const session = context.getSessionByAgentSessionId(request.sessionId);
      context.removeAgentSubscription(session, request.sessionId, request.subscriptionId);
      return { kind: "ok", id: request.id };
    }
    case "createPivotTable": {
      const session = context.getSessionByAgentSessionId(request.sessionId);
      session.engine.setPivotTable(request.sheetName, request.address, {
        name: request.name,
        source: request.source,
        groupBy: request.groupBy,
        values: request.values,
      });
      return { kind: "ok", id: request.id };
    }
    default: {
      const exhaustiveRequest: never = request;
      return {
        kind: "error",
        id: "unknown",
        code: "UNSUPPORTED_AGENT_REQUEST",
        message: `Unsupported agent request ${(exhaustiveRequest as { kind: string }).kind}`,
        retryable: false,
      };
    }
  }
}
