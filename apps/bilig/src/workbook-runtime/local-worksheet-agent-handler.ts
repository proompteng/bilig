import { SpreadsheetEngine } from "@bilig/core";
import type { CellRangeRef, CellValue, EngineEvent, WorkbookSnapshot } from "@bilig/protocol";
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
  decodeColumn,
  encodeColumn,
  iterateRange,
  splitAddress,
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

export function readRange(engine: SpreadsheetEngine, range: CellRangeRef): CellValue[][] {
  const [startColPart] = splitAddress(range.startAddress);
  const [endColPart] = splitAddress(range.endAddress);
  const [, startRowPart] = splitAddress(range.startAddress);
  const [, endRowPart] = splitAddress(range.endAddress);
  const startCol = decodeColumn(startColPart);
  const endCol = decodeColumn(endColPart);
  const startRow = Number.parseInt(startRowPart, 10);
  const endRow = Number.parseInt(endRowPart, 10);
  const width = decodeColumn(endColPart) - decodeColumn(startColPart) + 1;
  const rows: CellValue[][] = [];
  for (let row = startRow; row <= endRow; row += 1) {
    const nextRow: CellValue[] = Array.from<CellValue>({ length: width });
    for (let col = startCol; col <= endCol; col += 1) {
      nextRow[col - startCol] = engine.getCellValue(range.sheetName, `${encodeColumn(col)}${row}`);
    }
    rows.push(nextRow);
  }
  return rows;
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
        values: readRange(engine, request.range),
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

      const rangeCellCount = cellCountForRange(request.range);
      const changedAddresses =
        rangeCellCount <= context.largeRangeSubscriptionThreshold
          ? iterateRange(request.range)
          : null;
      const unsubscribe = changedAddresses
        ? session.engine.subscribeCells(request.range.sheetName, changedAddresses, () => {
            context.queueAgentEvent(session.documentId, {
              kind: "rangeChanged",
              subscriptionId: request.subscriptionId,
              range: request.range,
              changedAddresses: [...changedAddresses],
            });
          })
        : session.engine.subscribe((event: EngineEvent) => {
            const changedInRange = collectChangedAddressesForEvent(
              session.engine,
              request.range,
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
