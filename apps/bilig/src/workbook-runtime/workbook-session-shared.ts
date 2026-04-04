import type { AgentResponse, LoadWorkbookFileRequest } from "@bilig/agent-api";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  type AgentFrameContext,
  createWorkbookLoadedResponse,
  type PreparedWorkbookLoad,
  prepareWorkbookLoad,
  type WorkbookLoadPreparationOptions,
} from "./agent-routing.js";

export interface WorkbookLoadHandlerOptions extends WorkbookLoadPreparationOptions {
  registerPreparedSession?(
    prepared: PreparedWorkbookLoad,
    request: LoadWorkbookFileRequest,
  ): void | Promise<void>;
  publishImportedSnapshot(
    documentId: string,
    snapshot: WorkbookSnapshot,
    prepared: PreparedWorkbookLoad,
  ): void | Promise<void>;
}

export function documentIdFromSessionId(sessionId: string): string {
  return sessionId.split(":")[0] || sessionId;
}

export function createOpenWorkbookSessionResponse(
  requestId: string,
  sessionId: string,
): AgentResponse {
  return {
    kind: "ok",
    id: requestId,
    sessionId,
  };
}

export function createCloseWorkbookSessionResponse(requestId: string): AgentResponse {
  return {
    kind: "ok",
    id: requestId,
  };
}

export async function loadWorkbookIntoRuntime(
  request: LoadWorkbookFileRequest,
  context: AgentFrameContext,
  options: WorkbookLoadHandlerOptions,
): Promise<AgentResponse> {
  const prepared = prepareWorkbookLoad(request, context, options);
  await options.registerPreparedSession?.(prepared, request);
  await options.publishImportedSnapshot(prepared.documentId, prepared.imported.snapshot, prepared);
  return createWorkbookLoadedResponse(request.id, prepared);
}
