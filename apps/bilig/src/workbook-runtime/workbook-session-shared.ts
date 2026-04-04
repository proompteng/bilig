import type {
  AgentFrame,
  AgentRequest,
  AgentResponse,
  LoadWorkbookFileRequest,
} from "@bilig/agent-api";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  type AgentFrameContext,
  routeAgentFrame,
  createWorkbookLoadedResponse,
  type PreparedWorkbookLoad,
  prepareWorkbookLoad,
  worksheetHostUnavailableResponse,
  type WorkbookLoadPreparationOptions,
  type WorksheetAgentRequest,
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

export interface SharedWorkbookLoadOptions extends WorkbookLoadPreparationOptions {}

export interface WorkbookAgentFrameHandlerOptions {
  invalidFrameMessage: string;
  errorCode: string;
  loadWorkbookFile: (
    request: LoadWorkbookFileRequest,
    context: AgentFrameContext,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>;
  openWorkbookSession: (
    request: Extract<AgentRequest, { kind: "openWorkbookSession" }>,
  ) => string | AgentResponse | AgentFrame | Promise<string | AgentResponse | AgentFrame>;
  closeWorkbookSession: (
    request: Extract<AgentRequest, { kind: "closeWorkbookSession" }>,
  ) => void | AgentResponse | AgentFrame | Promise<void | AgentResponse | AgentFrame>;
  getMetrics: (
    request: Extract<AgentRequest, { kind: "getMetrics" }>,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>;
  handleWorksheetRequest?: (
    frame: Extract<AgentFrame, { kind: "request" }>,
    request: WorksheetAgentRequest,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>;
}

export function createWorkbookLoadOptions(
  baseOptions: SharedWorkbookLoadOptions,
  handlers: Pick<WorkbookLoadHandlerOptions, "registerPreparedSession" | "publishImportedSnapshot">,
): WorkbookLoadHandlerOptions {
  return {
    ...(baseOptions.maxImportBytes !== undefined
      ? { maxImportBytes: baseOptions.maxImportBytes }
      : {}),
    ...(baseOptions.publicServerUrl ? { publicServerUrl: baseOptions.publicServerUrl } : {}),
    ...(baseOptions.browserAppBaseUrl ? { browserAppBaseUrl: baseOptions.browserAppBaseUrl } : {}),
    ...handlers,
  };
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

export async function handleWorkbookAgentFrame(
  frame: AgentFrame,
  context: AgentFrameContext,
  options: WorkbookAgentFrameHandlerOptions,
): Promise<AgentFrame> {
  const handleWorksheetRequest = options.handleWorksheetRequest;
  return routeAgentFrame(frame, context, {
    invalidFrameMessage: options.invalidFrameMessage,
    errorCode: options.errorCode,
    loadWorkbookFile: options.loadWorkbookFile,
    openWorkbookSession: async (request) => {
      const result = await options.openWorkbookSession(request);
      return typeof result === "string"
        ? createOpenWorkbookSessionResponse(request.id, result)
        : result;
    },
    closeWorkbookSession: async (request) => {
      const result = await options.closeWorkbookSession(request);
      return result === undefined ? createCloseWorkbookSessionResponse(request.id) : result;
    },
    getMetrics: options.getMetrics,
    handleWorksheetRequest: handleWorksheetRequest
      ? (requestFrame, request) => handleWorksheetRequest(requestFrame, request)
      : (_requestFrame, request) => worksheetHostUnavailableResponse(request),
  });
}
