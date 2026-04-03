import {
  XLSX_CONTENT_TYPE,
  type AgentFrame,
  type AgentRequest,
  type AgentResponse,
  type LoadWorkbookFileRequest,
  type WorkbookLoadedResponse,
} from "@bilig/agent-api";
import { importXlsx, type ImportedWorkbook } from "@bilig/excel-import";
import {
  buildBrowserUrl,
  createImportedDocumentId,
  decodeWorkbookBase64,
  normalizeBaseUrl,
} from "./session-shared.js";

export interface AgentFrameContext {
  serverUrl?: string;
  browserAppBaseUrl?: string;
}

export interface WorkbookLoadPreparationOptions {
  maxImportBytes?: number;
  publicServerUrl?: string;
  browserAppBaseUrl?: string;
  defaultServerUrl?: string;
}

export interface PreparedWorkbookLoad {
  imported: ImportedWorkbook;
  documentId: string;
  sessionId: string;
  serverUrl: string;
  browserUrl?: string;
}

export type WorksheetAgentRequest = Exclude<
  AgentRequest,
  | LoadWorkbookFileRequest
  | { kind: "openWorkbookSession" }
  | { kind: "closeWorkbookSession" }
  | { kind: "getMetrics" }
>;

interface AgentFrameRouterOptions {
  invalidFrameMessage: string;
  errorCode: string;
  loadWorkbookFile: (
    request: LoadWorkbookFileRequest,
    context: AgentFrameContext,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>;
  openWorkbookSession: (
    request: Extract<AgentRequest, { kind: "openWorkbookSession" }>,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>;
  closeWorkbookSession: (
    request: Extract<AgentRequest, { kind: "closeWorkbookSession" }>,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>;
  getMetrics: (
    request: Extract<AgentRequest, { kind: "getMetrics" }>,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>;
  handleWorksheetRequest?: (
    frame: Extract<AgentFrame, { kind: "request" }>,
    request: WorksheetAgentRequest,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>;
}

export function normalizeSessionId(documentId: string, replicaId: string): string {
  return `${documentId}:${replicaId}`;
}

export function prepareWorkbookLoad(
  request: LoadWorkbookFileRequest,
  context: AgentFrameContext,
  options: WorkbookLoadPreparationOptions = {},
): PreparedWorkbookLoad {
  if (request.contentType !== XLSX_CONTENT_TYPE) {
    throw new Error("Unsupported workbook upload content type");
  }
  if (request.openMode === "replace" && !request.documentId) {
    throw new Error("Workbook replace uploads require documentId");
  }

  const bytes = decodeWorkbookBase64(request.bytesBase64);
  const maxImportBytes = options.maxImportBytes ?? 10 * 1024 * 1024;
  if (bytes.byteLength > maxImportBytes) {
    throw new Error(`Workbook upload exceeds ${maxImportBytes} bytes`);
  }

  const imported = importXlsx(bytes, request.fileName);
  const documentId = request.documentId ?? createImportedDocumentId();
  const sessionId = normalizeSessionId(documentId, request.replicaId);
  const serverUrl = normalizeBaseUrl(
    context.serverUrl ??
      options.publicServerUrl ??
      options.defaultServerUrl ??
      "http://127.0.0.1:4321",
  );
  const browserUrl = buildBrowserUrl(
    context.browserAppBaseUrl ?? options.browserAppBaseUrl,
    serverUrl,
    documentId,
  );

  return {
    imported,
    documentId,
    sessionId,
    serverUrl,
    ...(browserUrl ? { browserUrl } : {}),
  };
}

export function createWorkbookLoadedResponse(
  requestId: string,
  prepared: PreparedWorkbookLoad,
): WorkbookLoadedResponse {
  return {
    kind: "workbookLoaded",
    id: requestId,
    documentId: prepared.documentId,
    sessionId: prepared.sessionId,
    workbookName: prepared.imported.workbookName,
    sheetNames: prepared.imported.sheetNames,
    serverUrl: prepared.serverUrl,
    ...(prepared.browserUrl ? { browserUrl: prepared.browserUrl } : {}),
    warnings: prepared.imported.warnings,
  };
}

export async function routeAgentFrame(
  frame: AgentFrame,
  context: AgentFrameContext,
  options: AgentFrameRouterOptions,
): Promise<AgentFrame> {
  if (frame.kind !== "request") {
    return responseFrame({
      kind: "error",
      id: "unknown",
      code: "INVALID_AGENT_FRAME",
      message: options.invalidFrameMessage,
      retryable: false,
    });
  }

  const request = frame.request;
  try {
    if (request.kind === "loadWorkbookFile") {
      return normalizeAgentHandlerResult(await options.loadWorkbookFile(request, context));
    }
    if (request.kind === "openWorkbookSession") {
      return normalizeAgentHandlerResult(await options.openWorkbookSession(request));
    }
    if (request.kind === "closeWorkbookSession") {
      return normalizeAgentHandlerResult(await options.closeWorkbookSession(request));
    }
    if (request.kind === "getMetrics") {
      return normalizeAgentHandlerResult(await options.getMetrics(request));
    }

    const worksheetRequest = request as WorksheetAgentRequest;
    if (!options.handleWorksheetRequest) {
      return responseFrame(worksheetHostUnavailableResponse(worksheetRequest));
    }
    return normalizeAgentHandlerResult(
      await options.handleWorksheetRequest(frame, worksheetRequest),
    );
  } catch (error) {
    return responseFrame({
      kind: "error",
      id: request.id,
      code: options.errorCode,
      message: error instanceof Error ? error.message : String(error),
      retryable: false,
    });
  }
}

export function worksheetHostUnavailableResponse(request: WorksheetAgentRequest): AgentResponse {
  return {
    kind: "error",
    id: request.id,
    code: "WORKSHEET_HOST_UNAVAILABLE",
    message: `${request.kind} requires a live worksheet executor, but none is configured for this server`,
    retryable: true,
  };
}

function responseFrame(response: AgentResponse): AgentFrame {
  return { kind: "response", response };
}

function normalizeAgentHandlerResult(result: AgentResponse | AgentFrame): AgentFrame {
  if (result.kind === "request" || result.kind === "response" || result.kind === "event") {
    return result;
  }
  return responseFrame(result);
}
