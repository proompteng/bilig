import type {
  CellRangeRef,
  CellNumberFormatInput,
  CellStyleField,
  CellStylePatch,
  CellValue,
  LiteralInput,
  SyncState,
  WorkbookPivotValueSnapshot,
  WorkbookSnapshot,
} from "@bilig/protocol";

export const AGENT_PROTOCOL_VERSION = 1;
export const AGENT_STDIN_MAGIC = 0x41474e54;

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

export const XLSX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const CSV_CONTENT_TYPE = "text/csv";
export const WORKBOOK_IMPORT_CONTENT_TYPES = [XLSX_CONTENT_TYPE, CSV_CONTENT_TYPE] as const;
export type WorkbookImportContentType = (typeof WORKBOOK_IMPORT_CONTENT_TYPES)[number];
export type WorkbookFileOpenMode = "create" | "replace";

export interface LoadWorkbookFileRequest {
  kind: "loadWorkbookFile";
  id: string;
  replicaId: string;
  openMode: WorkbookFileOpenMode;
  documentId?: string;
  fileName: string;
  contentType: WorkbookImportContentType;
  bytesBase64: string;
}

export interface WorkbookLoadedResponse {
  kind: "workbookLoaded";
  id: string;
  documentId: string;
  sessionId: string;
  workbookName: string;
  sheetNames: string[];
  serverUrl: string;
  browserUrl?: string;
  warnings: string[];
}

export type AgentRequest =
  | { kind: "openWorkbookSession"; id: string; documentId: string; replicaId: string }
  | { kind: "closeWorkbookSession"; id: string; sessionId: string }
  | { kind: "readRange"; id: string; sessionId: string; range: CellRangeRef }
  | {
      kind: "writeRange";
      id: string;
      sessionId: string;
      range: CellRangeRef;
      values: LiteralInput[][];
    }
  | {
      kind: "setRangeFormulas";
      id: string;
      sessionId: string;
      range: CellRangeRef;
      formulas: string[][];
    }
  | {
      kind: "setRangeStyle";
      id: string;
      sessionId: string;
      range: CellRangeRef;
      patch: CellStylePatch;
    }
  | {
      kind: "clearRangeStyle";
      id: string;
      sessionId: string;
      range: CellRangeRef;
      fields?: CellStyleField[];
    }
  | {
      kind: "setRangeNumberFormat";
      id: string;
      sessionId: string;
      range: CellRangeRef;
      format: CellNumberFormatInput;
    }
  | {
      kind: "clearRangeNumberFormat";
      id: string;
      sessionId: string;
      range: CellRangeRef;
    }
  | { kind: "clearRange"; id: string; sessionId: string; range: CellRangeRef }
  | { kind: "fillRange"; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | { kind: "copyRange"; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | { kind: "moveRange"; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | {
      kind: "pasteRange";
      id: string;
      sessionId: string;
      source: CellRangeRef;
      target: CellRangeRef;
    }
  | { kind: "getDependents"; id: string; sessionId: string; sheetName: string; address: string }
  | { kind: "getPrecedents"; id: string; sessionId: string; sheetName: string; address: string }
  | {
      kind: "subscribeRange";
      id: string;
      sessionId: string;
      range: CellRangeRef;
      subscriptionId: string;
    }
  | { kind: "unsubscribe"; id: string; sessionId: string; subscriptionId: string }
  | { kind: "exportSnapshot"; id: string; sessionId: string }
  | { kind: "importSnapshot"; id: string; sessionId: string; snapshot: WorkbookSnapshot }
  | { kind: "getMetrics"; id: string; sessionId: string }
  | {
      kind: "createPivotTable";
      id: string;
      sessionId: string;
      name: string;
      sheetName: string;
      address: string;
      source: CellRangeRef;
      groupBy: string[];
      values: WorkbookPivotValueSnapshot[];
    }
  | LoadWorkbookFileRequest;

export type AgentResponse =
  | { kind: "ok"; id: string; sessionId?: string; value?: unknown }
  | { kind: "rangeValues"; id: string; values: CellValue[][] }
  | { kind: "dependencies"; id: string; addresses: string[] }
  | { kind: "snapshot"; id: string; snapshot: WorkbookSnapshot }
  | { kind: "metrics"; id: string; value: unknown }
  | WorkbookLoadedResponse
  | { kind: "error"; id: string; code: string; message: string; retryable: boolean };

export type AgentEvent =
  | {
      kind: "rangeChanged";
      subscriptionId: string;
      range: CellRangeRef;
      changedAddresses: string[];
    }
  | { kind: "syncState"; sessionId: string; state: SyncState };

export type AgentFrame =
  | { kind: "request"; request: AgentRequest }
  | { kind: "response"; response: AgentResponse }
  | { kind: "event"; event: AgentEvent };

export {
  WORKBOOK_AGENT_TOOL_NAMES,
  isWorkbookAgentToolName,
  normalizeWorkbookAgentToolName,
  type WorkbookAgentToolName,
} from "./workbook-agent-tool-names.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookImportContentType(value: unknown): value is WorkbookImportContentType {
  return value === XLSX_CONTENT_TYPE || value === CSV_CONTENT_TYPE;
}

function isLoadWorkbookFileRequest(value: unknown): value is LoadWorkbookFileRequest {
  return (
    isRecord(value) &&
    value["kind"] === "loadWorkbookFile" &&
    typeof value["id"] === "string" &&
    typeof value["replicaId"] === "string" &&
    (value["openMode"] === "create" || value["openMode"] === "replace") &&
    (value["documentId"] === undefined || typeof value["documentId"] === "string") &&
    typeof value["fileName"] === "string" &&
    isWorkbookImportContentType(value["contentType"]) &&
    typeof value["bytesBase64"] === "string"
  );
}

function isWorkbookLoadedResponse(value: unknown): value is WorkbookLoadedResponse {
  return (
    isRecord(value) &&
    value["kind"] === "workbookLoaded" &&
    typeof value["id"] === "string" &&
    typeof value["documentId"] === "string" &&
    typeof value["sessionId"] === "string" &&
    typeof value["workbookName"] === "string" &&
    Array.isArray(value["sheetNames"]) &&
    value["sheetNames"].every((entry) => typeof entry === "string") &&
    typeof value["serverUrl"] === "string" &&
    (value["browserUrl"] === undefined || typeof value["browserUrl"] === "string") &&
    Array.isArray(value["warnings"]) &&
    value["warnings"].every((entry) => typeof entry === "string")
  );
}

function isAgentRequest(value: unknown): value is AgentRequest {
  if (!isRecord(value) || typeof value["kind"] !== "string" || typeof value["id"] !== "string") {
    return false;
  }
  if (value["kind"] === "loadWorkbookFile") {
    return isLoadWorkbookFileRequest(value);
  }
  return true;
}

function isAgentResponse(value: unknown): value is AgentResponse {
  if (!isRecord(value) || typeof value["kind"] !== "string" || typeof value["id"] !== "string") {
    return false;
  }
  if (value["kind"] === "workbookLoaded") {
    return isWorkbookLoadedResponse(value);
  }
  return true;
}

function isAgentFrame(value: unknown): value is AgentFrame {
  const kind = isRecord(value) ? value["kind"] : undefined;
  if (typeof kind !== "string") {
    return false;
  }
  switch (kind) {
    case "request":
      return isRecord(value) && isAgentRequest(value["request"]);
    case "response":
      return isRecord(value) && isAgentResponse(value["response"]);
    case "event":
      return isRecord(value) && "event" in value;
    default:
      return false;
  }
}

export function encodeAgentFrame(frame: AgentFrame): Uint8Array {
  const payload = textEncoder.encode(JSON.stringify(frame));
  const output = new Uint8Array(10 + payload.byteLength);
  const view = new DataView(output.buffer);
  view.setUint32(0, AGENT_STDIN_MAGIC, true);
  view.setUint16(4, AGENT_PROTOCOL_VERSION, true);
  view.setUint32(6, payload.byteLength, true);
  output.set(payload, 10);
  return output;
}

export function decodeAgentFrame(bytes: Uint8Array | ArrayBuffer): AgentFrame {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (data.byteLength < 10) {
    throw new Error("Agent frame too short");
  }

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (view.getUint32(0, true) !== AGENT_STDIN_MAGIC) {
    throw new Error("Agent frame magic mismatch");
  }
  if (view.getUint16(4, true) !== AGENT_PROTOCOL_VERSION) {
    throw new Error("Unsupported agent protocol version");
  }
  const payloadLength = view.getUint32(6, true);
  if (payloadLength !== data.byteLength - 10) {
    throw new Error("Agent frame length mismatch");
  }
  const parsed: unknown = JSON.parse(textDecoder.decode(data.subarray(10)));
  if (!isAgentFrame(parsed)) {
    throw new Error("Invalid agent frame payload");
  }
  return parsed;
}

export function encodeStdioMessage(frame: AgentFrame): Uint8Array {
  const body = encodeAgentFrame(frame);
  const output = new Uint8Array(4 + body.byteLength);
  new DataView(output.buffer).setUint32(0, body.byteLength, true);
  output.set(body, 4);
  return output;
}

export function decodeStdioMessages(buffer: Uint8Array): {
  frames: AgentFrame[];
  remainder: Uint8Array;
} {
  const frames: AgentFrame[] = [];
  let offset = 0;

  while (offset + 4 <= buffer.byteLength) {
    const length = new DataView(buffer.buffer, buffer.byteOffset + offset, 4).getUint32(0, true);
    if (offset + 4 + length > buffer.byteLength) {
      break;
    }
    frames.push(decodeAgentFrame(buffer.subarray(offset + 4, offset + 4 + length)));
    offset += 4 + length;
  }

  return {
    frames,
    remainder: buffer.subarray(offset),
  };
}

export * from "./workbook-agent-bundles.js";
export * from "./codex-app-server-protocol.js";
export * from "./workbook-agent-execution-policy.js";
export * from "./workbook-agent-preview.js";
export * from "./workbook-agent-review-items.js";
export * from "./workbook-agent-skills.js";
export * from "./workbook-agent-annotation-commands.js";
export * from "./workbook-agent-conditional-format-commands.js";
export * from "./workbook-agent-object-commands.js";
export * from "./workbook-agent-structural-commands.js";
export * from "./workbook-agent-validation-commands.js";
