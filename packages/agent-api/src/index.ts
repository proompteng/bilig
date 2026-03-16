import type { CellValue, LiteralInput, WorkbookSnapshot } from "@bilig/protocol";

export const AGENT_PROTOCOL_VERSION = 1;
export const AGENT_STDIN_MAGIC = 0x41474e54;

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

export interface CellRangeRef {
  sheetName: string;
  startAddress: string;
  endAddress: string;
}

export type AgentRequest =
  | { kind: "openWorkbookSession"; id: string; documentId: string; replicaId: string }
  | { kind: "closeWorkbookSession"; id: string; sessionId: string }
  | { kind: "readRange"; id: string; sessionId: string; range: CellRangeRef }
  | { kind: "writeRange"; id: string; sessionId: string; range: CellRangeRef; values: LiteralInput[][] }
  | { kind: "setRangeFormulas"; id: string; sessionId: string; range: CellRangeRef; formulas: string[][] }
  | { kind: "clearRange"; id: string; sessionId: string; range: CellRangeRef }
  | { kind: "fillRange"; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | { kind: "copyRange"; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | { kind: "pasteRange"; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | { kind: "getDependents"; id: string; sessionId: string; sheetName: string; address: string }
  | { kind: "getPrecedents"; id: string; sessionId: string; sheetName: string; address: string }
  | { kind: "subscribeRange"; id: string; sessionId: string; range: CellRangeRef; subscriptionId: string }
  | { kind: "unsubscribe"; id: string; sessionId: string; subscriptionId: string }
  | { kind: "exportSnapshot"; id: string; sessionId: string }
  | { kind: "importSnapshot"; id: string; sessionId: string; snapshot: WorkbookSnapshot }
  | { kind: "getMetrics"; id: string; sessionId: string };

export type AgentResponse =
  | { kind: "ok"; id: string; sessionId?: string; value?: unknown }
  | { kind: "rangeValues"; id: string; values: CellValue[][] }
  | { kind: "dependencies"; id: string; addresses: string[] }
  | { kind: "snapshot"; id: string; snapshot: WorkbookSnapshot }
  | { kind: "metrics"; id: string; value: unknown }
  | { kind: "error"; id: string; code: string; message: string; retryable: boolean };

export type AgentEvent =
  | { kind: "rangeChanged"; subscriptionId: string; range: CellRangeRef; changedAddresses: string[] }
  | { kind: "syncState"; sessionId: string; state: "local-only" | "syncing" | "live" | "behind" | "reconnecting" };

export type AgentFrame =
  | { kind: "request"; request: AgentRequest }
  | { kind: "response"; response: AgentResponse }
  | { kind: "event"; event: AgentEvent };

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
  return JSON.parse(textDecoder.decode(data.subarray(10))) as AgentFrame;
}

export function encodeStdioMessage(frame: AgentFrame): Uint8Array {
  const body = encodeAgentFrame(frame);
  const output = new Uint8Array(4 + body.byteLength);
  new DataView(output.buffer).setUint32(0, body.byteLength, true);
  output.set(body, 4);
  return output;
}

export function decodeStdioMessages(buffer: Uint8Array): { frames: AgentFrame[]; remainder: Uint8Array } {
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
    remainder: buffer.subarray(offset)
  };
}
