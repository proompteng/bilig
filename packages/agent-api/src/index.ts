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
} from '@bilig/protocol'
import { MAX_AGENT_FRAME_BYTES, MAX_AGENT_FRAME_PAYLOAD_BYTES } from './agent-frame-limits.js'
import { isAgentFramePayload } from './agent-frame-schema.js'

export const AGENT_PROTOCOL_VERSION = 1
export const AGENT_STDIN_MAGIC = 0x41474e54

export { MAX_AGENT_WORKBOOK_IMPORT_BYTES } from './agent-frame-limits.js'

const textEncoder = new TextEncoder()
const textDecoder = new TextDecoder()

export {
  BILIG_CONTENT_TYPE,
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  WORKBOOK_IMPORT_CONTENT_TYPES,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  normalizeWorkbookImportContentType,
  type WorkbookImportContentType,
} from './workbook-import-content-types.js'
export type WorkbookFileOpenMode = 'create' | 'replace'

export interface LoadWorkbookFileRequest {
  kind: 'loadWorkbookFile'
  id: string
  replicaId: string
  openMode: WorkbookFileOpenMode
  documentId?: string
  fileName: string
  contentType: string
  bytesBase64: string
}

export interface WorkbookLoadedResponse {
  kind: 'workbookLoaded'
  id: string
  documentId: string
  sessionId: string
  workbookName: string
  sheetNames: string[]
  serverUrl: string
  browserUrl?: string
  warnings: string[]
}

export type AgentRequest =
  | { kind: 'openWorkbookSession'; id: string; documentId: string; replicaId: string }
  | { kind: 'closeWorkbookSession'; id: string; sessionId: string }
  | { kind: 'readRange'; id: string; sessionId: string; range: CellRangeRef }
  | {
      kind: 'writeRange'
      id: string
      sessionId: string
      range: CellRangeRef
      values: LiteralInput[][]
    }
  | {
      kind: 'setRangeFormulas'
      id: string
      sessionId: string
      range: CellRangeRef
      formulas: string[][]
    }
  | {
      kind: 'setRangeStyle'
      id: string
      sessionId: string
      range: CellRangeRef
      patch: CellStylePatch
    }
  | {
      kind: 'clearRangeStyle'
      id: string
      sessionId: string
      range: CellRangeRef
      fields?: CellStyleField[]
    }
  | {
      kind: 'setRangeNumberFormat'
      id: string
      sessionId: string
      range: CellRangeRef
      format: CellNumberFormatInput
    }
  | {
      kind: 'clearRangeNumberFormat'
      id: string
      sessionId: string
      range: CellRangeRef
    }
  | { kind: 'clearRange'; id: string; sessionId: string; range: CellRangeRef }
  | { kind: 'fillRange'; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | { kind: 'copyRange'; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | { kind: 'moveRange'; id: string; sessionId: string; source: CellRangeRef; target: CellRangeRef }
  | {
      kind: 'pasteRange'
      id: string
      sessionId: string
      source: CellRangeRef
      target: CellRangeRef
    }
  | { kind: 'getDependents'; id: string; sessionId: string; sheetName: string; address: string }
  | { kind: 'getPrecedents'; id: string; sessionId: string; sheetName: string; address: string }
  | {
      kind: 'subscribeRange'
      id: string
      sessionId: string
      range: CellRangeRef
      subscriptionId: string
    }
  | { kind: 'unsubscribe'; id: string; sessionId: string; subscriptionId: string }
  | { kind: 'exportSnapshot'; id: string; sessionId: string }
  | { kind: 'importSnapshot'; id: string; sessionId: string; snapshot: WorkbookSnapshot }
  | { kind: 'getMetrics'; id: string; sessionId: string }
  | {
      kind: 'createPivotTable'
      id: string
      sessionId: string
      name: string
      sheetName: string
      address: string
      source: CellRangeRef
      groupBy: string[]
      values: WorkbookPivotValueSnapshot[]
    }
  | LoadWorkbookFileRequest

export type AgentResponse =
  | { kind: 'ok'; id: string; sessionId?: string; value?: unknown }
  | { kind: 'rangeValues'; id: string; values: CellValue[][] }
  | { kind: 'dependencies'; id: string; addresses: string[] }
  | { kind: 'snapshot'; id: string; snapshot: WorkbookSnapshot }
  | { kind: 'metrics'; id: string; value: unknown }
  | WorkbookLoadedResponse
  | { kind: 'error'; id: string; code: string; message: string; retryable: boolean }

export type AgentEvent =
  | {
      kind: 'rangeChanged'
      subscriptionId: string
      range: CellRangeRef
      changedAddresses: string[]
    }
  | { kind: 'syncState'; sessionId: string; state: SyncState }

export type AgentFrame =
  | { kind: 'request'; request: AgentRequest }
  | { kind: 'response'; response: AgentResponse }
  | { kind: 'event'; event: AgentEvent }

export {
  WORKBOOK_AGENT_TOOL_NAMES,
  isWorkbookAgentToolName,
  normalizeWorkbookAgentToolName,
  type WorkbookAgentToolName,
} from './workbook-agent-tool-names.js'

export function encodeAgentFrame(frame: AgentFrame): Uint8Array {
  const payload = textEncoder.encode(JSON.stringify(frame))
  if (payload.byteLength > MAX_AGENT_FRAME_PAYLOAD_BYTES) {
    throw new Error('Agent frame exceeds the protocol size limit')
  }
  const output = new Uint8Array(10 + payload.byteLength)
  const view = new DataView(output.buffer)
  view.setUint32(0, AGENT_STDIN_MAGIC, true)
  view.setUint16(4, AGENT_PROTOCOL_VERSION, true)
  view.setUint32(6, payload.byteLength, true)
  output.set(payload, 10)
  return output
}

export function decodeAgentFrame(bytes: Uint8Array | ArrayBuffer): AgentFrame {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  if (data.byteLength > MAX_AGENT_FRAME_BYTES) {
    throw new Error('Agent frame exceeds the protocol size limit')
  }
  if (data.byteLength < 10) {
    throw new Error('Agent frame too short')
  }

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  if (view.getUint32(0, true) !== AGENT_STDIN_MAGIC) {
    throw new Error('Agent frame magic mismatch')
  }
  if (view.getUint16(4, true) !== AGENT_PROTOCOL_VERSION) {
    throw new Error('Unsupported agent protocol version')
  }
  const payloadLength = view.getUint32(6, true)
  if (payloadLength !== data.byteLength - 10) {
    throw new Error('Agent frame length mismatch')
  }
  const parsed: unknown = JSON.parse(textDecoder.decode(data.subarray(10)))
  if (!isAgentFramePayload(parsed)) {
    throw new Error('Invalid agent frame payload')
  }
  return parsed
}

export function encodeStdioMessage(frame: AgentFrame): Uint8Array {
  const body = encodeAgentFrame(frame)
  const output = new Uint8Array(4 + body.byteLength)
  new DataView(output.buffer).setUint32(0, body.byteLength, true)
  output.set(body, 4)
  return output
}

export function decodeStdioMessages(buffer: Uint8Array): {
  frames: AgentFrame[]
  remainder: Uint8Array
} {
  const frames: AgentFrame[] = []
  let offset = 0

  while (offset + 4 <= buffer.byteLength) {
    const length = new DataView(buffer.buffer, buffer.byteOffset + offset, 4).getUint32(0, true)
    if (length > MAX_AGENT_FRAME_BYTES) {
      throw new Error('Agent frame exceeds the protocol size limit')
    }
    if (offset + 4 + length > buffer.byteLength) {
      break
    }
    frames.push(decodeAgentFrame(buffer.subarray(offset + 4, offset + 4 + length)))
    offset += 4 + length
  }

  return {
    frames,
    remainder: buffer.subarray(offset),
  }
}

export * from './workbook-agent-bundles.js'
export * from './workbook-agent-command-handoff.js'
export * from './codex-app-server-protocol.js'
export * from './workbook-agent-execution-policy.js'
export * from './workbook-agent-preview.js'
export * from './workbook-agent-review-items.js'
export * from './workbook-agent-skills.js'
export * from './workbook-agent-annotation-commands.js'
export * from './workbook-agent-conditional-format-commands.js'
export * from './workbook-agent-media-commands.js'
export * from './workbook-agent-object-commands.js'
export * from './workbook-agent-protection-commands.js'
export * from './workbook-agent-structural-commands.js'
export * from './workbook-agent-validation-commands.js'
