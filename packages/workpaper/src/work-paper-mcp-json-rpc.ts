import { WORKPAPER_VERSION } from './work-paper-version.js'
import type { WorkPaperMcpToolServer } from './work-paper-mcp-server.js'

type JsonObject = Record<string, unknown>
type JsonRpcId = string | number | null

interface WorkPaperMcpJsonRpcRequest {
  jsonrpc: '2.0'
  id: JsonRpcId | undefined
  method: string
  params?: JsonObject
}

interface WorkPaperMcpJsonRpcErrorResponse {
  jsonrpc: '2.0'
  id: JsonRpcId
  error: {
    code: number
    message: string
  }
}

interface WorkPaperMcpJsonRpcDispatchOptions {
  server: WorkPaperMcpToolServer
  protocolVersion?: string
  serverName?: string
  serverTitle?: string
  serverVersion?: string
}

type WorkPaperMcpJsonRpcDispatchResult =
  | {
      kind: 'response'
      response: unknown
    }
  | {
      kind: 'notification'
    }

const WORKPAPER_MCP_PROTOCOL_VERSION = '2025-11-25'
const WORKPAPER_MCP_SUPPORTED_PROTOCOL_VERSIONS = ['2025-11-25', '2025-06-18', '2025-03-26'] as const

function dispatchWorkPaperMcpJsonRpc(payload: unknown, options: WorkPaperMcpJsonRpcDispatchOptions): WorkPaperMcpJsonRpcDispatchResult {
  if (Array.isArray(payload)) {
    return {
      kind: 'response',
      response: createWorkPaperMcpJsonRpcError(null, -32600, 'JSON-RPC batch requests are not supported'),
    }
  }

  if (!isWorkPaperMcpJsonRpcRequest(payload)) {
    if (isWorkPaperMcpJsonRpcResponse(payload)) {
      return {
        kind: 'notification',
      }
    }

    const id = isRecord(payload) && isWorkPaperMcpJsonRpcId(payload['id']) ? payload['id'] : null
    return {
      kind: 'response',
      response: createWorkPaperMcpJsonRpcError(id, -32600, 'Invalid JSON-RPC 2.0 request'),
    }
  }

  if (payload.method === 'initialize') {
    return {
      kind: 'response',
      response: {
        jsonrpc: '2.0',
        id: payload.id,
        result: {
          protocolVersion: options.protocolVersion ?? WORKPAPER_MCP_PROTOCOL_VERSION,
          capabilities: options.server.capabilities,
          serverInfo: {
            name: options.serverName ?? 'bilig-headless-workpaper',
            title: options.serverTitle ?? 'Bilig WorkPaper',
            version: options.serverVersion ?? WORKPAPER_VERSION,
          },
        },
      },
    }
  }

  if (payload.method === 'ping') {
    return {
      kind: 'response',
      response: {
        jsonrpc: '2.0',
        id: payload.id,
        result: {},
      },
    }
  }

  if (payload.method === 'notifications/initialized' || payload.id === undefined) {
    return {
      kind: 'notification',
    }
  }

  try {
    return {
      kind: 'response',
      response: options.server.handleJsonRpc(payload),
    }
  } catch (error) {
    return {
      kind: 'response',
      response: createWorkPaperMcpJsonRpcError(payload.id ?? null, errorCode(error), errorMessage(error)),
    }
  }
}

function createWorkPaperMcpJsonRpcError(id: JsonRpcId, code: number, message: string): WorkPaperMcpJsonRpcErrorResponse {
  return {
    jsonrpc: '2.0',
    id,
    error: {
      code,
      message,
    },
  }
}

function isWorkPaperMcpProtocolVersion(value: string): value is (typeof WORKPAPER_MCP_SUPPORTED_PROTOCOL_VERSIONS)[number] {
  return WORKPAPER_MCP_SUPPORTED_PROTOCOL_VERSIONS.some((version) => version === value)
}

function isWorkPaperMcpJsonRpcRequest(value: unknown): value is WorkPaperMcpJsonRpcRequest {
  if (!isRecord(value) || value['jsonrpc'] !== '2.0' || typeof value['method'] !== 'string') {
    return false
  }

  return value['id'] === undefined || isWorkPaperMcpJsonRpcId(value['id'])
}

function isWorkPaperMcpJsonRpcResponse(value: unknown): boolean {
  return (
    isRecord(value) &&
    value['jsonrpc'] === '2.0' &&
    value['method'] === undefined &&
    isWorkPaperMcpJsonRpcId(value['id']) &&
    (Object.hasOwn(value, 'result') || Object.hasOwn(value, 'error'))
  )
}

function isWorkPaperMcpJsonRpcId(value: unknown): value is JsonRpcId {
  return value === null || typeof value === 'string' || typeof value === 'number'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function errorCode(error: unknown): number {
  return errorMessage(error).startsWith('Unsupported MCP method: ') ? -32601 : -32603
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export {
  WORKPAPER_MCP_PROTOCOL_VERSION,
  WORKPAPER_MCP_SUPPORTED_PROTOCOL_VERSIONS,
  createWorkPaperMcpJsonRpcError,
  dispatchWorkPaperMcpJsonRpc,
  isWorkPaperMcpProtocolVersion,
  type JsonRpcId as WorkPaperMcpJsonRpcId,
  type WorkPaperMcpJsonRpcDispatchOptions,
  type WorkPaperMcpJsonRpcDispatchResult,
  type WorkPaperMcpJsonRpcErrorResponse,
  type WorkPaperMcpJsonRpcRequest,
}
