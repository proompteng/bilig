import { parseJsonRecord } from './workpaper-external-smoke-parser-helpers.ts'

export function parseNodeMcpStdioErrorOutput(output: string): {
  invalidJson: {
    code: number
    id: null
  }
  invalidRequest: {
    code: number
    id: null
  }
} {
  const responses = output
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0)
    .map((line, index) => parseJsonRecord(line, `node MCP stdio error response ${index + 1}`))

  const invalidJsonResponse = requireJsonRpcErrorResponseAt(responses, 0, null, -32700, 'node MCP stdio invalid JSON response')
  const invalidRequestResponse = requireJsonRpcErrorResponseAt(responses, 1, null, -32700, 'node MCP stdio invalid request response')

  return {
    invalidJson: {
      code: Number(parseRecordValue(invalidJsonResponse.error, 'node MCP stdio invalid JSON error').code),
      id: null,
    },
    invalidRequest: {
      code: Number(parseRecordValue(invalidRequestResponse.error, 'node MCP stdio invalid request error').code),
      id: null,
    },
  }
}

function requireJsonRpcErrorResponseAt(
  responses: Record<string, unknown>[],
  index: number,
  id: number | null,
  code: number,
  context: string,
): Record<string, unknown> {
  const response = responses[index]
  if (response === undefined || response.id !== id || response.jsonrpc !== '2.0') {
    throw new Error(`Missing ${context}: ${JSON.stringify(responses)}`)
  }

  const error = parseRecordValue(response.error, `${context} error`)
  if (error.code !== code || typeof error.message !== 'string' || error.message.length === 0) {
    throw new Error(`Unexpected ${context}: ${JSON.stringify(response)}`)
  }

  return response
}

function parseRecordValue(candidate: unknown, context: string): Record<string, unknown> {
  if (!isRecord(candidate)) {
    throw new Error(`Unexpected ${context}: ${JSON.stringify(candidate)}`)
  }
  return candidate
}

function isRecord(candidate: unknown): candidate is Record<string, unknown> {
  return typeof candidate === 'object' && candidate !== null && !Array.isArray(candidate)
}
