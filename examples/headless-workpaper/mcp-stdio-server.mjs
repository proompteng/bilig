import { buildWorkbook, createMcpWorkPaperToolServer } from './mcp-tool-server.mjs'

const server = createMcpWorkPaperToolServer(buildWorkbook())
let inputBuffer = ''

process.stdin.setEncoding('utf8')
process.stdin.on('data', (chunk) => {
  inputBuffer += chunk
  drainInputLines(false)
})
process.stdin.on('end', () => {
  drainInputLines(true)
})

function drainInputLines(flush) {
  let newlineIndex = inputBuffer.indexOf('\n')
  while (newlineIndex !== -1) {
    const line = inputBuffer.slice(0, newlineIndex).trim()
    inputBuffer = inputBuffer.slice(newlineIndex + 1)
    if (line.length > 0) {
      handleLine(line)
    }
    newlineIndex = inputBuffer.indexOf('\n')
  }

  const trailingLine = inputBuffer.trim()
  if (flush && trailingLine.length > 0) {
    inputBuffer = ''
    handleLine(trailingLine)
  }
}

function handleLine(line) {
  let request

  try {
    request = JSON.parse(line)
  } catch (error) {
    writeJsonRpcError(null, -32700, `Parse error: ${errorMessage(error)}`)
    return
  }

  if (request?.jsonrpc !== '2.0' || typeof request.method !== 'string') {
    writeJsonRpcError(request?.id ?? null, -32600, 'Invalid JSON-RPC 2.0 request')
    return
  }

  try {
    const response = dispatchJsonRpc(request)
    if (response !== undefined) {
      writeJson(response)
    }
  } catch (error) {
    writeJsonRpcError(request.id ?? null, -32603, errorMessage(error))
  }
}

function dispatchJsonRpc(request) {
  if (request.method === 'initialize') {
    return {
      jsonrpc: '2.0',
      id: request.id,
      result: {
        protocolVersion: '2025-06-18',
        capabilities: server.capabilities,
        serverInfo: {
          name: 'bilig-headless-workpaper-example',
          version: '0.1.0',
        },
      },
    }
  }

  if (request.method === 'notifications/initialized' || request.id === undefined) {
    return undefined
  }

  return server.handleJsonRpc(request)
}

function writeJsonRpcError(id, code, message) {
  writeJson({
    jsonrpc: '2.0',
    id,
    error: {
      code,
      message,
    },
  })
}

function writeJson(value) {
  process.stdout.write(`${JSON.stringify(value)}\n`)
}

function errorMessage(error) {
  return error instanceof Error ? error.message : String(error)
}
