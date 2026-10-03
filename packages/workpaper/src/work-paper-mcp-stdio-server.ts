import { buildDemoWorkPaper, createWorkPaperMcpToolServer, type WorkPaperMcpToolServer } from './work-paper-mcp-server.js'
import { WORKPAPER_VERSION } from './work-paper-version.js'
import { createWorkPaperMcpJsonRpcError, dispatchWorkPaperMcpJsonRpc } from './work-paper-mcp-json-rpc.js'
import type { WorkPaper } from './work-paper.js'
import type { Readable, Writable } from 'node:stream'

interface WorkPaperMcpStdioOptions {
  input?: Readable
  output?: Writable
  server?: WorkPaperMcpToolServer
  workbook?: WorkPaper
  serverName?: string
  serverVersion?: string
}

function runDemoWorkPaperMcpStdioServer(options: WorkPaperMcpStdioOptions = {}): void {
  const input = options.input ?? process.stdin
  const output = options.output ?? process.stdout
  const server = options.server ?? createWorkPaperMcpToolServer(options.workbook ?? buildDemoWorkPaper())
  const serverName = options.serverName ?? 'bilig-headless-workpaper'
  const serverVersion = options.serverVersion ?? WORKPAPER_VERSION
  let inputBuffer = ''

  input.setEncoding('utf8')
  input.on('data', (chunk: string) => {
    inputBuffer += chunk
    drainInputLines(false)
  })
  input.on('end', () => {
    drainInputLines(true)
  })

  function drainInputLines(flush: boolean): void {
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

  function handleLine(line: string): void {
    let request: unknown

    try {
      request = JSON.parse(line)
    } catch (error) {
      writeJson(createWorkPaperMcpJsonRpcError(null, -32700, `Parse error: ${errorMessage(error)}`))
      return
    }

    const result = dispatchWorkPaperMcpJsonRpc(request, {
      server,
      protocolVersion: '2025-06-18',
      serverName,
      serverVersion,
    })
    if (result.kind === 'response') {
      writeJson(result.response)
    }
  }

  function writeJson(value: unknown): void {
    output.write(`${JSON.stringify(value)}\n`)
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export { runDemoWorkPaperMcpStdioServer, type WorkPaperMcpStdioOptions }
