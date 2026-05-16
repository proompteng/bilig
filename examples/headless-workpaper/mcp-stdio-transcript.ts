import { spawn } from 'node:child_process'

type JsonObject = Record<string, unknown>

interface TranscriptOutput {
  transport: 'stdio'
  requestLines: string[]
  responseSummary: {
    protocolVersion: string
    serverName: string
    toolNames: string[]
    write: JsonObject
  }
  verified: {
    initialized: boolean
    listedTools: boolean
    editedCell: 'Inputs!B3'
    formulasPersisted: boolean
    restoredMatchesAfter: boolean
    expectedArrChanged: boolean
  }
}

const requestLines = [
  jsonLine({ jsonrpc: '2.0', id: 1, method: 'initialize' }),
  jsonLine({ jsonrpc: '2.0', method: 'notifications/initialized' }),
  jsonLine({ jsonrpc: '2.0', id: 2, method: 'tools/list' }),
  jsonLine({
    jsonrpc: '2.0',
    id: 3,
    method: 'tools/call',
    params: {
      name: 'set_workpaper_input_cell',
      arguments: {
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      },
    },
  }),
]

const responseLines = await runStdioServer(requestLines)
const transcriptOutput = createTranscriptOutput(requestLines, responseLines)
assertTranscriptOutput(transcriptOutput)

console.log(JSON.stringify(transcriptOutput, null, 2))

function jsonLine(value: unknown): string {
  return JSON.stringify(value)
}

async function runStdioServer(lines: string[]): Promise<string[]> {
  const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm'
  const child = spawn(npmCommand, ['run', '--silent', 'agent:mcp-stdio'], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      NODE_NO_WARNINGS: '1',
    },
    stdio: ['pipe', 'pipe', 'pipe'],
  })

  let stdout = ''
  let stderr = ''
  child.stdout.setEncoding('utf8')
  child.stderr.setEncoding('utf8')
  child.stdout.on('data', (chunk: string) => {
    stdout += chunk
  })
  child.stderr.on('data', (chunk: string) => {
    stderr += chunk
  })

  child.stdin.end(`${lines.join('\n')}\n`)

  const exitCode = await new Promise<number | null>((resolve, reject) => {
    child.on('error', reject)
    child.on('close', resolve)
  })

  if (exitCode !== 0) {
    throw new Error(`MCP stdio transcript failed with exit ${String(exitCode)}: ${stderr}`)
  }
  if (stderr.trim().length > 0) {
    throw new Error(`MCP stdio transcript wrote stderr: ${stderr}`)
  }

  return stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
}

function createTranscriptOutput(requests: string[], responses: string[]): TranscriptOutput {
  const initializeResponse = requireResponse(responses, 1)
  const toolsResponse = requireResponse(responses, 2)
  const writeResponse = requireResponse(responses, 3)
  const initializeResult = readRecord(initializeResponse.result, 'initialize result')
  const serverInfo = readRecord(initializeResult.serverInfo, 'server info')
  const toolsResult = readRecord(toolsResponse.result, 'tools/list result')
  const writeResult = readRecord(writeResponse.result, 'tools/call result')
  const structuredContent = readRecord(writeResult.structuredContent, 'structured content')
  const checks = readRecord(structuredContent.checks, 'structured content checks')

  return {
    transport: 'stdio',
    requestLines: requests,
    responseSummary: {
      protocolVersion: readString(initializeResult.protocolVersion, 'protocolVersion'),
      serverName: readString(serverInfo.name, 'server name'),
      toolNames: readToolNames(toolsResult.tools),
      write: structuredContent,
    },
    verified: {
      initialized: true,
      listedTools: true,
      editedCell: readEditedCell(structuredContent.editedCell),
      formulasPersisted: readBoolean(checks.formulasPersisted, 'formulasPersisted'),
      restoredMatchesAfter: readBoolean(checks.restoredMatchesAfter, 'restoredMatchesAfter'),
      expectedArrChanged: readBoolean(checks.expectedArrChanged, 'expectedArrChanged'),
    },
  }
}

function requireResponse(lines: string[], id: number): JsonObject {
  for (const line of lines) {
    const parsed = readRecord(JSON.parse(line), `response ${id.toString()}`)
    if (parsed.id === id && parsed.jsonrpc === '2.0') {
      return parsed
    }
  }
  throw new Error(`Missing JSON-RPC response id ${id.toString()}: ${lines.join('\n')}`)
}

function readToolNames(value: unknown): string[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected tools/list result to be an array, got ${JSON.stringify(value)}`)
  }

  return value.map((entry, index) => readString(readRecord(entry, `tool ${index.toString()}`).name, `tool ${index.toString()} name`))
}

function assertTranscriptOutput(output: TranscriptOutput): void {
  const write = output.responseSummary.write
  const before = readSummary(write.before, 'before')
  const after = readSummary(write.after, 'after')
  const restored = readSummary(write.restored, 'restored')
  const checks = readRecord(write.checks, 'checks')

  if (
    output.responseSummary.protocolVersion !== '2025-06-18' ||
    output.responseSummary.serverName !== 'bilig-headless-workpaper-example' ||
    JSON.stringify(output.responseSummary.toolNames) !== JSON.stringify(['read_workpaper_summary', 'set_workpaper_input_cell']) ||
    output.verified.editedCell !== 'Inputs!B3' ||
    before.expectedArr !== 60000 ||
    after.expectedArr !== 96000 ||
    restored.expectedArr !== 96000 ||
    checks.previousValue !== 0.25 ||
    checks.newValue !== 0.4 ||
    !output.verified.formulasPersisted ||
    !output.verified.restoredMatchesAfter ||
    !output.verified.expectedArrChanged
  ) {
    throw new Error(`Unexpected MCP stdio transcript: ${JSON.stringify(output)}`)
  }
}

function readSummary(value: unknown, label: string): { expectedArr: number } {
  const record = readRecord(value, label)
  return {
    expectedArr: readNumber(record.expectedArr, `${label} expectedArr`),
  }
}

function readEditedCell(value: unknown): 'Inputs!B3' {
  if (value !== 'Inputs!B3') {
    throw new Error(`Expected editedCell to be Inputs!B3, got ${JSON.stringify(value)}`)
  }
  return value
}

function readRecord(value: unknown, label: string): JsonObject {
  if (!isJsonObject(value)) {
    throw new Error(`Expected ${label} to be an object, got ${JSON.stringify(value)}`)
  }
  return value
}

function isJsonObject(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${label} to be a string, got ${JSON.stringify(value)}`)
  }
  return value
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`Expected ${label} to be a number, got ${JSON.stringify(value)}`)
  }
  return value
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`Expected ${label} to be a boolean, got ${JSON.stringify(value)}`)
  }
  return value
}
