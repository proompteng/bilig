import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http'

import { createN8nForecastProof } from './n8n-forecast-proof.js'
import { createN8nWorkPaperEvaluationProof } from './n8n-workpaper-evaluation-proof.js'

export type N8nForecastServerCliHost = {
  argv: string[]
  env?: Record<string, string | undefined>
  writeStdout?: (text: string) => void
  writeStderr?: (text: string) => void
}

export type N8nForecastServerCliArgs = {
  help: boolean
  host: string
  port: number
}

const DEFAULT_HOST = '127.0.0.1'
const DEFAULT_PORT = 4321
const MAX_REQUEST_BYTES = 1024 * 1024

export function runN8nForecastServerCli(host: N8nForecastServerCliHost): number {
  const writeStdout = host.writeStdout ?? ((text) => process.stdout.write(text))
  const writeStderr = host.writeStderr ?? ((text) => process.stderr.write(text))

  try {
    const options = parseN8nForecastServerCliArgs(host.argv, host.env)
    if (options.help) {
      writeStdout(n8nForecastServerHelpText())
      return 0
    }

    const server = createN8nForecastHttpServer()
    server.on('error', (error) => {
      writeStderr(`Bilig n8n formula server failed: ${error instanceof Error ? error.message : String(error)}\n`)
      process.exitCode = 1
    })
    server.listen(options.port, options.host, () => {
      const baseUrl = `http://${options.host}:${options.port}`
      writeStdout(`Bilig n8n formula server listening on ${baseUrl}\n`)
      writeStdout(`POST ${baseUrl}/api/workpaper/n8n/forecast\n`)
      writeStdout(`POST ${baseUrl}/api/workpaper/n8n/evaluate\n`)
      writeStdout('Use host.docker.internal from n8n Docker when this server runs on the host.\n')
    })
    return 0
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    return 1
  }
}

export function createN8nForecastHttpServer(): Server {
  return createServer((request, response) => {
    void handleN8nForecastHttpRequest(request, response)
  })
}

export function parseN8nForecastServerCliArgs(argv: string[], env: Record<string, string | undefined> = {}): N8nForecastServerCliArgs {
  let host = env['BILIG_N8N_HOST'] ?? DEFAULT_HOST
  let port = readPort(env['PORT'] ?? env['BILIG_N8N_PORT'], DEFAULT_PORT)
  let help = false

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    if (arg === '--help' || arg === '-h') {
      help = true
    } else if (arg === '--host') {
      host = readRequiredValue(argv, index, arg)
      index += 1
    } else if (arg === '--port') {
      port = readPort(readRequiredValue(argv, index, arg), DEFAULT_PORT)
      index += 1
    } else {
      throw new Error(`Unknown bilig-n8n-formula-server argument: ${arg}`)
    }
  }

  return { help, host, port }
}

export function n8nForecastServerHelpText(): string {
  return [
    'Usage: bilig-n8n-formula-server [--host 127.0.0.1] [--port 4321]',
    '',
    'Start a tiny local HTTP server for the importable n8n WorkPaper formula-readback workflow.',
    '',
    'Endpoint:',
    '  POST /api/workpaper/n8n/forecast',
    '  POST /api/workpaper/n8n/evaluate',
    '',
    'Example:',
    '  npm exec --package @bilig/workpaper@latest -- bilig-n8n-formula-server --port 4321',
    '  curl -sS -X POST http://localhost:4321/api/workpaper/n8n/forecast \\',
    "    -H 'content-type: application/json' \\",
    '    --data \'{"sheetName":"Inputs","address":"B3","value":0.4}\'',
    '  curl -sS -X POST http://localhost:4321/api/workpaper/n8n/evaluate \\',
    "    -H 'content-type: application/json' \\",
    '    --data \'{"document":{"format":"bilig.headless.work-paper.document.v1","sheets":[{"name":"Inputs","content":[["Metric","Value"],["Win rate",0.25]]},{"name":"Summary","content":[["Metric","Value"],["Expected customers","=Inputs!B2*20"]]}],"namedExpressions":[]},"edits":[{"cell":"Inputs!B2","value":0.4}],"readCells":["Summary!B2"]}\'',
    '',
    'Environment:',
    '  BILIG_N8N_HOST   Override the listen host. Defaults to 127.0.0.1.',
    '  BILIG_N8N_PORT   Override the listen port. Defaults to 4321.',
    '  PORT             Also accepted for platform hosts.',
    '',
  ].join('\n')
}

async function handleN8nForecastHttpRequest(request: IncomingMessage, response: ServerResponse): Promise<void> {
  if (request.method === 'OPTIONS') {
    writeCorsHeaders(response)
    response.statusCode = 204
    response.end()
    return
  }

  if (request.method === 'GET' && request.url === '/healthz') {
    writeJson(response, 200, { ok: true })
    return
  }

  if (request.method !== 'POST' || !isN8nPostEndpoint(request.url)) {
    writeJson(response, 404, {
      verified: false,
      error: 'Use POST /api/workpaper/n8n/forecast or POST /api/workpaper/n8n/evaluate',
    })
    return
  }

  try {
    const body = await readJsonBody(request)
    if (request.url === '/api/workpaper/n8n/evaluate') {
      writeJson(response, 200, createN8nWorkPaperEvaluationProof(body))
      return
    }
    writeJson(response, 200, createN8nForecastProof(body))
  } catch (error) {
    writeJson(response, 400, {
      verified: false,
      error: error instanceof Error ? error.message : 'Invalid n8n WorkPaper forecast request',
    })
  }
}

function isN8nPostEndpoint(url: string | undefined): boolean {
  return url === '/api/workpaper/n8n/forecast' || url === '/api/workpaper/n8n/evaluate'
}

async function readJsonBody(request: IncomingMessage): Promise<Record<string, unknown>> {
  const chunks: Buffer[] = []
  let totalBytes = 0

  for await (const chunk of request) {
    const buffer = typeof chunk === 'string' ? Buffer.from(chunk) : chunk
    totalBytes += buffer.byteLength
    if (totalBytes > MAX_REQUEST_BYTES) {
      throw new Error('Request body is too large')
    }
    chunks.push(buffer)
  }

  if (chunks.length === 0) {
    return {}
  }

  const text = Buffer.concat(chunks).toString('utf8').trim()
  if (text.length === 0) {
    return {}
  }

  const parsed: unknown = JSON.parse(text)
  if (!isJsonObject(parsed)) {
    throw new Error('Request body must be a JSON object')
  }
  return parsed
}

function isJsonObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function writeJson(response: ServerResponse, statusCode: number, body: unknown): void {
  writeCorsHeaders(response)
  response.statusCode = statusCode
  response.setHeader('cache-control', 'no-store')
  response.setHeader('content-type', 'application/json; charset=utf-8')
  response.end(`${JSON.stringify(body, null, 2)}\n`)
}

function writeCorsHeaders(response: ServerResponse): void {
  response.setHeader('access-control-allow-origin', '*')
  response.setHeader('access-control-allow-methods', 'GET, POST, OPTIONS')
  response.setHeader('access-control-allow-headers', 'content-type, accept')
}

function readRequiredValue(argv: string[], index: number, label: string): string {
  const value = argv[index + 1]
  if (value === undefined || value.startsWith('-')) {
    throw new Error(`${label} requires a value`)
  }
  return value
}

function readPort(value: string | undefined, fallback: number): number {
  if (value === undefined || value.length === 0) {
    return fallback
  }
  const port = Number(value)
  if (!Number.isInteger(port) || port <= 0 || port > 65535) {
    throw new Error(`Port must be an integer from 1 to 65535, received ${value}`)
  }
  return port
}
