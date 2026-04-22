import readline from 'node:readline'
import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import type {
  CodexDynamicToolCallRequest,
  CodexDynamicToolCallResult,
  CodexDynamicToolSpec,
  CodexInitializeCapabilities,
  CodexInitializeResponse,
  CodexJsonRpcError,
  CodexJsonRpcResponse,
  CodexRequestId,
  CodexServerNotification,
  CodexThread,
  CodexThreadStartResponse,
  CodexTurn,
  CodexTurnStartResponse,
  CodexUserInput,
} from '@bilig/agent-api'

export interface CodexAppServerClientOptions {
  command?: string
  args?: string[]
  cwd?: string
  env?: NodeJS.ProcessEnv
  onLog?: (message: string) => void
  handleDynamicToolCall: (request: CodexDynamicToolCallRequest) => Promise<CodexDynamicToolCallResult>
}

export type CodexAppServerJsonValue =
  | boolean
  | number
  | string
  | null
  | CodexAppServerJsonValue[]
  | { [key: string]: CodexAppServerJsonValue }

export type CodexAppServerApprovalPolicy = 'never'
export type CodexAppServerSandboxMode = 'read-only' | 'workspace-write' | 'danger-full-access'
export type CodexAppServerThreadConfig = { readonly [key: string]: CodexAppServerJsonValue }

export interface CodexAppServerTransport {
  ensureReady(): Promise<CodexInitializeResponse>
  subscribe(listener: (notification: CodexServerNotification) => void): () => void
  threadStart(input: {
    model: string
    approvalPolicy: CodexAppServerApprovalPolicy
    sandbox: CodexAppServerSandboxMode
    config?: CodexAppServerThreadConfig
    baseInstructions: string
    developerInstructions: string
    dynamicTools: readonly CodexDynamicToolSpec[]
  }): Promise<CodexThread>
  threadResume(input: { threadId: string; baseInstructions: string; developerInstructions: string }): Promise<CodexThread>
  turnStart(input: { threadId: string; prompt: string }): Promise<CodexTurn>
  turnInterrupt(threadId: string): Promise<void>
  close(): Promise<void>
}

interface PendingResponse {
  readonly resolve: (value: unknown) => void
  readonly reject: (error: Error) => void
}

type ParsedJsonValue = CodexAppServerJsonValue

type ParsedThreadItem =
  | {
      type: 'userMessage'
      id: string
      content: CodexUserInput[]
    }
  | {
      type: 'agentMessage'
      id: string
      text: string
      phase: string | null
      memoryCitation: unknown
    }
  | {
      type: 'plan'
      id: string
      text: string
    }
  | {
      type: 'dynamicToolCall'
      id: string
      tool: string
      arguments: ParsedJsonValue
      status: 'inProgress' | 'completed' | 'failed'
      contentItems: Array<
        | {
            type: 'inputText'
            text: string
          }
        | {
            type: 'inputImage'
            imageUrl: string
          }
      > | null
      success: boolean | null
      durationMs: number | null
    }
  | {
      type: string
      id: string
      [key: string]: ParsedJsonValue | undefined
    }

type ParsedServerRequest =
  | {
      method: 'item/tool/call'
      id: CodexRequestId
      params: CodexDynamicToolCallRequest
    }
  | {
      method: string
      id: CodexRequestId
      params?: ParsedJsonValue
    }

function isDynamicToolCallServerRequest(
  request: ParsedServerRequest,
): request is Extract<ParsedServerRequest, { method: 'item/tool/call' }> {
  return (
    request.method === 'item/tool/call' &&
    typeof request.params === 'object' &&
    request.params !== null &&
    'threadId' in request.params &&
    typeof request.params.threadId === 'string'
  )
}

const JSON_RPC_INTERNAL_ERROR = -32603
const JSON_RPC_METHOD_NOT_FOUND = -32601
const CODEX_INITIALIZE_CAPABILITIES: CodexInitializeCapabilities = {
  experimentalApi: true,
}
const TELEMETRY_ENV_PREFIXES = ['OTEL_'] as const

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isString(value: unknown): value is string {
  return typeof value === 'string'
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function isJsonValue(value: unknown): value is ParsedJsonValue {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return true
  }
  if (Array.isArray(value)) {
    return value.every((entry) => isJsonValue(entry))
  }
  if (!isRecord(value)) {
    return false
  }
  return Object.values(value).every((entry) => isJsonValue(entry))
}

function asError(value: unknown, fallback: string): Error {
  return value instanceof Error ? value : new Error(fallback)
}

function stripTelemetryEnv(inputEnv: NodeJS.ProcessEnv | undefined): NodeJS.ProcessEnv {
  const nextEnv: NodeJS.ProcessEnv = {
    ...(inputEnv ?? process.env),
    OTEL_SDK_DISABLED: 'true',
  }
  for (const key of Object.keys(nextEnv)) {
    if (key === 'OTEL_SDK_DISABLED') {
      continue
    }
    if (TELEMETRY_ENV_PREFIXES.some((prefix) => key.startsWith(prefix))) {
      delete nextEnv[key]
    }
  }
  return nextEnv
}

function parseInitializeResponse(value: unknown): CodexInitializeResponse | null {
  if (!isRecord(value)) {
    return null
  }
  const { codexHome, platformFamily, platformOs, userAgent } = value
  if (!isString(userAgent)) {
    return null
  }
  return {
    userAgent,
    ...(isString(codexHome) ? { codexHome } : {}),
    ...(isString(platformFamily) ? { platformFamily } : {}),
    ...(isString(platformOs) ? { platformOs } : {}),
  }
}

function parseTurnError(value: unknown): CodexTurn['error'] | null {
  if (value === null) {
    return null
  }
  if (!isRecord(value)) {
    return null
  }
  if (value['message'] !== undefined && !isString(value['message'])) {
    return null
  }
  return isString(value['message']) ? { message: value['message'] } : {}
}

function parseToolContentItem(value: unknown): CodexDynamicToolCallResult['contentItems'][number] | null {
  if (!isRecord(value) || !isString(value['type'])) {
    return null
  }
  if (value['type'] === 'inputText' && isString(value['text'])) {
    return {
      type: 'inputText',
      text: value['text'],
    }
  }
  if (value['type'] === 'inputImage' && isString(value['imageUrl'])) {
    return {
      type: 'inputImage',
      imageUrl: value['imageUrl'],
    }
  }
  return null
}

function parseThreadItem(value: unknown): ParsedThreadItem | null {
  if (!isRecord(value) || !isString(value['type']) || !isString(value['id'])) {
    return null
  }
  const type = value['type']
  const id = value['id']
  switch (type) {
    case 'userMessage': {
      if (!Array.isArray(value['content'])) {
        return null
      }
      const content: CodexUserInput[] = []
      for (const entry of value['content']) {
        if (!isRecord(entry) || entry['type'] !== 'text' || !isString(entry['text'])) {
          return null
        }
        content.push({
          type: 'text',
          text: entry['text'],
        })
      }
      return {
        type,
        id,
        content,
      }
    }
    case 'agentMessage': {
      if (!isString(value['text'])) {
        return null
      }
      const phase = value['phase']
      if (phase !== null && !isString(phase)) {
        return null
      }
      return {
        type,
        id,
        text: value['text'],
        phase,
        memoryCitation: value['memoryCitation'],
      }
    }
    case 'plan': {
      if (!isString(value['text'])) {
        return null
      }
      return {
        type,
        id,
        text: value['text'],
      }
    }
    case 'dynamicToolCall': {
      if (
        !isString(value['tool']) ||
        !isJsonValue(value['arguments']) ||
        (value['status'] !== 'inProgress' && value['status'] !== 'completed' && value['status'] !== 'failed') ||
        (value['success'] !== null && typeof value['success'] !== 'boolean') ||
        (value['durationMs'] !== null && !isFiniteNumber(value['durationMs']))
      ) {
        return null
      }
      const contentItemsValue = value['contentItems']
      if (contentItemsValue !== null && !Array.isArray(contentItemsValue)) {
        return null
      }
      const contentItems = contentItemsValue === null ? null : contentItemsValue.map((entry) => parseToolContentItem(entry))
      if (contentItems && contentItems.some((entry) => entry === null)) {
        return null
      }
      return {
        type,
        id,
        tool: value['tool'],
        arguments: value['arguments'],
        status: value['status'],
        contentItems,
        success: value['success'],
        durationMs: value['durationMs'],
      }
    }
    default: {
      const additionalEntries: Record<string, ParsedJsonValue | undefined> = {}
      for (const [key, entry] of Object.entries(value)) {
        if (key === 'type' || key === 'id') {
          continue
        }
        if (entry !== undefined && !isJsonValue(entry)) {
          return null
        }
        additionalEntries[key] = entry
      }
      return {
        type,
        id,
        ...additionalEntries,
      }
    }
  }
}

function parseTurn(value: unknown): CodexTurn | null {
  if (!isRecord(value) || !isString(value['id']) || !Array.isArray(value['items'])) {
    return null
  }
  const status = value['status']
  if (status !== 'completed' && status !== 'interrupted' && status !== 'failed' && status !== 'inProgress') {
    return null
  }
  const items: ParsedThreadItem[] = []
  for (const entry of value['items']) {
    const item = parseThreadItem(entry)
    if (!item) {
      return null
    }
    items.push(item)
  }
  const error = parseTurnError(value['error'])
  if (error === null && value['error'] !== null) {
    return null
  }
  return {
    id: value['id'],
    status,
    items,
    error,
  }
}

function parseThread(value: unknown): CodexThread | null {
  if (!isRecord(value) || !isString(value['id']) || !isString(value['preview'])) {
    return null
  }
  if (!Array.isArray(value['turns'])) {
    return null
  }
  const turns: CodexTurn[] = []
  for (const entry of value['turns']) {
    const turn = parseTurn(entry)
    if (!turn) {
      return null
    }
    turns.push(turn)
  }
  return {
    id: value['id'],
    preview: value['preview'],
    turns,
  }
}

function parseThreadStartResponse(value: unknown): CodexThreadStartResponse | null {
  if (!isRecord(value)) {
    return null
  }
  const thread = parseThread(value['thread'])
  return thread ? { thread } : null
}

function parseTurnStartResponse(value: unknown): CodexTurnStartResponse | null {
  if (!isRecord(value)) {
    return null
  }
  const turn = parseTurn(value['turn'])
  return turn ? { turn } : null
}

function parseJsonRpcError(value: unknown): CodexJsonRpcError | null {
  if (!isRecord(value) || !isFiniteNumber(value['code']) || !isString(value['message'])) {
    return null
  }
  if (value['data'] !== undefined && !isJsonValue(value['data'])) {
    return null
  }
  return {
    code: value['code'],
    message: value['message'],
    ...(value['data'] !== undefined ? { data: value['data'] } : {}),
  }
}

function parseJsonRpcResponse(value: unknown): CodexJsonRpcResponse<unknown> | null {
  if (!isRecord(value) || !isFiniteNumber(value['id'])) {
    return null
  }
  const hasResult = Object.hasOwn(value, 'result')
  const hasError = Object.hasOwn(value, 'error')
  if (!hasResult && !hasError) {
    return null
  }
  const response: CodexJsonRpcResponse<unknown> = {
    id: value['id'],
  }
  if (hasResult) {
    response.result = value['result']
  }
  if (hasError) {
    const error = parseJsonRpcError(value['error'])
    if (!error) {
      return null
    }
    response.error = error
  }
  return response
}

function parseDynamicToolCallRequest(value: unknown): CodexDynamicToolCallRequest | null {
  if (
    !isRecord(value) ||
    !isString(value['threadId']) ||
    !isString(value['turnId']) ||
    !isString(value['callId']) ||
    !isString(value['tool']) ||
    !isJsonValue(value['arguments'])
  ) {
    return null
  }
  return {
    threadId: value['threadId'],
    turnId: value['turnId'],
    callId: value['callId'],
    tool: value['tool'],
    arguments: value['arguments'],
  }
}

function parseServerRequest(value: unknown): ParsedServerRequest | null {
  if (!isRecord(value) || !isFiniteNumber(value['id']) || !isString(value['method'])) {
    return null
  }
  if (value['method'] === 'item/tool/call') {
    const params = parseDynamicToolCallRequest(value['params'])
    return params
      ? {
          method: 'item/tool/call',
          id: value['id'],
          params,
        }
      : null
  }
  if (value['params'] !== undefined && !isJsonValue(value['params'])) {
    return null
  }
  return {
    method: value['method'],
    id: value['id'],
    ...(value['params'] !== undefined ? { params: value['params'] } : {}),
  }
}

function parseServerNotification(value: unknown): CodexServerNotification | null {
  if (!isRecord(value) || !isString(value['method']) || !isRecord(value['params'])) {
    return null
  }
  const method = value['method']
  const params = value['params']
  switch (method) {
    case 'thread/started': {
      const thread = parseThread(params['thread'])
      return thread ? { method, params: { thread } } : null
    }
    case 'turn/started':
    case 'turn/completed': {
      if (!isString(params['threadId'])) {
        return null
      }
      const turn = parseTurn(params['turn'])
      return turn
        ? {
            method,
            params: {
              threadId: params['threadId'],
              turn,
            },
          }
        : null
    }
    case 'item/started':
    case 'item/completed': {
      if (!isString(params['threadId']) || !isString(params['turnId'])) {
        return null
      }
      const item = parseThreadItem(params['item'])
      return item
        ? {
            method,
            params: {
              threadId: params['threadId'],
              turnId: params['turnId'],
              item,
            },
          }
        : null
    }
    case 'item/agentMessage/delta':
    case 'item/plan/delta':
    case 'item/reasoning/delta':
      return isString(params['threadId']) && isString(params['turnId']) && isString(params['itemId']) && isString(params['delta'])
        ? {
            method,
            params: {
              threadId: params['threadId'],
              turnId: params['turnId'],
              itemId: params['itemId'],
              delta: params['delta'],
            },
          }
        : null
    case 'error': {
      if (params['message'] !== undefined && !isString(params['message'])) {
        return null
      }
      const extraParams: Record<string, ParsedJsonValue | undefined> = {}
      for (const [key, entry] of Object.entries(params)) {
        if (entry !== undefined && !isJsonValue(entry)) {
          return null
        }
        extraParams[key] = entry
      }
      return {
        method,
        params: extraParams,
      }
    }
    default:
      return null
  }
}

function expectInitializeResponse(value: unknown): CodexInitializeResponse {
  const response = parseInitializeResponse(value)
  if (!response) {
    throw new Error('Invalid Codex initialize response')
  }
  return response
}

function expectThreadStartResponse(value: unknown): CodexThreadStartResponse {
  const response = parseThreadStartResponse(value)
  if (!response) {
    throw new Error('Invalid Codex thread response')
  }
  return response
}

function expectTurnStartResponse(value: unknown): CodexTurnStartResponse {
  const response = parseTurnStartResponse(value)
  if (!response) {
    throw new Error('Invalid Codex turn response')
  }
  return response
}

export class CodexAppServerClient implements CodexAppServerTransport {
  private readonly command: string
  private readonly args: string[]
  private readonly cwd: string | undefined
  private readonly env: NodeJS.ProcessEnv
  private readonly onLog: ((message: string) => void) | undefined
  private readonly handleDynamicToolCall
  private readonly pending = new Map<CodexRequestId, PendingResponse>()
  private readonly notificationListeners = new Set<(notification: CodexServerNotification) => void>()
  private process: ChildProcessWithoutNullStreams | null = null
  private reader: readline.Interface | null = null
  private nextRequestId = 1
  private initializePromise: Promise<CodexInitializeResponse> | null = null

  constructor(options: CodexAppServerClientOptions) {
    this.command = options.command ?? 'codex'
    this.args = options.args ?? ['app-server']
    this.cwd = options.cwd
    this.env = stripTelemetryEnv(options.env)
    this.onLog = options.onLog
    this.handleDynamicToolCall = options.handleDynamicToolCall
  }

  subscribe(listener: (notification: CodexServerNotification) => void): () => void {
    this.notificationListeners.add(listener)
    return () => {
      this.notificationListeners.delete(listener)
    }
  }

  async ensureReady(): Promise<CodexInitializeResponse> {
    if (this.initializePromise) {
      return await this.initializePromise
    }

    this.initializePromise = this.startProcess()
    try {
      return await this.initializePromise
    } catch (error) {
      this.initializePromise = null
      throw error
    }
  }

  async threadStart(input: {
    model: string
    approvalPolicy: CodexAppServerApprovalPolicy
    sandbox: CodexAppServerSandboxMode
    config?: CodexAppServerThreadConfig
    baseInstructions: string
    developerInstructions: string
    dynamicTools: readonly CodexDynamicToolSpec[]
  }): Promise<CodexThread> {
    await this.ensureReady()
    const result = expectThreadStartResponse(
      await this.request('thread/start', {
        model: input.model,
        approvalPolicy: input.approvalPolicy,
        sandbox: input.sandbox,
        ...(input.config === undefined ? {} : { config: input.config }),
        baseInstructions: input.baseInstructions,
        developerInstructions: input.developerInstructions,
        dynamicTools: [...input.dynamicTools],
        experimentalRawEvents: false,
        persistExtendedHistory: true,
      }),
    )
    return result.thread
  }

  async threadResume(input: { threadId: string; baseInstructions: string; developerInstructions: string }): Promise<CodexThread> {
    await this.ensureReady()
    const result = expectThreadStartResponse(
      await this.request('thread/resume', {
        threadId: input.threadId,
        baseInstructions: input.baseInstructions,
        developerInstructions: input.developerInstructions,
        persistExtendedHistory: true,
      }),
    )
    return result.thread
  }

  async turnStart(input: { threadId: string; prompt: string }): Promise<CodexTurn> {
    await this.ensureReady()
    const result = expectTurnStartResponse(
      await this.request('turn/start', {
        threadId: input.threadId,
        input: [
          {
            type: 'text',
            text: input.prompt,
          } satisfies CodexUserInput,
        ],
      }),
    )
    return result.turn
  }

  async turnInterrupt(threadId: string): Promise<void> {
    await this.ensureReady()
    await this.request('turn/interrupt', { threadId })
  }

  async close(): Promise<void> {
    const activeProcess = this.process
    this.process = null
    this.reader?.close()
    this.reader = null
    this.initializePromise = null
    if (!activeProcess) {
      return
    }
    if (!activeProcess.killed) {
      activeProcess.kill('SIGTERM')
    }
    await new Promise<void>((resolve) => {
      activeProcess.once('close', () => resolve())
      activeProcess.once('error', () => resolve())
      setTimeout(resolve, 1_000)
    })
  }

  private async startProcess(): Promise<CodexInitializeResponse> {
    const child = spawn(this.command, this.args, {
      stdio: ['pipe', 'pipe', 'pipe'],
      ...(this.cwd ? { cwd: this.cwd } : {}),
      env: this.env,
    })
    this.process = child

    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => {
      this.onLog?.(chunk.trim())
    })

    child.once('error', (error) => {
      this.rejectAllPending(error instanceof Error ? error : new Error(String(error)))
    })
    child.once('close', (code, signal) => {
      this.rejectAllPending(new Error(`Codex app-server exited unexpectedly (${signal ?? 'code'}:${String(code ?? 'unknown')})`))
    })

    this.reader = readline.createInterface({ input: child.stdout })
    this.reader.on('line', (line) => {
      if (line.trim().length === 0) {
        return
      }
      void this.handleLine(line)
    })

    const initialized = expectInitializeResponse(
      await this.request('initialize', {
        clientInfo: {
          name: 'monolith',
          title: 'Bilig Monolith',
          version: '0.1.0',
        },
        capabilities: CODEX_INITIALIZE_CAPABILITIES,
      }),
    )
    this.notify('initialized', {})
    return initialized
  }

  private async handleLine(line: string): Promise<void> {
    let message: unknown
    try {
      message = JSON.parse(line)
    } catch (error) {
      this.onLog?.(`Failed to parse Codex app-server message: ${String(error)}`)
      return
    }

    if (!isRecord(message)) {
      return
    }

    const response = parseJsonRpcResponse(message)
    if (response) {
      this.handleResponse(response)
      return
    }

    const request = parseServerRequest(message)
    if (request) {
      await this.handleServerRequest(request)
      return
    }

    const notification = parseServerNotification(message)
    if (notification) {
      this.emitNotification(notification)
    }
  }

  private handleResponse(message: CodexJsonRpcResponse<unknown>): void {
    const pending = this.pending.get(message.id)
    if (!pending) {
      return
    }
    this.pending.delete(message.id)
    if (message.error) {
      pending.reject(new Error(message.error.message))
      return
    }
    pending.resolve(message.result)
  }

  private async handleServerRequest(request: ParsedServerRequest): Promise<void> {
    try {
      if (!isDynamicToolCallServerRequest(request)) {
        this.respondWithError(request.id, JSON_RPC_METHOD_NOT_FOUND, `Unsupported server request: ${request.method}`)
        return
      }
      const result = await this.handleDynamicToolCall(request.params)
      this.respondWithResult(request.id, result)
    } catch (error) {
      this.respondWithError(request.id, JSON_RPC_INTERNAL_ERROR, asError(error, 'Dynamic tool call failed').message)
    }
  }

  private emitNotification(notification: CodexServerNotification): void {
    this.notificationListeners.forEach((listener) => {
      listener(notification)
    })
  }

  private async request(method: string, params: unknown): Promise<unknown> {
    const activeProcess = this.process
    if (!activeProcess) {
      throw new Error('Codex app-server process is not running')
    }
    const id = this.nextRequestId++
    const response = new Promise<unknown>((resolve, reject) => {
      this.pending.set(id, {
        resolve,
        reject,
      })
    })
    this.write({
      method,
      id,
      params,
    })
    return await response
  }

  private notify(method: string, params: unknown): void {
    this.write({
      method,
      params,
    })
  }

  private write(payload: Record<string, unknown>): void {
    const activeProcess = this.process
    if (!activeProcess) {
      throw new Error('Codex app-server process is not running')
    }
    activeProcess.stdin.write(`${JSON.stringify(payload)}\n`)
  }

  private respondWithResult(id: CodexRequestId, result: CodexDynamicToolCallResult): void {
    this.write({
      id,
      result,
    })
  }

  private respondWithError(id: CodexRequestId, code: number, message: string): void {
    this.write({
      id,
      error: {
        code,
        message,
      },
    })
  }

  private rejectAllPending(error: Error): void {
    if (this.pending.size === 0) {
      return
    }
    const entries = [...this.pending.values()]
    this.pending.clear()
    entries.forEach((pending) => {
      pending.reject(error)
    })
  }
}
