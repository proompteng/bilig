import { Context, Data, Effect, Layer, Schema } from 'effect'

import type { AgentFrame } from '@bilig/agent-api'
import type { HelloFrame, ProtocolFrame } from '@bilig/binary-protocol'
import type { ErrorEnvelope, RuntimeSession } from '@bilig/contracts'
import type { DocumentStateSummary } from '@bilig/contracts'

export class TransportError extends Data.TaggedError('TransportError')<{
  readonly message: string
  readonly cause?: unknown
}> {}

export class HttpError extends Data.TaggedError('HttpError')<{
  readonly status: number
  readonly message: string
  readonly body?: string
}> {}

export class DecodeError extends Data.TaggedError('DecodeError')<{
  readonly message: string
  readonly cause?: unknown
}> {}

export interface ServerRuntimeConfig {
  readonly browserAppBaseUrl?: string
  readonly corsOrigin?: string
}

export interface BasicRequestLike {
  readonly protocol: string
  readonly headers: {
    readonly host?: string | string[] | undefined
  }
}

export type NormalizedWebSocket = {
  on(event: string, listener: (...args: unknown[]) => void): void
  send(data: Uint8Array): void
}

export interface AgentFrameContext {
  readonly serverUrl?: string
  readonly browserAppBaseUrl?: string
}

export interface SnapshotPayload {
  readonly cursor: number
  readonly contentType: string
  readonly bytes: Uint8Array
}

export interface DocumentControlService {
  readonly attachBrowser: (
    documentId: string,
    subscriberId: string,
    send: (frame: ProtocolFrame) => void,
  ) => Effect.Effect<() => void, TransportError>
  readonly openBrowserSession: (frame: HelloFrame) => Effect.Effect<ProtocolFrame[], TransportError>
  readonly handleSyncFrame: (frame: ProtocolFrame) => Effect.Effect<ProtocolFrame | ProtocolFrame[], TransportError>
  readonly handleAgentFrame: (frame: AgentFrame, context?: AgentFrameContext) => Effect.Effect<AgentFrame, TransportError>
  readonly getDocumentState: (documentId: string) => Effect.Effect<DocumentStateSummary, TransportError>
  readonly getLatestSnapshot: (documentId: string) => Effect.Effect<SnapshotPayload | null, TransportError>
}

export interface FetchService {
  readonly fetch: (input: RequestInfo | URL, init?: RequestInit) => Effect.Effect<Response, TransportError>
}

export const FetchService = Context.GenericTag<FetchService>('@bilig/runtime-kernel/FetchService')

export const BrowserFetchLayer = Layer.succeed(FetchService, {
  fetch(input: RequestInfo | URL, init?: RequestInit) {
    const target = typeof input === 'string' ? input : input instanceof URL ? input.toString() : 'request'
    return Effect.tryPromise({
      try: () => fetch(input, init),
      catch: (cause) =>
        new TransportError({
          message: `Failed to fetch ${target}`,
          cause,
        }),
    })
  },
})

export function provideBrowserFetch<Success, Failure, Requirements>(
  effect: Effect.Effect<Success, Failure, Requirements | FetchService>,
): Effect.Effect<Success, Failure, Requirements> {
  return effect.pipe(Effect.provide(BrowserFetchLayer))
}

export function normalizeBaseUrl(value: string): string {
  return value.endsWith('/') ? value.slice(0, -1) : value
}

export function resolveServerRuntimeConfig(env: Record<string, string | undefined>): ServerRuntimeConfig {
  const browserAppBaseUrl = env['BILIG_WEB_APP_BASE_URL']?.trim()
  const corsOrigin = env['BILIG_CORS_ORIGIN']?.trim()

  return {
    ...(browserAppBaseUrl ? { browserAppBaseUrl } : {}),
    ...(corsOrigin ? { corsOrigin } : {}),
  }
}

export function resolveRequestBaseUrl(request: BasicRequestLike, fallbackHost: string): string {
  const host = firstHeaderValue(request.headers.host) ?? fallbackHost
  return normalizeBaseUrl(`${request.protocol}://${host}`)
}

export function createErrorEnvelope(error: string, message: string, retryable: boolean): ErrorEnvelope {
  return {
    error,
    message,
    retryable,
  }
}

export function createGuestRuntimeSession(userId: string, roles: string[] = ['editor']): RuntimeSession {
  return {
    authToken: userId,
    userId,
    roles,
    isAuthenticated: false,
    authSource: 'guest',
  }
}

export function createRuntimeSession(details: {
  readonly authToken: string
  readonly userId: string
  readonly roles: string[]
  readonly isAuthenticated: boolean
  readonly authSource: RuntimeSession['authSource']
}): RuntimeSession {
  return {
    authToken: details.authToken,
    userId: details.userId,
    roles: details.roles,
    isAuthenticated: details.isAuthenticated,
    authSource: details.authSource,
  }
}

export function runPromise<Success, Failure>(effect: Effect.Effect<Success, Failure>): Promise<Success> {
  return Effect.runPromise(effect)
}

export function decodeWithSchema<Decoded, Encoded>(
  schema: Schema.Schema<Decoded, Encoded>,
  input: unknown,
): Effect.Effect<Decoded, DecodeError> {
  return Effect.try({
    try: () => Schema.decodeUnknownSync(schema)(input),
    catch: (cause) =>
      new DecodeError({
        message: 'Failed to decode payload',
        cause,
      }),
  })
}

export function ensureOkResponse(response: Response, message = 'Request failed'): Effect.Effect<Response, HttpError> {
  return response.ok
    ? Effect.succeed(response)
    : Effect.tryPromise({
        try: async () => {
          const body = await response.text()
          throw new HttpError({
            status: response.status,
            message,
            body,
          })
        },
        catch: (cause) => {
          if (cause instanceof HttpError) {
            return cause
          }
          return new HttpError({
            status: response.status,
            message,
          })
        },
      })
}

export function toMessageBytes(raw: unknown): Uint8Array {
  if (raw instanceof Buffer) {
    return new Uint8Array(raw)
  }
  if (raw instanceof ArrayBuffer) {
    return new Uint8Array(raw)
  }
  if (ArrayBuffer.isView(raw)) {
    return new Uint8Array(raw.buffer, raw.byteOffset, raw.byteLength)
  }
  throw new Error('Unsupported websocket payload')
}

type EventTargetWebSocket = {
  addEventListener(event: string, listener: (event: unknown) => void): void
  send(data: Uint8Array): void
}

function isNormalizedWebSocket(value: unknown): value is NormalizedWebSocket {
  return (
    typeof value === 'object' &&
    value !== null &&
    'on' in value &&
    typeof value.on === 'function' &&
    'send' in value &&
    typeof value.send === 'function'
  )
}

function hasSocket(value: unknown): value is { socket: unknown } {
  return typeof value === 'object' && value !== null && 'socket' in value
}

function hasWebSocket(value: unknown): value is { websocket: unknown } {
  return typeof value === 'object' && value !== null && 'websocket' in value
}

function isEventTargetWebSocket(value: unknown): value is EventTargetWebSocket {
  return (
    typeof value === 'object' &&
    value !== null &&
    'addEventListener' in value &&
    typeof value.addEventListener === 'function' &&
    'send' in value &&
    typeof value.send === 'function'
  )
}

function asNormalizedEventTargetSocket(socket: EventTargetWebSocket): NormalizedWebSocket {
  return {
    on(event, listener) {
      socket.addEventListener(event, (payload) => {
        if (event === 'message' && typeof payload === 'object' && payload !== null && 'data' in payload) {
          listener(payload.data)
          return
        }
        listener(payload)
      })
    },
    send(data) {
      socket.send(data)
    },
  }
}

export function normalizeWebSocket(candidate: unknown): NormalizedWebSocket {
  if (isNormalizedWebSocket(candidate)) {
    return candidate
  }
  if (isEventTargetWebSocket(candidate)) {
    return asNormalizedEventTargetSocket(candidate)
  }
  if (hasSocket(candidate) && isNormalizedWebSocket(candidate.socket)) {
    return candidate.socket
  }
  if (hasSocket(candidate) && isEventTargetWebSocket(candidate.socket)) {
    return asNormalizedEventTargetSocket(candidate.socket)
  }
  if (hasWebSocket(candidate)) {
    return normalizeWebSocket(candidate.websocket)
  }
  throw new Error('Unsupported websocket connection shape')
}

function firstHeaderValue(value: string | string[] | undefined): string | undefined {
  if (Array.isArray(value)) {
    return value.find((entry) => entry.length > 0)
  }
  return typeof value === 'string' && value.length > 0 ? value : undefined
}
