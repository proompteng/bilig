import { decodeFrame, encodeFrame, type ProtocolFrame } from '@bilig/binary-protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'

export interface UpstreamSyncRelay {
  send(batch: EngineOpBatch): Promise<void>
  disconnect(): Promise<void>
}

export interface HttpSyncRelayOptions {
  documentId: string
  baseUrl: string
  fetchImpl?: typeof fetch
  replicaId?: string
}

function normalizeBaseUrl(value: string): string {
  return value.endsWith('/') ? value.slice(0, -1) : value
}

async function readProtocolFrame(response: Response): Promise<ProtocolFrame> {
  const bytes = new Uint8Array(await response.arrayBuffer())
  return decodeFrame(bytes)
}

export function createHttpSyncRelay(options: HttpSyncRelayOptions): UpstreamSyncRelay {
  const fetchImpl = options.fetchImpl ?? fetch
  const baseUrl = normalizeBaseUrl(options.baseUrl)
  const replicaId = options.replicaId ?? `worksheet-host:${options.documentId}`
  const sessionId = `${options.documentId}:${replicaId}`
  let lastServerCursor = 0
  let connectPromise: Promise<void> | null = null
  let operationChain: Promise<void> = Promise.resolve()

  const sendFrame = async (frame: ProtocolFrame): Promise<ProtocolFrame> => {
    const response = await fetchImpl(`${baseUrl}/v2/documents/${encodeURIComponent(options.documentId)}/frames`, {
      method: 'POST',
      headers: {
        'content-type': 'application/octet-stream',
      },
      body: Buffer.from(encodeFrame(frame)),
    })
    if (!response.ok) {
      throw new Error(`Sync relay request failed with status ${response.status}`)
    }
    const nextFrame = await readProtocolFrame(response)
    if ('cursor' in nextFrame && typeof nextFrame.cursor === 'number') {
      lastServerCursor = Math.max(lastServerCursor, nextFrame.cursor)
    }
    if (nextFrame.kind === 'error') {
      throw new Error(nextFrame.message)
    }
    return nextFrame
  }

  const ensureConnected = async (): Promise<void> => {
    if (connectPromise) {
      return connectPromise
    }
    connectPromise = (async () => {
      await sendFrame({
        kind: 'hello',
        documentId: options.documentId,
        replicaId,
        sessionId,
        protocolVersion: 1,
        lastServerCursor,
        capabilities: ['local-relay'],
      })
    })()
    try {
      await connectPromise
    } catch (error) {
      connectPromise = null
      throw error
    }
  }

  const serializeOperation = <T>(operation: () => Promise<T>): Promise<T> => {
    const previous = operationChain
    const next = (async () => {
      try {
        await previous
      } catch {
        // Continue processing subsequent operations after a failure.
      }
      return await operation()
    })()
    operationChain = (async () => {
      try {
        await next
      } catch {
        // Swallow to keep the operation chain reusable.
      }
    })()
    return next
  }

  return {
    async send(batch) {
      await serializeOperation(async () => {
        await ensureConnected()
        await sendFrame({
          kind: 'appendBatch',
          documentId: options.documentId,
          cursor: lastServerCursor,
          batch,
        })
      })
    },
    async disconnect() {
      await serializeOperation(async () => {
        connectPromise = null
      })
    },
  }
}
