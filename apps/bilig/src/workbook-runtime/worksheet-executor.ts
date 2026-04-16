import { decodeAgentFrame, encodeAgentFrame, type AgentFrame } from '@bilig/agent-api'
import { runPromise, type DocumentControlService } from '@bilig/runtime-kernel'
import type { AgentFrameContext } from './agent-routing.js'
import { normalizeBaseUrl } from './session-shared.js'

export interface WorksheetExecutor {
  execute(frame: AgentFrame): Promise<AgentFrame>
}

export interface HttpWorksheetExecutorOptions {
  baseUrl: string
  fetchImpl?: typeof fetch
}

export interface InProcessWorksheetExecutorOptions {
  documentService: DocumentControlService
  serverUrl?: string
  browserAppBaseUrl?: string
}

export function createHttpWorksheetExecutor(options: HttpWorksheetExecutorOptions): WorksheetExecutor {
  const baseUrl = normalizeBaseUrl(options.baseUrl)
  const fetchImpl = options.fetchImpl ?? fetch

  return {
    async execute(frame) {
      const response = await fetchImpl(`${baseUrl}/v2/agent/frames`, {
        method: 'POST',
        headers: {
          'content-type': 'application/octet-stream',
        },
        body: Buffer.from(encodeAgentFrame(frame)),
      })
      if (!response.ok) {
        throw new Error(`Worksheet executor request failed with status ${response.status}`)
      }
      return decodeAgentFrame(new Uint8Array(await response.arrayBuffer()))
    },
  }
}

export function createInProcessWorksheetExecutor(options: InProcessWorksheetExecutorOptions): WorksheetExecutor {
  const context: AgentFrameContext = {
    ...(options.serverUrl ? { serverUrl: options.serverUrl } : {}),
    ...(options.browserAppBaseUrl ? { browserAppBaseUrl: options.browserAppBaseUrl } : {}),
  }
  return {
    execute(frame) {
      return runPromise(options.documentService.handleAgentFrame(frame, context))
    },
  }
}
