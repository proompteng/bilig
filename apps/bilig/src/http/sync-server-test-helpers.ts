import { toWorkbookAgentReviewQueueItem, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { WorkbookAgentThreadSnapshot } from '@bilig/contracts'
import type { DocumentControlService } from '@bilig/runtime-kernel'
import { Effect } from 'effect'
import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import { createHmac } from 'node:crypto'
import { afterEach } from 'vitest'
import type { WorkbookAgentService } from '../codex-app/workbook-agent-service.js'
import type { ZeroSyncService } from '../zero/service.js'
import { createRequestSessionResolver } from './session.js'
export type { WorkbookAgentThreadSnapshot } from '@bilig/contracts'
export type { WorkbookSnapshot } from '@bilig/protocol'
export type { DocumentControlService } from '@bilig/runtime-kernel'
export { Effect } from 'effect'
export { describe, expect, it, vi } from 'vitest'
export type { WorkbookAgentService } from '../codex-app/workbook-agent-service.js'
export { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
export type { ZeroSyncService } from '../zero/service.js'
export { createAgentSkillDiscoveryIndex } from './agent-skill-discovery-routes.js'
export { resolveCanonicalDocsRedirectUrl } from './sync-server-spa.js'
export { createSyncServer } from './sync-server.js'

export type TestServer = Awaited<ReturnType<typeof startHttpServer>>

export async function startHttpServer(handler: (request: IncomingMessage, response: ServerResponse) => void) {
  const server = createServer(handler)
  await new Promise<void>((resolve, reject) => {
    server.listen(0, '127.0.0.1', () => resolve())
    server.once('error', reject)
  })
  const address = server.address()
  if (!address || typeof address === 'string') {
    throw new Error('Expected TCP test server address')
  }
  return {
    server,
    origin: `http://127.0.0.1:${String(address.port)}`,
    async close() {
      await new Promise<void>((resolve, reject) => {
        server.close((error) => {
          if (error) {
            reject(error)
            return
          }
          resolve()
        })
      })
    },
  }
}

export const upstreamServers: TestServer[] = []

export function createSignedProxyTestSession(userId: string, roles: readonly string[] = ['editor']) {
  const proxySecret = 'sync-server-test-proxy-secret-at-least-32-bytes'
  const timestamp = '1735689600'
  const roleHeader = roles.join(',')
  return {
    sessionResolver: createRequestSessionResolver({
      env: {
        BILIG_AUTH_MODE: 'signed-proxy',
        BILIG_AUTH_PROXY_SECRET: proxySecret,
        BILIG_SESSION_SECRET: 'sync-server-test-session-secret-at-least-32-bytes',
      },
      now: () => 1_735_689_600_000,
    }),
    headers: {
      'x-bilig-auth-user': userId,
      'x-bilig-auth-roles': roleHeader,
      'x-bilig-auth-timestamp': timestamp,
      'x-bilig-auth-signature': createHmac('sha256', proxySecret).update(`${timestamp}\n${userId}\n${roleHeader}`).digest('base64url'),
    },
  }
}

afterEach(async () => {
  delete process.env['BILIG_ZERO_PROXY_UPSTREAM']
  delete process.env['BILIG_PERSIST_STATE']
  delete process.env['BILIG_REMOTE_MCP_ALLOWED_ORIGINS']
  delete process.env['BILIG_REMOTE_MCP_ALLOW_LOCAL_ORIGINS']
  await Promise.all(upstreamServers.splice(0).map((server) => server.close()))
})

export function createZeroSyncStub(overrides: Partial<ZeroSyncService> = {}): ZeroSyncService {
  return {
    enabled: true,
    isReady: () => true,
    async initialize() {},
    async close() {},
    async handleQuery() {
      throw new Error('not used')
    },
    async handleMutate() {
      throw new Error('not used')
    },
    async assertWorkbookAccess() {},
    async inspectWorkbook<T>(_documentId: string, _task: (runtime: never) => T | Promise<T>) {
      throw new Error('not used')
    },
    async applyServerMutator() {
      throw new Error('not used')
    },
    async applyAgentCommandBundle() {
      throw new Error('not used')
    },
    async applyWorkbookPlanData() {
      throw new Error('not used')
    },
    async listWorkbookChanges() {
      return []
    },
    async listWorkbookAgentRuns() {
      return []
    },
    async listWorkbookAgentThreadRuns() {
      return []
    },
    async appendWorkbookAgentRun() {
      throw new Error('not used')
    },
    async listWorkbookAgentThreadSummaries() {
      return []
    },
    async loadWorkbookAgentThreadState() {
      return null
    },
    async saveWorkbookAgentThreadState() {
      throw new Error('not used')
    },
    async listWorkbookThreadWorkflowRuns() {
      return []
    },
    async upsertWorkbookWorkflowRun() {
      throw new Error('not used')
    },
    async getWorkbookHeadRevision() {
      return 1
    },
    async loadAuthoritativeEvents() {
      throw new Error('not used')
    },
    ...overrides,
  }
}

export function createWorkbookAgentServiceStub(overrides: Partial<WorkbookAgentService> = {}): WorkbookAgentService {
  return {
    enabled: true,
    async createSession() {
      throw new Error('not used')
    },
    async updateContext() {
      throw new Error('not used')
    },
    async startTurn() {
      throw new Error('not used')
    },
    async startWorkflow() {
      throw new Error('not used')
    },
    async cancelWorkflow() {
      throw new Error('not used')
    },
    async interruptTurn() {
      throw new Error('not used')
    },
    async applyReviewItem() {
      throw new Error('not used')
    },
    async reviewReviewItem() {
      throw new Error('not used')
    },
    async dismissReviewItem() {
      throw new Error('not used')
    },
    async replayExecutionRecord() {
      throw new Error('not used')
    },
    async listThreads() {
      return []
    },
    getObservabilitySnapshot() {
      return {
        enabled: true,
        generatedAtUnixMs: 1,
        featureFlags: {
          sharedThreadsEnabled: true,
          workflowRunnerEnabled: true,
          autoApplyLowRiskEnabled: true,
          formulaWorkflowFamilyEnabled: true,
          formattingWorkflowFamilyEnabled: true,
          importWorkflowFamilyEnabled: true,
          rollupWorkflowFamilyEnabled: true,
          structuralWorkflowFamilyEnabled: true,
          allowlistedUserCount: 0,
          allowlistedDocumentCount: 0,
        },
        sessions: {
          sessionCount: 0,
          subscriberThreadCount: 0,
          subscriberCount: 0,
          activeTurnCount: 0,
          runningWorkflowCount: 0,
          reviewQueueSessionCount: 0,
          sharedPendingReviewCount: 0,
        },
        pool: {
          slotCount: 0,
          boundThreadCount: 0,
          activeTurnCount: 0,
          queuedTurnCount: 0,
          maxClients: 0,
          maxConcurrentTurnsPerClient: 0,
          maxQueuedTurnsPerClient: 0,
        },
        counters: {
          turnBackpressureCount: 0,
          workflowStartedCount: 0,
          workflowCompletedCount: 0,
          workflowFailedCount: 0,
          workflowCancelledCount: 0,
          sharedReviewApprovedCount: 0,
          sharedReviewRejectedCount: 0,
          sharedRecommendationApprovedCount: 0,
          sharedRecommendationRejectedCount: 0,
        },
      }
    },
    getSnapshot() {
      throw new Error('not used')
    },
    subscribe() {
      return () => {}
    },
    async close() {},
    ...overrides,
  }
}

export function createDocumentServiceStub(overrides: Partial<DocumentControlService> = {}): DocumentControlService {
  return {
    attachBrowser() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    openBrowserSession() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    handleSyncFrame() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    handleAgentFrame() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    getDocumentState() {
      return Effect.sync(() => {
        throw new Error('not used')
      })
    },
    getLatestSnapshot() {
      return Effect.succeed(null)
    },
    ...overrides,
  }
}

export function createAgentSessionSnapshot(overrides: Partial<WorkbookAgentThreadSnapshot> = {}): WorkbookAgentThreadSnapshot {
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    executionPolicy: 'autoApplyAll',
    scope: 'private',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
    context: {
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    },
    entries: [],
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
  }
}

export function createReviewQueueItem(bundle: WorkbookAgentCommandBundle) {
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: bundle.sharedReview ? 'ownerReview' : 'manual',
    ...(bundle.sharedReview ? { sharedReview: bundle.sharedReview } : {}),
  })
}

export function readMcpToolNames(responseBody: unknown): string[] {
  const tools = readMcpTools(responseBody)
  return tools.map((tool) => {
    if (!isRecord(tool) || typeof tool['name'] !== 'string') {
      throw new Error(`Expected MCP tool definition, received ${JSON.stringify(tool)}`)
    }
    return tool['name']
  })
}

export function readMcpToolAnnotations(responseBody: unknown, toolName: string): Record<string, unknown> {
  const tool = readMcpTools(responseBody).find((candidate) => isRecord(candidate) && candidate['name'] === toolName)
  if (!isRecord(tool) || !isRecord(tool['annotations'])) {
    throw new Error(`Expected MCP tool annotations for ${toolName}, received ${JSON.stringify(responseBody)}`)
  }
  return tool['annotations']
}

function readMcpTools(responseBody: unknown): unknown[] {
  const result = isRecord(responseBody) ? responseBody['result'] : undefined
  const tools = isRecord(result) ? result['tools'] : undefined
  if (!Array.isArray(tools)) {
    throw new Error(`Expected tools/list response, received ${JSON.stringify(responseBody)}`)
  }
  return tools
}

export function readMcpServerCardToolNames(card: unknown): string[] {
  const tools = isRecord(card) ? card['tools'] : undefined
  if (!Array.isArray(tools)) {
    throw new Error(`Expected MCP server-card tools, received ${JSON.stringify(card)}`)
  }
  return tools.map((tool) => {
    if (!isRecord(tool) || typeof tool['name'] !== 'string') {
      throw new Error(`Expected MCP server-card tool, received ${JSON.stringify(tool)}`)
    }
    return tool['name']
  })
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
