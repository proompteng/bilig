import type { WorkbookAgentExecutionRecord, WorkbookAgentCommandBundle, CodexServerNotification } from '@bilig/agent-api'
import type { WorkbookAgentThreadSnapshot, WorkbookAgentStreamEvent } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import {
  CodexAppServerClient,
  type CodexAppServerClientOptions,
  type CodexAppServerThreadConfig,
  type CodexAppServerTransport,
} from './codex-app-server-client.js'
import {
  CodexAppServerClientPool,
  type CodexAppServerClientPoolStats,
  type CodexAppServerPoolBackpressureError,
} from './codex-app-server-pool.js'
import { routeWorkbookAgentCodexNotification } from './workbook-agent-codex-notification-router.js'
import { createWorkbookAgentDynamicToolHandler } from './workbook-agent-dynamic-tool-handler.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import { workbookAgentDynamicToolSpecs, type WorkbookAgentStartWorkflowRequest } from './workbook-agent-tools.js'
import { createWorkbookAgentBaseInstructions, createWorkbookAgentDeveloperInstructions } from './workbook-agent-session-model.js'

function parsePositiveIntegerEnv(value: string | undefined, fallback: number): number {
  if (!value) {
    return fallback
  }
  const parsed = Number.parseInt(value, 10)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback
}

const DEFAULT_MODEL = process.env['BILIG_CODEX_MODEL']?.trim() || 'gpt-5.5'
const CODEX_APP_SERVER_ARGS = [
  'app-server',
  '-c',
  'analytics.enabled=false',
  '-c',
  'approval_policy="never"',
  '-c',
  'sandbox_mode="danger-full-access"',
  '-c',
  'network_access=true',
  '-c',
  'web_search="live"',
] as const
const WORKBOOK_AGENT_CODEX_THREAD_CONFIG = {
  approval_policy: 'never',
  sandbox_mode: 'danger-full-access',
  network_access: true,
  web_search: 'live',
  tools: {
    view_image: true,
  },
} as const satisfies CodexAppServerThreadConfig

export const DEFAULT_MAX_CODEX_CLIENTS = parsePositiveIntegerEnv(process.env['BILIG_CODEX_MAX_CLIENTS'], 4)
export const DEFAULT_MAX_CODEX_CONCURRENT_TURNS_PER_CLIENT = parsePositiveIntegerEnv(
  process.env['BILIG_CODEX_MAX_CONCURRENT_TURNS_PER_CLIENT'],
  1,
)
export const DEFAULT_MAX_CODEX_QUEUED_TURNS_PER_CLIENT = parsePositiveIntegerEnv(process.env['BILIG_CODEX_MAX_QUEUED_TURNS_PER_CLIENT'], 8)

export interface WorkbookAgentCodexRuntimeOptions {
  readonly zeroSyncService: ZeroSyncService
  readonly codexClientFactory?: (options: CodexAppServerClientOptions) => CodexAppServerTransport
  readonly now: () => number
  readonly maxCodexClients: number
  readonly maxConcurrentTurnsPerCodexClient: number
  readonly maxQueuedTurnsPerCodexClient: number
  readonly getSessionByThreadId: (threadId: string) => WorkbookAgentThreadState
  readonly tryGetSessionByThreadId: (threadId: string) => WorkbookAgentThreadState | null
  readonly listSessions: () => readonly WorkbookAgentThreadState[]
  readonly resolveTurnActorUserId: (sessionState: WorkbookAgentThreadState, turnId: string) => string
  readonly resolveTurnContext: (sessionState: WorkbookAgentThreadState, turnId: string) => WorkbookAgentThreadState['durable']['context']
  readonly stageReviewBundle: (sessionState: WorkbookAgentThreadState, turnId: string, bundle: WorkbookAgentCommandBundle) => void
  readonly shouldApplyToolBundleImmediately: (sessionState: WorkbookAgentThreadState, bundle: WorkbookAgentCommandBundle) => boolean
  readonly applyToolBundleAutomatically: (input: {
    sessionState: WorkbookAgentThreadState
    actorUserId: string
    bundle: WorkbookAgentCommandBundle
  }) => Promise<WorkbookAgentExecutionRecord | null>
  readonly persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>
  readonly emitSnapshot: (threadId: string) => void
  readonly emit: (threadId: string, event: WorkbookAgentStreamEvent) => void
  readonly finalizeCompletedTurn: (
    sessionState: WorkbookAgentThreadState,
    turnId: string,
    turnStatus: 'completed' | 'failed',
  ) => Promise<void>
  readonly startWorkflow: (input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: WorkbookAgentStartWorkflowRequest & {
      context?: WorkbookAgentThreadState['durable']['context']
    }
  }) => Promise<WorkbookAgentThreadSnapshot>
}

export class WorkbookAgentCodexRuntime {
  private readonly codexClientFactory: (options: CodexAppServerClientOptions) => CodexAppServerTransport
  private codexClient: CodexAppServerClientPool | null = null
  private unsubscribeCodex: (() => void) | null = null

  constructor(private readonly options: WorkbookAgentCodexRuntimeOptions) {
    this.codexClientFactory = options.codexClientFactory ?? ((clientOptions) => new CodexAppServerClient(clientOptions))
  }

  async getClient(): Promise<CodexAppServerTransport> {
    if (!this.codexClient) {
      this.codexClient = new CodexAppServerClientPool({
        codexClientFactory: this.codexClientFactory,
        maxClients: this.options.maxCodexClients,
        maxConcurrentTurnsPerClient: this.options.maxConcurrentTurnsPerCodexClient,
        maxQueuedTurnsPerClient: this.options.maxQueuedTurnsPerCodexClient,
        clientOptions: {
          command: process.env['BILIG_CODEX_BIN']?.trim() || 'codex',
          args: [...CODEX_APP_SERVER_ARGS],
          cwd: process.cwd(),
          env: process.env,
          handleDynamicToolCall: createWorkbookAgentDynamicToolHandler({
            zeroSyncService: this.options.zeroSyncService,
            now: this.options.now,
            getSessionByThreadId: this.options.getSessionByThreadId,
            resolveTurnActorUserId: this.options.resolveTurnActorUserId,
            resolveTurnContext: this.options.resolveTurnContext,
            stageReviewBundle: this.options.stageReviewBundle,
            shouldApplyToolBundleImmediately: this.options.shouldApplyToolBundleImmediately,
            applyToolBundleAutomatically: this.options.applyToolBundleAutomatically,
            persistSessionState: this.options.persistSessionState,
            emitSnapshot: this.options.emitSnapshot,
            startWorkflow: this.options.startWorkflow,
          }),
        },
      })
      await this.codexClient.ensureReady()
      this.unsubscribeCodex = this.codexClient.subscribe((notification) => {
        void this.handleNotification(notification)
      })
    }
    return this.codexClient
  }

  getStats(): CodexAppServerClientPoolStats | null {
    return this.codexClient?.getStats() ?? null
  }

  releaseThread(threadId: string): void {
    this.codexClient?.releaseThread(threadId)
  }

  async close(): Promise<void> {
    this.unsubscribeCodex?.()
    this.unsubscribeCodex = null
    await this.codexClient?.close()
    this.codexClient = null
  }

  private async handleNotification(notification: CodexServerNotification): Promise<void> {
    try {
      await routeWorkbookAgentCodexNotification({
        notification,
        listSessions: this.options.listSessions,
        tryGetSessionByThreadId: this.options.tryGetSessionByThreadId,
        finalizeCompletedTurn: this.options.finalizeCompletedTurn,
        persistSessionState: this.options.persistSessionState,
        emitSnapshot: this.options.emitSnapshot,
        emit: this.options.emit,
        now: this.options.now,
      })
    } catch {
      return
    }
  }
}

export function createWorkbookAgentThreadStartInput() {
  return {
    model: DEFAULT_MODEL,
    approvalPolicy: 'never' as const,
    sandbox: 'danger-full-access' as const,
    config: WORKBOOK_AGENT_CODEX_THREAD_CONFIG,
    baseInstructions: createWorkbookAgentBaseInstructions(),
    developerInstructions: createWorkbookAgentDeveloperInstructions(),
    dynamicTools: workbookAgentDynamicToolSpecs,
  }
}

export function createWorkbookAgentThreadResumeInput(threadId: string) {
  return {
    threadId,
    baseInstructions: createWorkbookAgentBaseInstructions(),
    developerInstructions: createWorkbookAgentDeveloperInstructions(),
  }
}

export function isWorkbookAgentCodexBackpressureError(value: unknown): value is CodexAppServerPoolBackpressureError {
  return value instanceof Error && value.name === 'CodexAppServerPoolBackpressureError'
}
