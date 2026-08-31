import { chmodSync, existsSync, lstatSync, mkdirSync } from 'node:fs'
import os from 'node:os'
import { join, resolve } from 'node:path'
import type { WorkbookAgentExecutionRecord, WorkbookAgentCommandBundle, CodexServerNotification } from '@bilig/agent-api'
import type { WorkbookAgentThreadSnapshot, WorkbookAgentStreamEvent } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import { logError } from '../runtime-logger.js'
import type { ZeroSyncService } from '../zero/service.js'
import {
  CodexAppServerClient,
  createCodexChildEnvironment,
  type CodexAppServerClientOptions,
  type CodexAppServerThreadConfig,
  type CodexAppServerTransport,
} from './codex-app-server-client.js'
import { CodexAppServerClientPool, type CodexAppServerClientPoolStats } from './codex-app-server-pool.js'
import { routeWorkbookAgentCodexNotification } from './workbook-agent-codex-notification-router.js'
import { createWorkbookAgentDynamicToolHandler } from './workbook-agent-dynamic-tool-handler.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import { workbookAgentDynamicToolSpecs, type WorkbookAgentStartWorkflowRequest } from './workbook-agent-tools.js'
import { createWorkbookAgentBaseInstructions, createWorkbookAgentDeveloperInstructions } from './workbook-agent-session-model.js'
import { parsePositiveIntegerEnv } from './workbook-agent-env.js'

const DEFAULT_MODEL = process.env['BILIG_CODEX_MODEL']?.trim() || 'gpt-5.5'
const UNSAFE_LOCAL_CODEX_OPT_IN = 'BILIG_CODEX_ALLOW_UNSAFE_LOCAL'
const SERVICE_CODEX_HOME = 'BILIG_CODEX_HOME'

const DISABLED_CODEX_FEATURES = [
  'apps',
  'browser_use',
  'browser_use_external',
  'browser_use_full_cdp_access',
  'code_mode',
  'code_mode_host',
  'computer_use',
  'goals',
  'hooks',
  'image_generation',
  'in_app_browser',
  'memories',
  'multi_agent',
  'multi_agent_v2',
  'plugins',
  'remote_plugin',
  'shell_snapshot',
  'shell_tool',
  'skill_mcp_dependency_install',
  'skill_search',
  'tool_suggest',
  'unified_exec',
  'view_image',
  'workspace_dependencies',
] as const

const DISABLED_CODEX_FEATURE_CONFIG = Object.fromEntries(DISABLED_CODEX_FEATURES.map((feature) => [feature, false]))

type CodexChildEnvironment = Readonly<Record<string, string | undefined>>

export interface WorkbookAgentCodexSecurityConfig {
  readonly command: string
  readonly args: readonly string[]
  readonly cwd: string
  readonly env: NodeJS.ProcessEnv
  readonly approvalPolicy: NonNullable<CodexAppServerThreadConfig['approval_policy']>
  readonly sandbox: NonNullable<CodexAppServerThreadConfig['sandbox_mode']>
  readonly networkAccess: boolean
  readonly webSearch: NonNullable<CodexAppServerThreadConfig['web_search']>
  readonly threadConfig: CodexAppServerThreadConfig
  readonly unsafeLocalOptIn: boolean
}

const SAFE_CODEX_SECURITY = {
  approvalPolicy: 'on-request' as const,
  sandbox: 'read-only' as const,
  networkAccess: false,
  webSearch: 'disabled' as const,
}

const LOCAL_CODEX_SECURITY = {
  approvalPolicy: 'on-request' as const,
  sandbox: 'workspace-write' as const,
  networkAccess: true,
  webSearch: 'live' as const,
}

type CodexSecurityPolicy = typeof SAFE_CODEX_SECURITY | typeof LOCAL_CODEX_SECURITY

function parseUnsafeLocalOptIn(env: CodexChildEnvironment): boolean {
  const value = env[UNSAFE_LOCAL_CODEX_OPT_IN]?.trim().toLowerCase()
  if (value === undefined || value.length === 0) {
    return false
  }
  if (value !== '1' && value !== 'true') {
    throw new Error(`${UNSAFE_LOCAL_CODEX_OPT_IN} must be true only when explicitly enabled`)
  }
  if (env['NODE_ENV'] !== 'development') {
    throw new Error(`${UNSAFE_LOCAL_CODEX_OPT_IN} is only permitted when NODE_ENV=development`)
  }
  return true
}

function resolveServiceCodexHome(env: CodexChildEnvironment, unsafeLocalOptIn: boolean): string {
  const configuredHome = env[SERVICE_CODEX_HOME]?.trim()
  const codexHome = resolve(configuredHome && configuredHome.length > 0 ? configuredHome : join(os.tmpdir(), 'bilig-codex-home'))

  if (!unsafeLocalOptIn) {
    const inheritedHomes = [env['CODEX_HOME'], env['HOME'] ? join(env['HOME'], '.codex') : undefined]
      .filter((value): value is string => value !== undefined && value.trim().length > 0)
      .map((value) => resolve(value))
    if (inheritedHomes.includes(codexHome)) {
      throw new Error(`${SERVICE_CODEX_HOME} must use a dedicated service directory, not the inherited user Codex home`)
    }
  }
  mkdirSync(codexHome, { recursive: true, mode: 0o700 })
  const stats = lstatSync(codexHome)
  if (!stats.isDirectory() || stats.isSymbolicLink()) {
    throw new Error(`${SERVICE_CODEX_HOME} must resolve to a real directory`)
  }
  if (!unsafeLocalOptIn) {
    chmodSync(codexHome, 0o700)
    for (const filename of ['config.toml', 'requirements.toml']) {
      if (existsSync(join(codexHome, filename))) {
        throw new Error(`${SERVICE_CODEX_HOME} must not contain ${filename}; app-server policy is supplied by Bilig`)
      }
    }
  }
  return codexHome
}

function createCodexAppServerArgs(security: CodexSecurityPolicy): readonly string[] {
  const args = [
    'app-server',
    '--strict-config',
    '-c',
    'analytics.enabled=false',
    '-c',
    `approval_policy="${security.approvalPolicy}"`,
    '-c',
    `sandbox_mode="${security.sandbox}"`,
    '-c',
    `sandbox_workspace_write.network_access=${String(security.networkAccess)}`,
    '-c',
    `web_search="${security.webSearch}"`,
    '-c',
    'mcp_servers={}',
  ]
  for (const feature of DISABLED_CODEX_FEATURES) {
    args.push('-c', `features.${feature}=false`)
  }
  return args
}

function createCodexThreadConfig(security: CodexSecurityPolicy): CodexAppServerThreadConfig {
  return {
    approval_policy: security.approvalPolicy,
    sandbox_mode: security.sandbox,
    sandbox_workspace_write: {
      network_access: security.networkAccess,
    },
    web_search: security.webSearch,
    mcp_servers: {},
    features: DISABLED_CODEX_FEATURE_CONFIG,
  }
}

/**
 * Resolve the Codex app-server trust boundary once at service startup.
 *
 * The production policy is intentionally explicit: app-server commands are
 * read-only, require approval for escalations, cannot use network/web search,
 * run outside the repository, and receive only the environment they need.
 * Local networked development requires an explicit development-only opt-in.
 */
export function resolveWorkbookAgentCodexSecurityConfig(env: CodexChildEnvironment = process.env): WorkbookAgentCodexSecurityConfig {
  const unsafeLocalOptIn = parseUnsafeLocalOptIn(env)
  const security = unsafeLocalOptIn ? LOCAL_CODEX_SECURITY : SAFE_CODEX_SECURITY
  const command = env['BILIG_CODEX_BIN']?.trim() || 'codex'
  const threadConfig = createCodexThreadConfig(security)
  const codexHome = resolveServiceCodexHome(env, unsafeLocalOptIn)

  return {
    command,
    args: createCodexAppServerArgs(security),
    cwd: os.tmpdir(),
    env: createCodexChildEnvironment(env, { codexHome }),
    approvalPolicy: security.approvalPolicy,
    sandbox: security.sandbox,
    networkAccess: security.networkAccess,
    webSearch: security.webSearch,
    threadConfig,
    unsafeLocalOptIn,
  }
}

export const DEFAULT_MAX_CODEX_CLIENTS = parsePositiveIntegerEnv(process.env['BILIG_CODEX_MAX_CLIENTS'], 4, 'BILIG_CODEX_MAX_CLIENTS')
export const DEFAULT_MAX_CODEX_CONCURRENT_TURNS_PER_CLIENT = parsePositiveIntegerEnv(
  process.env['BILIG_CODEX_MAX_CONCURRENT_TURNS_PER_CLIENT'],
  1,
  'BILIG_CODEX_MAX_CONCURRENT_TURNS_PER_CLIENT',
)
export const DEFAULT_MAX_CODEX_QUEUED_TURNS_PER_CLIENT = parsePositiveIntegerEnv(
  process.env['BILIG_CODEX_MAX_QUEUED_TURNS_PER_CLIENT'],
  8,
  'BILIG_CODEX_MAX_QUEUED_TURNS_PER_CLIENT',
)

export interface WorkbookAgentCodexRuntimeOptions {
  readonly zeroSyncService: ZeroSyncService
  readonly env?: CodexChildEnvironment
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
    assertApplyStillAuthorized?: (() => void) | null | undefined
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
  private readonly securityConfig: WorkbookAgentCodexSecurityConfig
  private codexClient: CodexAppServerClientPool | null = null
  private unsubscribeCodex: (() => void) | null = null

  constructor(private readonly options: WorkbookAgentCodexRuntimeOptions) {
    this.codexClientFactory = options.codexClientFactory ?? ((clientOptions) => new CodexAppServerClient(clientOptions))
    this.securityConfig = resolveWorkbookAgentCodexSecurityConfig(options.env)
  }

  async getClient(): Promise<CodexAppServerTransport> {
    if (!this.codexClient) {
      this.codexClient = new CodexAppServerClientPool({
        codexClientFactory: this.codexClientFactory,
        maxClients: this.options.maxCodexClients,
        maxConcurrentTurnsPerClient: this.options.maxConcurrentTurnsPerCodexClient,
        maxQueuedTurnsPerClient: this.options.maxQueuedTurnsPerCodexClient,
        clientOptions: {
          command: this.securityConfig.command,
          args: [...this.securityConfig.args],
          cwd: this.securityConfig.cwd,
          env: this.securityConfig.env,
          onLog: (message) => {
            if (message.length > 0) {
              logError(message)
            }
          },
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

  createThreadStartInput() {
    return createWorkbookAgentThreadStartInput(this.securityConfig)
  }

  createThreadResumeInput(threadId: string) {
    return createWorkbookAgentThreadResumeInput(threadId, this.securityConfig)
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
    } catch (error) {
      logError(error)
    }
  }
}

export function createWorkbookAgentThreadStartInput(
  securityConfig: Pick<
    WorkbookAgentCodexSecurityConfig,
    'approvalPolicy' | 'sandbox' | 'threadConfig'
  > = resolveWorkbookAgentCodexSecurityConfig(),
) {
  return {
    model: DEFAULT_MODEL,
    approvalPolicy: securityConfig.approvalPolicy,
    sandbox: securityConfig.sandbox,
    cwd: os.tmpdir(),
    runtimeWorkspaceRoots: [],
    environments: [],
    config: securityConfig.threadConfig,
    baseInstructions: createWorkbookAgentBaseInstructions(),
    developerInstructions: createWorkbookAgentDeveloperInstructions(),
    dynamicTools: workbookAgentDynamicToolSpecs,
  }
}

export function createWorkbookAgentThreadResumeInput(
  threadId: string,
  securityConfig: Pick<
    WorkbookAgentCodexSecurityConfig,
    'approvalPolicy' | 'sandbox' | 'threadConfig'
  > = resolveWorkbookAgentCodexSecurityConfig(),
) {
  return {
    threadId,
    approvalPolicy: securityConfig.approvalPolicy,
    sandbox: securityConfig.sandbox,
    cwd: os.tmpdir(),
    runtimeWorkspaceRoots: [],
    config: securityConfig.threadConfig,
    baseInstructions: createWorkbookAgentBaseInstructions(),
    developerInstructions: createWorkbookAgentDeveloperInstructions(),
  }
}
