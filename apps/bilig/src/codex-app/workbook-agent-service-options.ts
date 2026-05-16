import type { WorkbookAgentAppliedBy } from '@bilig/agent-api'
import type { WorkbookAgentStreamEvent, WorkbookAgentThreadSnapshot, WorkbookAgentThreadSummary } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import {
  DEFAULT_MAX_CODEX_CLIENTS,
  DEFAULT_MAX_CODEX_CONCURRENT_TURNS_PER_CLIENT,
  DEFAULT_MAX_CODEX_QUEUED_TURNS_PER_CLIENT,
} from './workbook-agent-codex-runtime.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import type { WorkbookAgentFeatureFlags } from './workbook-agent-feature-flags.js'
import type { WorkbookAgentObservabilitySnapshot } from './workbook-agent-session-registry.js'

const DEFAULT_MAX_ACTIVE_TURNS_PER_USER = parsePositiveIntegerEnv(process.env['BILIG_CODEX_MAX_ACTIVE_TURNS_PER_USER'], 4)
const DEFAULT_MAX_ACTIVE_TURNS_PER_DOCUMENT = parsePositiveIntegerEnv(process.env['BILIG_CODEX_MAX_ACTIVE_TURNS_PER_DOCUMENT'], 16)

export interface WorkbookAgentService {
  readonly enabled: boolean
  createSession(input: { documentId: string; session: SessionIdentity; body: unknown }): Promise<WorkbookAgentThreadSnapshot>
  updateContext(input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot>
  startTurn(input: { documentId: string; threadId: string; session: SessionIdentity; body: unknown }): Promise<WorkbookAgentThreadSnapshot>
  startWorkflow(input: {
    documentId: string
    threadId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot>
  cancelWorkflow(input: {
    documentId: string
    threadId: string
    runId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot>
  interruptTurn(input: { documentId: string; threadId: string; session: SessionIdentity }): Promise<WorkbookAgentThreadSnapshot>
  applyReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null
    preview: unknown
  }): Promise<WorkbookAgentThreadSnapshot>
  reviewReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
    body: unknown
  }): Promise<WorkbookAgentThreadSnapshot>
  dismissReviewItem(input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot>
  replayExecutionRecord(input: {
    documentId: string
    threadId: string
    recordId: string
    session: SessionIdentity
  }): Promise<WorkbookAgentThreadSnapshot>
  listThreads(input: { documentId: string; session: SessionIdentity }): Promise<WorkbookAgentThreadSummary[]>
  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot
  getSnapshot(input: { documentId: string; threadId: string; session: SessionIdentity }): WorkbookAgentThreadSnapshot
  subscribe(threadId: string, listener: (event: WorkbookAgentStreamEvent) => void): () => void
  close(): Promise<void>
}

export interface EnabledWorkbookAgentServiceOptions {
  zeroSyncService: ZeroSyncService
  codexClientFactory?: (options: CodexAppServerClientOptions) => CodexAppServerTransport
  now?: () => number
  maxSessions?: number
  maxCodexClients?: number
  maxConcurrentTurnsPerCodexClient?: number
  maxQueuedTurnsPerCodexClient?: number
  maxActiveTurnsPerUser?: number
  maxActiveTurnsPerDocument?: number
  featureFlags?: Partial<WorkbookAgentFeatureFlags>
}

export function resolveWorkbookAgentServiceLimits(options: EnabledWorkbookAgentServiceOptions): {
  readonly maxSessions: number
  readonly maxCodexClients: number
  readonly maxConcurrentTurnsPerCodexClient: number
  readonly maxQueuedTurnsPerCodexClient: number
  readonly maxActiveTurnsPerUser: number
  readonly maxActiveTurnsPerDocument: number
} {
  return {
    maxSessions: options.maxSessions ?? 64,
    maxCodexClients: options.maxCodexClients ?? DEFAULT_MAX_CODEX_CLIENTS,
    maxConcurrentTurnsPerCodexClient: options.maxConcurrentTurnsPerCodexClient ?? DEFAULT_MAX_CODEX_CONCURRENT_TURNS_PER_CLIENT,
    maxQueuedTurnsPerCodexClient: options.maxQueuedTurnsPerCodexClient ?? DEFAULT_MAX_CODEX_QUEUED_TURNS_PER_CLIENT,
    maxActiveTurnsPerUser: options.maxActiveTurnsPerUser ?? DEFAULT_MAX_ACTIVE_TURNS_PER_USER,
    maxActiveTurnsPerDocument: options.maxActiveTurnsPerDocument ?? DEFAULT_MAX_ACTIVE_TURNS_PER_DOCUMENT,
  }
}

function parsePositiveIntegerEnv(value: string | undefined, fallback: number): number {
  if (!value) {
    return fallback
  }
  const parsed = Number.parseInt(value, 10)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback
}
