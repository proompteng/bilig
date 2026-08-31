import { handleMutateRequest, handleQueryRequest } from '@rocicorp/zero/server'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { resolveRequestBaseUrl } from '@bilig/runtime-kernel'
import {
  checkWorkbookCommandResultForBundle,
  type WorkbookCommandResult,
  type WorkbookPlanData,
  type WorkbookRunResult,
  type WorkbookUndoRef,
} from '@bilig/workbook'
import {
  type AuthoritativeWorkbookEventBatch,
  executeZeroQueryTransform,
  isAuthoritativeWorkbookEventBatch,
  schema,
  type WorkbookChangeUndoBundle,
} from '@bilig/zero-sync'
import {
  areWorkbookAgentPreviewSummariesEqual,
  buildWorkbookAgentPreview,
  toAppliedWorkbookCommandResult,
  toWorkbookCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentExecutionRecord,
  type WorkbookAgentPreviewSummary,
} from '@bilig/agent-api'
import type { BiligAuthMode, SessionIdentity } from '../http/session.js'
import { WorkbookRuntimeManager, type WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { createWorkbookRuntimeStoreConnection, createZeroDbProvider, createZeroPool, resolveZeroDatabaseUrl } from './db.js'
import { handleServerMutator } from './server-mutators.js'
import { ZeroRecalcWorker } from './recalc-worker.js'
import { loadWorkbookEventRecordsAfter } from './store.js'
import type { WorkbookRuntimeStoreConnection } from './store.js'
import { applyWorkbookAgentCommandBundleWithUndoCapture } from './workbook-agent-apply.js'
import {
  runStrictWorkbookPlanData,
  workbookPlanRunAppliedOps,
  workbookPlanRunResultProof,
  workbookPlanRunUndoBundle,
} from './workbook-plan-data-apply.js'
import {
  assertZeroDataMigrationsReady,
  ensureZeroDataMigrationSchema,
  resolveAllowPendingCleanupMigrations,
  resolveRunDataMigrationsOnBoot,
  runPendingZeroDataMigrations,
} from './data-migration-runner.js'
import { persistWorkbookMutation } from './workbook-mutation-store.js'
import { acquireWorkbookMutationLock, loadWorkbookRuntimeMetadata, loadWorkbookState } from './workbook-runtime-store.js'
import { ensureZeroPublication } from './publication-store.js'
import { createWorkbookChangeStoreConnection, listWorkbookChanges, type WorkbookChangeRecord } from './workbook-change-store.js'
import { ensureZeroServiceSchema } from './schema-bootstrap.js'
import { ensureWorkbookDocumentExists } from './workbook-migration-store.js'
import {
  appendWorkbookAgentRun,
  createWorkbookAgentRunStoreConnection,
  listWorkbookAgentThreadRuns,
  listWorkbookAgentRuns,
} from './workbook-agent-run-store.js'
import {
  createWorkbookChatThreadStoreConnection,
  listWorkbookAgentThreadSummaries,
  loadWorkbookAgentThreadState,
  saveWorkbookAgentThreadState,
  type WorkbookAgentThreadStateRecord,
} from './workbook-chat-thread-store.js'
import type { WorkbookAgentThreadSummary, WorkbookAgentWorkflowRun } from '@bilig/contracts'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import {
  createWorkbookWorkflowRunStoreConnection,
  listWorkbookThreadWorkflowRuns,
  upsertWorkbookWorkflowRun,
} from './workbook-workflow-run-store.js'

export interface ZeroSyncRequestLike {
  readonly protocol: string
  readonly method: string
  readonly url: string
  readonly headers: {
    readonly [key: string]: string | string[] | undefined
    readonly host?: string | string[] | undefined
  }
  readonly body?: unknown
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readRequiredDocumentId(args: unknown): string {
  if (!isRecord(args) || typeof args['documentId'] !== 'string' || args['documentId'].trim().length === 0) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_DOCUMENT_ID_REQUIRED',
      message: 'Workbook mutations require a non-empty documentId',
      statusCode: 400,
      retryable: false,
    })
  }
  return args['documentId']
}

function readZeroQueryDocumentIds(body: unknown): string[] {
  if (!Array.isArray(body) || body[0] !== 'transform' || !Array.isArray(body[1])) {
    return []
  }
  const documentIds = new Set<string>()
  for (const request of body[1]) {
    if (!isRecord(request) || !Array.isArray(request['args'])) {
      continue
    }
    documentIds.add(readRequiredDocumentId(request['args'][0]))
  }
  return [...documentIds]
}

export interface ZeroSyncService {
  readonly enabled: boolean
  /**
   * Reports whether an enabled service has completed its startup contract.
   * Disabled services are intentionally ready because local-only operation is
   * a supported mode of the app.
   */
  readonly isReady: () => boolean
  initialize(): Promise<void>
  close(): Promise<void>
  handleQuery(request: ZeroSyncRequestLike, session: SessionIdentity, authMode: BiligAuthMode): Promise<unknown>
  handleMutate(request: ZeroSyncRequestLike, session: SessionIdentity, authMode: BiligAuthMode): Promise<unknown>
  assertWorkbookAccess?(
    documentId: string,
    session: SessionIdentity,
    authMode: BiligAuthMode,
    options?: { readonly createIfMissing?: boolean },
  ): Promise<void>
  inspectWorkbook<T>(documentId: string, task: (runtime: WorkbookRuntime) => Promise<T> | T): Promise<T>
  applyServerMutator(name: string, args: unknown, session?: SessionIdentity): Promise<void>
  applyAgentCommandBundle(
    documentId: string,
    bundle: WorkbookAgentCommandBundle,
    preview: WorkbookAgentPreviewSummary,
    session?: SessionIdentity,
  ): Promise<{ revision: number; preview: WorkbookAgentPreviewSummary; commandResult?: WorkbookCommandResult }>
  applyWorkbookPlanData(
    documentId: string,
    plan: WorkbookPlanData,
    session?: SessionIdentity,
  ): Promise<{ revision: number; result: WorkbookRunResult }>
  listWorkbookChanges(documentId: string, limit?: number): Promise<WorkbookChangeRecord[]>
  listWorkbookAgentRuns(documentId: string, actorUserId: string, limit?: number): Promise<WorkbookAgentExecutionRecord[]>
  listWorkbookAgentThreadRuns(
    documentId: string,
    actorUserId: string,
    threadId: string,
    limit?: number,
  ): Promise<WorkbookAgentExecutionRecord[]>
  appendWorkbookAgentRun(record: WorkbookAgentExecutionRecord): Promise<void>
  listWorkbookAgentThreadSummaries(documentId: string, actorUserId: string): Promise<WorkbookAgentThreadSummary[]>
  loadWorkbookAgentThreadState(documentId: string, actorUserId: string, threadId: string): Promise<WorkbookAgentThreadStateRecord | null>
  saveWorkbookAgentThreadState(record: WorkbookAgentThreadStateRecord): Promise<void>
  listWorkbookThreadWorkflowRuns(
    documentId: string,
    actorUserId: string,
    threadId: string,
    limit?: number,
  ): Promise<WorkbookAgentWorkflowRun[]>
  upsertWorkbookWorkflowRun(documentId: string, run: WorkbookAgentWorkflowRun): Promise<void>
  getWorkbookHeadRevision(documentId: string): Promise<number>
  ensureWorkbookDocument?(documentId: string, ownerUserId?: string): Promise<void>
  loadLatestWorkbookSnapshot?(
    documentId: string,
  ): Promise<{ revision: number; calculatedRevision: number; snapshot: WorkbookSnapshot } | null>
  loadAuthoritativeEvents(documentId: string, afterRevision: number): Promise<AuthoritativeWorkbookEventBatch>
}

function fastifyRequestToWebRequest(request: ZeroSyncRequestLike): Request {
  const origin = resolveRequestBaseUrl(request, 'localhost')
  const headers = new Headers()
  for (const [key, value] of Object.entries(request.headers)) {
    if (Array.isArray(value)) {
      for (const entry of value) {
        headers.append(key, entry)
      }
      continue
    }
    if (typeof value === 'string') {
      headers.set(key, value)
    }
  }
  const body = request.body === undefined || request.body === null ? undefined : JSON.stringify(request.body)

  const init: RequestInit = {
    method: request.method,
    headers,
  }
  if (body !== undefined) {
    init.body = body
  }

  return new Request(new URL(request.url, origin), init)
}

class DisabledZeroSyncService implements ZeroSyncService {
  readonly enabled = false

  isReady(): boolean {
    return true
  }

  async initialize(): Promise<void> {}

  async close(): Promise<void> {}

  async handleQuery(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async handleMutate(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async assertWorkbookAccess(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async inspectWorkbook<T>(_documentId: string, _task: (runtime: WorkbookRuntime) => Promise<T> | T): Promise<T> {
    throw new Error('Zero sync is not configured')
  }

  async applyServerMutator(_name: string, _args: unknown, _session?: SessionIdentity): Promise<void> {
    throw new Error('Zero sync is not configured')
  }

  async applyAgentCommandBundle(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async applyWorkbookPlanData(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async listWorkbookChanges(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async listWorkbookAgentRuns(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async appendWorkbookAgentRun(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async listWorkbookAgentThreadRuns(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async listWorkbookAgentThreadSummaries(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async loadWorkbookAgentThreadState(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async saveWorkbookAgentThreadState(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async listWorkbookThreadWorkflowRuns(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async upsertWorkbookWorkflowRun(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async getWorkbookHeadRevision(): Promise<never> {
    throw new Error('Zero sync is not configured')
  }

  async loadAuthoritativeEvents(_documentId: string, _afterRevision: number): Promise<never> {
    throw new Error('Zero sync is not configured')
  }
}

class EnabledZeroSyncService implements ZeroSyncService {
  readonly enabled = true
  private readonly pool: ReturnType<typeof createZeroPool>
  private readonly dbProvider
  private readonly runtimeStore: WorkbookRuntimeStoreConnection
  private readonly runtimeManager: WorkbookRuntimeManager
  private readonly recalcWorker: ZeroRecalcWorker
  private lifecycle: 'new' | 'initializing' | 'ready' | 'failed' | 'closing' | 'closed' = 'new'
  private initializePromise: Promise<void> | null = null
  private closePromise: Promise<void> | null = null

  constructor(connectionString: string) {
    this.pool = createZeroPool(connectionString)
    this.dbProvider = createZeroDbProvider(connectionString)
    this.runtimeStore = createWorkbookRuntimeStoreConnection(this.pool, this.dbProvider)
    this.runtimeManager = new WorkbookRuntimeManager()
    this.recalcWorker = new ZeroRecalcWorker(this.pool, this.runtimeStore, this.runtimeManager)
  }

  isReady(): boolean {
    return this.lifecycle === 'ready'
  }

  async initialize(): Promise<void> {
    if (this.lifecycle === 'ready') {
      return
    }
    if (this.lifecycle === 'closing' || this.lifecycle === 'closed') {
      throw new Error('Zero sync cannot be initialized after shutdown has started')
    }
    if (this.initializePromise) {
      return await this.initializePromise
    }

    this.lifecycle = 'initializing'
    this.initializePromise = this.initializeOnce()
    try {
      await this.initializePromise
    } catch (error) {
      this.lifecycle = 'failed'
      throw error
    }
  }

  async close(): Promise<void> {
    this.closePromise ??= this.closeOnce()
    return await this.closePromise
  }

  private async initializeOnce(): Promise<void> {
    await ensureZeroServiceSchema(this.pool)
    await ensureZeroPublication(this.pool)
    await ensureZeroDataMigrationSchema(this.pool)
    if (resolveRunDataMigrationsOnBoot()) {
      await runPendingZeroDataMigrations(this.pool)
    }
    await assertZeroDataMigrationsReady(this.pool, {
      allowPendingCleanup: resolveAllowPendingCleanupMigrations(),
    })
    this.recalcWorker.start()
    this.lifecycle = 'ready'
  }

  private async closeOnce(): Promise<void> {
    this.lifecycle = 'closing'
    if (this.initializePromise) {
      await this.initializePromise.catch(() => undefined)
    }
    const failures: unknown[] = []
    try {
      this.recalcWorker.stop()
    } catch (error) {
      failures.push(error)
    }
    try {
      await this.runtimeManager.close()
    } catch (error) {
      failures.push(error)
    }
    try {
      await this.pool.end()
    } catch (error) {
      failures.push(error)
    }
    this.lifecycle = 'closed'
    if (failures.length > 0) {
      throw new AggregateError(failures, 'Failed to close Zero sync cleanly')
    }
  }

  async handleQuery(request: ZeroSyncRequestLike, session: SessionIdentity, authMode: BiligAuthMode): Promise<unknown> {
    await this.assertQueryWorkbookAccess(request.body, session, authMode)
    return await handleQueryRequest(
      (name, args) => executeZeroQueryTransform(name, args, session.userID),
      schema,
      fastifyRequestToWebRequest(request),
    )
  }

  async handleMutate(request: ZeroSyncRequestLike, session: SessionIdentity, authMode: BiligAuthMode): Promise<unknown> {
    return await handleMutateRequest(
      this.dbProvider,
      (transact) =>
        transact(async (tx, name, args) => {
          await this.assertWorkbookAccess(readRequiredDocumentId(args), session, authMode)
          return await handleServerMutator(tx, name, args, this.runtimeManager, session)
        }),
      fastifyRequestToWebRequest(request),
    )
  }

  async assertWorkbookAccess(
    documentId: string,
    session: SessionIdentity,
    authMode: BiligAuthMode,
    options: { readonly createIfMissing?: boolean } = {},
  ): Promise<void> {
    if (authMode === 'demo') {
      return
    }
    if (options.createIfMissing) {
      await ensureWorkbookDocumentExists(this.pool, documentId, session.userID)
    }
    if (session.roles.includes('admin')) {
      return
    }
    const metadata = await loadWorkbookRuntimeMetadata(this.runtimeStore, documentId)
    if (metadata.ownerUserId === session.userID) {
      return
    }
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_ACCESS_DENIED',
      message: `User ${session.userID} does not have access to workbook ${documentId}`,
      statusCode: 403,
      retryable: false,
    })
  }

  private async assertQueryWorkbookAccess(body: unknown, session: SessionIdentity, authMode: BiligAuthMode): Promise<void> {
    const documentIds = readZeroQueryDocumentIds(body)
    await Promise.all(documentIds.map((documentId) => this.assertWorkbookAccess(documentId, session, authMode)))
  }

  async inspectWorkbook<T>(documentId: string, task: (runtime: WorkbookRuntime) => Promise<T> | T): Promise<T> {
    return await this.runtimeManager.runExclusive(documentId, async () => {
      const runtime = await this.runtimeManager.loadRuntime(this.runtimeStore, documentId)
      return await task(runtime)
    })
  }

  async applyServerMutator(name: string, args: unknown, session?: SessionIdentity): Promise<void> {
    const client = await this.pool.connect()
    try {
      await client.query('BEGIN')
      const transactionDbProvider = createZeroDbProvider(client)
      await handleServerMutator(
        {
          run: transactionDbProvider.run.bind(transactionDbProvider),
          dbTransaction: {
            wrappedTransaction: client,
          },
        },
        name,
        args,
        this.runtimeManager,
        session,
      )
      await client.query('COMMIT')
    } catch (error) {
      await client.query('ROLLBACK').catch(() => undefined)
      throw error
    } finally {
      client.release()
    }
  }

  async applyAgentCommandBundle(
    documentId: string,
    bundle: WorkbookAgentCommandBundle,
    preview: WorkbookAgentPreviewSummary,
    session?: SessionIdentity,
  ): Promise<{ revision: number; preview: WorkbookAgentPreviewSummary; commandResult: WorkbookCommandResult }> {
    const client = await this.pool.connect()
    try {
      return await this.runtimeManager.runExclusive(documentId, async () => {
        await client.query('BEGIN')
        try {
          await acquireWorkbookMutationLock(client, documentId)
          const transactionRuntimeStore = createWorkbookRuntimeStoreConnection(client, createZeroDbProvider(client))
          const state = await this.runtimeManager.loadRuntime(transactionRuntimeStore, documentId)
          const commandBundle = assertWorkbookCommandBundleHandoff(bundle)
          if (state.headRevision !== commandBundle.targetRevision) {
            throw createWorkbookAgentServiceError({
              code: 'WORKBOOK_AGENT_PREVIEW_STALE',
              message: 'Workbook changed while the change set was being prepared. Run the request again to prepare a fresh change set.',
              statusCode: 409,
              retryable: true,
            })
          }
          const authoritativePreview = await buildWorkbookAgentPreview({
            snapshot: state.engine.exportSnapshot(),
            replicaId: `server:${documentId}:agent-preview:r${String(state.headRevision)}`,
            bundle,
          })
          if (!areWorkbookAgentPreviewSummariesEqual(preview, authoritativePreview)) {
            throw createWorkbookAgentServiceError({
              code: 'WORKBOOK_AGENT_PREVIEW_MISMATCH',
              message: 'Local workbook state changed before apply. Run the request again to refresh the change set.',
              statusCode: 409,
              retryable: true,
            })
          }
          const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(state.engine, bundle)
          const ownerUserId = resolveOwnerUserId(state, session)
          const result = await persistWorkbookMutation(client, documentId, {
            previousState: state,
            nextEngine: state.engine,
            updatedBy: session?.userID ?? 'system',
            ownerUserId,
            eventPayload: {
              kind: 'applyAgentCommandBundle',
              bundle,
            },
            undoBundle,
          })
          const commandResult = assertAppliedWorkbookCommandResult({
            bundle,
            revision: result.revision,
            ...optionalWorkbookUndoRef(workbookUndoRefForAgentCommandBundle(bundle, undoBundle)),
          })
          this.runtimeManager.commitMutation(documentId, {
            projectionCommit: result.projectionCommit,
            headRevision: result.revision,
            calculatedRevision: result.calculatedRevision,
            ownerUserId,
          })
          await client.query('COMMIT')
          return {
            revision: result.revision,
            preview: authoritativePreview,
            commandResult,
          }
        } catch (error) {
          this.runtimeManager.invalidate(documentId)
          await client.query('ROLLBACK').catch(() => undefined)
          throw error
        }
      })
    } finally {
      client.release()
    }
  }

  async applyWorkbookPlanData(
    documentId: string,
    plan: WorkbookPlanData,
    session?: SessionIdentity,
  ): Promise<{ revision: number; result: WorkbookRunResult }> {
    const client = await this.pool.connect()
    try {
      return await this.runtimeManager.runExclusive(documentId, async () => {
        await client.query('BEGIN')
        try {
          await acquireWorkbookMutationLock(client, documentId)
          const transactionRuntimeStore = createWorkbookRuntimeStoreConnection(client, createZeroDbProvider(client))
          const state = await this.runtimeManager.loadRuntime(transactionRuntimeStore, documentId)
          const result = await runStrictWorkbookPlanData(state.engine, plan, state.headRevision)
          if (result.status === 'failed') {
            this.runtimeManager.invalidate(documentId)
            await client.query('ROLLBACK').catch(() => undefined)
            return { revision: state.headRevision, result }
          }
          const appliedOps = workbookPlanRunAppliedOps(result)
          if (appliedOps.length === 0) {
            await client.query('COMMIT')
            return { revision: state.headRevision, result }
          }
          const runResultProof = workbookPlanRunResultProof(result)
          const ownerUserId = resolveOwnerUserId(state, session)
          const persisted = await persistWorkbookMutation(client, documentId, {
            previousState: state,
            nextEngine: state.engine,
            updatedBy: session?.userID ?? 'system',
            ownerUserId,
            eventPayload: {
              kind: 'applyWorkbookPlanData',
              plan,
              appliedOps: structuredClone([...appliedOps]),
              result: runResultProof,
            },
            undoBundle: workbookPlanRunUndoBundle(result),
          })
          this.runtimeManager.commitMutation(documentId, {
            projectionCommit: persisted.projectionCommit,
            headRevision: persisted.revision,
            calculatedRevision: persisted.calculatedRevision,
            ownerUserId,
          })
          await client.query('COMMIT')
          return { revision: persisted.revision, result }
        } catch (error) {
          this.runtimeManager.invalidate(documentId)
          await client.query('ROLLBACK').catch(() => undefined)
          throw error
        }
      })
    } finally {
      client.release()
    }
  }

  async listWorkbookAgentRuns(documentId: string, actorUserId: string, limit = 20): Promise<WorkbookAgentExecutionRecord[]> {
    return await listWorkbookAgentRuns(createWorkbookAgentRunStoreConnection(this.runtimeStore), {
      documentId,
      actorUserId,
      limit,
    })
  }

  async listWorkbookChanges(documentId: string, limit = 10): Promise<WorkbookChangeRecord[]> {
    return await listWorkbookChanges(createWorkbookChangeStoreConnection(this.runtimeStore), {
      documentId,
      limit,
    })
  }

  async listWorkbookAgentThreadRuns(
    documentId: string,
    actorUserId: string,
    threadId: string,
    limit?: number,
  ): Promise<WorkbookAgentExecutionRecord[]> {
    return await listWorkbookAgentThreadRuns(createWorkbookAgentRunStoreConnection(this.runtimeStore), {
      documentId,
      actorUserId,
      threadId,
      ...(limit === undefined ? {} : { limit }),
    })
  }

  async appendWorkbookAgentRun(record: WorkbookAgentExecutionRecord): Promise<void> {
    await appendWorkbookAgentRun(this.pool, record)
  }

  async listWorkbookAgentThreadSummaries(documentId: string, actorUserId: string): Promise<WorkbookAgentThreadSummary[]> {
    return await listWorkbookAgentThreadSummaries(createWorkbookChatThreadStoreConnection(this.runtimeStore), {
      documentId,
      actorUserId,
    })
  }

  async loadWorkbookAgentThreadState(
    documentId: string,
    actorUserId: string,
    threadId: string,
  ): Promise<WorkbookAgentThreadStateRecord | null> {
    return await loadWorkbookAgentThreadState(createWorkbookChatThreadStoreConnection(this.runtimeStore), {
      documentId,
      actorUserId,
      threadId,
    })
  }

  async saveWorkbookAgentThreadState(record: WorkbookAgentThreadStateRecord): Promise<void> {
    await saveWorkbookAgentThreadState(this.pool, record)
  }

  async listWorkbookThreadWorkflowRuns(
    documentId: string,
    actorUserId: string,
    threadId: string,
    limit?: number,
  ): Promise<WorkbookAgentWorkflowRun[]> {
    return await listWorkbookThreadWorkflowRuns(createWorkbookWorkflowRunStoreConnection(this.runtimeStore), {
      documentId,
      actorUserId,
      threadId,
      ...(limit === undefined ? {} : { limit }),
    })
  }

  async upsertWorkbookWorkflowRun(documentId: string, run: WorkbookAgentWorkflowRun): Promise<void> {
    await upsertWorkbookWorkflowRun(this.pool, {
      documentId,
      run,
    })
  }

  async getWorkbookHeadRevision(documentId: string): Promise<number> {
    const metadata = await loadWorkbookRuntimeMetadata(this.runtimeStore, documentId)
    return metadata.headRevision
  }

  async ensureWorkbookDocument(documentId: string, ownerUserId = 'system'): Promise<void> {
    await ensureWorkbookDocumentExists(this.pool, documentId, ownerUserId)
  }

  async loadLatestWorkbookSnapshot(
    documentId: string,
  ): Promise<{ revision: number; calculatedRevision: number; snapshot: WorkbookSnapshot } | null> {
    const metadata = await loadWorkbookRuntimeMetadata(this.runtimeStore, documentId)
    if (metadata.headRevision === 0) {
      return null
    }
    const state = await loadWorkbookState(this.runtimeStore, documentId)
    return {
      revision: state.headRevision,
      calculatedRevision: state.calculatedRevision,
      snapshot: state.snapshot,
    }
  }

  async loadAuthoritativeEvents(documentId: string, afterRevision: number): Promise<AuthoritativeWorkbookEventBatch> {
    const metadata = await loadWorkbookRuntimeMetadata(this.runtimeStore, documentId)
    const events = metadata.headRevision > afterRevision ? await loadWorkbookEventRecordsAfter(this.pool, documentId, afterRevision) : []
    const eventBatch = {
      afterRevision,
      headRevision: metadata.headRevision,
      calculatedRevision: metadata.calculatedRevision,
      events,
    }
    if (!isAuthoritativeWorkbookEventBatch(eventBatch)) {
      throw new Error(
        `Invalid authoritative workbook event batch for ${documentId}: expected contiguous events from r${String(afterRevision + 1)} through r${String(metadata.headRevision)}`,
      )
    }
    return eventBatch
  }
}

function resolveOwnerUserId(state: { ownerUserId: string }, session?: SessionIdentity): string {
  if (state.ownerUserId !== 'system' || !session?.userID) {
    return state.ownerUserId
  }
  return session.userID
}

function assertWorkbookCommandBundleHandoff(bundle: WorkbookAgentCommandBundle) {
  try {
    return toWorkbookCommandBundle(bundle)
  } catch (error) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_INVALID_COMMAND_BUNDLE',
      message: error instanceof Error ? error.message : String(error),
      statusCode: 400,
      retryable: false,
    })
  }
}

function workbookUndoRefForAgentCommandBundle(
  bundle: WorkbookAgentCommandBundle,
  undoBundle: WorkbookChangeUndoBundle | null,
): WorkbookUndoRef | undefined {
  if (undoBundle === null) {
    return undefined
  }
  if (undoBundle.kind !== 'engineOps') {
    return {
      id: `${bundle.id}:undo`,
    }
  }
  if (undoBundle.ops.length === 0) {
    return undefined
  }
  return {
    id: `${bundle.id}:undo`,
    ops: undoBundle.ops,
  }
}

function assertAppliedWorkbookCommandResult(input: {
  readonly bundle: WorkbookAgentCommandBundle
  readonly revision: number
  readonly undo?: WorkbookUndoRef
}): WorkbookCommandResult {
  try {
    const commandBundle = assertWorkbookCommandBundleHandoff(input.bundle)
    const commandResult = toAppliedWorkbookCommandResult(input)
    const check = checkWorkbookCommandResultForBundle(commandBundle, commandResult)
    if (check.status === 'invalid') {
      const [firstIssue] = check.issues
      throw new Error(firstIssue === undefined ? 'Workbook command result is invalid' : firstIssue.message)
    }
    return check.result
  } catch (error) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_INVALID_COMMAND_RESULT',
      message: error instanceof Error ? error.message : String(error),
      statusCode: 500,
      retryable: false,
    })
  }
}

function optionalWorkbookUndoRef(undo: WorkbookUndoRef | undefined): { readonly undo?: WorkbookUndoRef } {
  return undo === undefined ? {} : { undo }
}

export function createZeroSyncService(): ZeroSyncService {
  const connectionString = resolveZeroDatabaseUrl()
  if (!connectionString) {
    return new DisabledZeroSyncService()
  }
  return new EnabledZeroSyncService(connectionString)
}
