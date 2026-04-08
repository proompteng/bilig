import { handleMutateRequest, handleQueryRequest } from "@rocicorp/zero/server";
import type { SpreadsheetEngine } from "@bilig/core";
import { schema } from "@bilig/zero-sync";
import {
  type WorkbookChangeUndoBundle,
  type AuthoritativeWorkbookEventBatch,
  queries,
  workbookCellArgsSchema,
  workbookColumnTileArgsSchema,
  workbookQueryArgsSchema,
  workbookRowTileArgsSchema,
  workbookTileArgsSchema,
} from "@bilig/zero-sync";
import {
  applyWorkbookAgentCommandBundle,
  areWorkbookAgentPreviewSummariesEqual,
  buildWorkbookAgentPreview,
  type WorkbookAgentPreviewSummary,
} from "@bilig/agent-api";
import type { EngineOp } from "@bilig/workbook-domain";
import type { FastifyRequest } from "fastify";
import type { SessionIdentity } from "../http/session.js";
import { resolveSessionIdentity } from "../http/session.js";
import {
  WorkbookRuntimeManager,
  type WorkbookRuntime,
} from "../workbook-runtime/runtime-manager.js";
import { createZeroDbProvider, createZeroPool, resolveZeroDatabaseUrl } from "./db.js";
import { handleServerMutator } from "./server-mutators.js";
import { ZeroRecalcWorker } from "./recalc-worker.js";
import {
  ensureZeroSyncSchema,
  loadWorkbookEventRecordsAfter,
  persistWorkbookMutation,
} from "./store.js";
import {
  backfillAuthoritativeCellEval,
  dropLegacyZeroSyncSchemaObjects,
} from "./workbook-migration-store.js";
import {
  acquireWorkbookMutationLock,
  loadWorkbookRuntimeMetadata,
} from "./workbook-runtime-store.js";
import { ensureWorkbookPresenceSchema } from "./presence-store.js";
import { ensureZeroPublication } from "./publication-store.js";
import { backfillWorkbookChanges, ensureWorkbookChangeSchema } from "./workbook-change-store.js";
import {
  appendWorkbookAgentRun,
  ensureWorkbookAgentRunSchema,
  listWorkbookAgentRuns,
} from "./workbook-agent-run-store.js";
import type { WorkbookAgentCommandBundle, WorkbookAgentExecutionRecord } from "@bilig/agent-api";
import { createWorkbookAgentServiceError } from "../workbook-agent-errors.js";

export interface ZeroSyncService {
  readonly enabled: boolean;
  initialize(): Promise<void>;
  close(): Promise<void>;
  handleQuery(request: FastifyRequest): Promise<unknown>;
  handleMutate(request: FastifyRequest): Promise<unknown>;
  inspectWorkbook<T>(
    documentId: string,
    task: (runtime: WorkbookRuntime) => Promise<T> | T,
  ): Promise<T>;
  applyServerMutator(name: string, args: unknown, session?: SessionIdentity): Promise<void>;
  applyAgentCommandBundle(
    documentId: string,
    bundle: WorkbookAgentCommandBundle,
    preview: WorkbookAgentPreviewSummary,
    session?: SessionIdentity,
  ): Promise<{ revision: number; preview: WorkbookAgentPreviewSummary }>;
  listWorkbookAgentRuns(
    documentId: string,
    actorUserId: string,
    limit?: number,
  ): Promise<WorkbookAgentExecutionRecord[]>;
  appendWorkbookAgentRun(record: WorkbookAgentExecutionRecord): Promise<void>;
  getWorkbookHeadRevision(documentId: string): Promise<number>;
  loadAuthoritativeEvents(
    documentId: string,
    afterRevision: number,
  ): Promise<AuthoritativeWorkbookEventBatch>;
}

function fastifyRequestToWebRequest(request: FastifyRequest): Request {
  const origin =
    typeof request.headers.host === "string"
      ? `http://${request.headers.host}`
      : "http://localhost";
  const headers = new Headers();
  for (const [key, value] of Object.entries(request.headers)) {
    if (Array.isArray(value)) {
      for (const entry of value) {
        headers.append(key, entry);
      }
      continue;
    }
    if (typeof value === "string") {
      headers.set(key, value);
    }
  }
  const body =
    request.body === undefined || request.body === null ? undefined : JSON.stringify(request.body);

  const init: RequestInit = {
    method: request.method,
    headers,
  };
  if (body !== undefined) {
    init.body = body;
  }

  return new Request(new URL(request.url, origin), init);
}

class DisabledZeroSyncService implements ZeroSyncService {
  readonly enabled = false;

  async initialize(): Promise<void> {}

  async close(): Promise<void> {}

  async handleQuery(): Promise<never> {
    throw new Error("Zero sync is not configured");
  }

  async handleMutate(): Promise<never> {
    throw new Error("Zero sync is not configured");
  }

  async inspectWorkbook<T>(
    _documentId: string,
    _task: (runtime: WorkbookRuntime) => Promise<T> | T,
  ): Promise<T> {
    throw new Error("Zero sync is not configured");
  }

  async applyServerMutator(
    _name: string,
    _args: unknown,
    _session?: SessionIdentity,
  ): Promise<void> {
    throw new Error("Zero sync is not configured");
  }

  async applyAgentCommandBundle(): Promise<never> {
    throw new Error("Zero sync is not configured");
  }

  async listWorkbookAgentRuns(): Promise<never> {
    throw new Error("Zero sync is not configured");
  }

  async appendWorkbookAgentRun(): Promise<never> {
    throw new Error("Zero sync is not configured");
  }

  async getWorkbookHeadRevision(): Promise<never> {
    throw new Error("Zero sync is not configured");
  }

  async loadAuthoritativeEvents(_documentId: string, _afterRevision: number): Promise<never> {
    throw new Error("Zero sync is not configured");
  }
}

class EnabledZeroSyncService implements ZeroSyncService {
  readonly enabled = true;
  private readonly pool: ReturnType<typeof createZeroPool>;
  private readonly dbProvider;
  private readonly runtimeManager: WorkbookRuntimeManager;
  private readonly recalcWorker: ZeroRecalcWorker;

  constructor(connectionString: string) {
    this.pool = createZeroPool(connectionString);
    this.dbProvider = createZeroDbProvider(connectionString);
    this.runtimeManager = new WorkbookRuntimeManager();
    this.recalcWorker = new ZeroRecalcWorker(this.pool, this.runtimeManager);
  }

  async initialize(): Promise<void> {
    await ensureZeroSyncSchema(this.pool);
    await ensureWorkbookPresenceSchema(this.pool);
    await ensureWorkbookChangeSchema(this.pool);
    await ensureWorkbookAgentRunSchema(this.pool);
    await ensureZeroPublication(this.pool);
    await backfillAuthoritativeCellEval(this.pool);
    await backfillWorkbookChanges(this.pool);
    await dropLegacyZeroSyncSchemaObjects(this.pool);
    this.recalcWorker.start();
  }

  async close(): Promise<void> {
    this.recalcWorker.stop();
    await this.runtimeManager.close();
    await this.pool.end();
  }

  async handleQuery(request: FastifyRequest): Promise<unknown> {
    const session = resolveSessionIdentity(request);
    const queryLookup = {
      "workbook.get": {
        query: queries.workbook.get,
        schema: workbookQueryArgsSchema,
      },
      "sheet.byWorkbook": {
        query: queries.sheet.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "cellInput.one": {
        query: queries.cellInput.one,
        schema: workbookCellArgsSchema,
      },
      "cellInput.tile": {
        query: queries.cellInput.tile,
        schema: workbookTileArgsSchema,
      },
      "cellEval.one": {
        query: queries.cellEval.one,
        schema: workbookCellArgsSchema,
      },
      "cellEval.tile": {
        query: queries.cellEval.tile,
        schema: workbookTileArgsSchema,
      },
      "cellRender.tile": {
        query: queries.cellRender.tile,
        schema: workbookTileArgsSchema,
      },
      "sheetRow.tile": {
        query: queries.sheetRow.tile,
        schema: workbookRowTileArgsSchema,
      },
      "sheetCol.tile": {
        query: queries.sheetCol.tile,
        schema: workbookColumnTileArgsSchema,
      },
      "cellStyle.byWorkbook": {
        query: queries.cellStyle.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "numberFormat.byWorkbook": {
        query: queries.numberFormat.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "presenceCoarse.byWorkbook": {
        query: queries.presenceCoarse.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "presence.byWorkbook": {
        query: queries.presence.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "workbookChange.byWorkbook": {
        query: queries.workbookChange.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "workbookChanges.byWorkbook": {
        query: queries.workbookChanges.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
    } as const;

    return await handleQueryRequest(
      (name, args) => {
        if (!hasOwn(queryLookup, name)) {
          throw new Error(`Unknown Zero query: ${name}`);
        }
        const query = queryLookup[name];
        return query.query.fn({
          // Zero's query registry erases the specific arg type when accessed through the name map.
          // oxlint-disable-next-line @typescript-eslint/no-unsafe-type-assertion
          args: query.schema.parse(args) as never,
          ctx: { userID: session.userID },
        });
      },
      schema,
      fastifyRequestToWebRequest(request),
    );
  }

  async handleMutate(request: FastifyRequest): Promise<unknown> {
    const session = resolveSessionIdentity(request);
    return await handleMutateRequest(
      this.dbProvider,
      (transact) =>
        transact(async (tx, name, args) => {
          return await handleServerMutator(tx, name, args, this.runtimeManager, session);
        }),
      fastifyRequestToWebRequest(request),
    );
  }

  async inspectWorkbook<T>(
    documentId: string,
    task: (runtime: WorkbookRuntime) => Promise<T> | T,
  ): Promise<T> {
    return await this.runtimeManager.runExclusive(documentId, async () => {
      const runtime = await this.runtimeManager.loadRuntime(this.pool, documentId);
      return await task(runtime);
    });
  }

  async applyServerMutator(name: string, args: unknown, session?: SessionIdentity): Promise<void> {
    const client = await this.pool.connect();
    try {
      await client.query("BEGIN");
      await handleServerMutator(
        {
          dbTransaction: {
            wrappedTransaction: client,
          },
        },
        name,
        args,
        this.runtimeManager,
        session,
      );
      await client.query("COMMIT");
    } catch (error) {
      await client.query("ROLLBACK").catch(() => undefined);
      throw error;
    } finally {
      client.release();
    }
  }

  async applyAgentCommandBundle(
    documentId: string,
    bundle: WorkbookAgentCommandBundle,
    preview: WorkbookAgentPreviewSummary,
    session?: SessionIdentity,
  ): Promise<{ revision: number; preview: WorkbookAgentPreviewSummary }> {
    const client = await this.pool.connect();
    try {
      return await this.runtimeManager.runExclusive(documentId, async () => {
        await client.query("BEGIN");
        try {
          await acquireWorkbookMutationLock(client, documentId);
          const state = await this.runtimeManager.loadRuntime(client, documentId);
          if (state.headRevision !== bundle.baseRevision) {
            throw createWorkbookAgentServiceError({
              code: "WORKBOOK_AGENT_PREVIEW_STALE",
              message:
                "Workbook changed after preview. Replay the plan to stage a fresh preview bundle.",
              statusCode: 409,
              retryable: true,
            });
          }
          const authoritativePreview = await buildWorkbookAgentPreview({
            snapshot: state.engine.exportSnapshot(),
            replicaId: `server:${documentId}:agent-preview:r${String(state.headRevision)}`,
            bundle,
          });
          if (!areWorkbookAgentPreviewSummariesEqual(preview, authoritativePreview)) {
            throw createWorkbookAgentServiceError({
              code: "WORKBOOK_AGENT_PREVIEW_MISMATCH",
              message:
                "Local preview no longer matches the authoritative workbook state. Replay the plan to refresh the preview.",
              statusCode: 409,
              retryable: true,
            });
          }
          const undoBundle = captureAgentUndoBundle(state.engine, bundle);
          const ownerUserId = resolveOwnerUserId(state, session);
          const result = await persistWorkbookMutation(client, documentId, {
            previousState: state,
            nextEngine: state.engine,
            updatedBy: session?.userID ?? "system",
            ownerUserId,
            eventPayload: {
              kind: "applyAgentCommandBundle",
              bundle,
            },
            undoBundle,
          });
          this.runtimeManager.commitMutation(documentId, {
            projectionCommit: result.projectionCommit,
            headRevision: result.revision,
            calculatedRevision: result.calculatedRevision,
            ownerUserId,
          });
          await client.query("COMMIT");
          return {
            revision: result.revision,
            preview: authoritativePreview,
          };
        } catch (error) {
          this.runtimeManager.invalidate(documentId);
          await client.query("ROLLBACK").catch(() => undefined);
          throw error;
        }
      });
    } finally {
      client.release();
    }
  }

  async listWorkbookAgentRuns(
    documentId: string,
    actorUserId: string,
    limit = 20,
  ): Promise<WorkbookAgentExecutionRecord[]> {
    return await listWorkbookAgentRuns(this.pool, {
      documentId,
      actorUserId,
      limit,
    });
  }

  async appendWorkbookAgentRun(record: WorkbookAgentExecutionRecord): Promise<void> {
    await appendWorkbookAgentRun(this.pool, record);
  }

  async getWorkbookHeadRevision(documentId: string): Promise<number> {
    const metadata = await loadWorkbookRuntimeMetadata(this.pool, documentId);
    return metadata.headRevision;
  }

  async loadAuthoritativeEvents(
    documentId: string,
    afterRevision: number,
  ): Promise<AuthoritativeWorkbookEventBatch> {
    const metadata = await loadWorkbookRuntimeMetadata(this.pool, documentId);
    const events =
      metadata.headRevision > afterRevision
        ? await loadWorkbookEventRecordsAfter(this.pool, documentId, afterRevision)
        : [];
    return {
      afterRevision,
      headRevision: metadata.headRevision,
      calculatedRevision: metadata.calculatedRevision,
      events,
    };
  }
}

function resolveOwnerUserId(state: { ownerUserId: string }, session?: SessionIdentity): string {
  if (state.ownerUserId !== "system" || !session?.userID) {
    return state.ownerUserId;
  }
  return session.userID;
}

function toEngineUndoBundle(undoOps: readonly EngineOp[] | null): WorkbookChangeUndoBundle | null {
  if (!undoOps || undoOps.length === 0) {
    return null;
  }
  return {
    kind: "engineOps",
    ops: structuredClone([...undoOps]),
  };
}

function captureAgentUndoBundle(
  engine: SpreadsheetEngine,
  bundle: WorkbookAgentCommandBundle,
): WorkbookChangeUndoBundle | null {
  return toEngineUndoBundle(
    engine.captureUndoOps(() => {
      applyWorkbookAgentCommandBundle(engine, bundle);
    }).undoOps,
  );
}

function hasOwn<ObjectType extends object>(
  object: ObjectType,
  key: PropertyKey,
): key is keyof ObjectType {
  return Object.prototype.hasOwnProperty.call(object, key);
}

export function createZeroSyncService(): ZeroSyncService {
  const connectionString = resolveZeroDatabaseUrl();
  if (!connectionString) {
    return new DisabledZeroSyncService();
  }
  return new EnabledZeroSyncService(connectionString);
}
