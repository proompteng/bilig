import { handleMutateRequest, handleQueryRequest } from "@rocicorp/zero/server";
import { schema } from "@bilig/zero-sync";
import {
  type AuthoritativeWorkbookEventBatch,
  queries,
  workbookCellArgsSchema,
  workbookColumnTileArgsSchema,
  workbookQueryArgsSchema,
  workbookRowTileArgsSchema,
  workbookTileArgsSchema,
} from "@bilig/zero-sync";
import type { FastifyRequest } from "fastify";
import { resolveSessionIdentity } from "../http/session.js";
import { WorkbookRuntimeManager } from "../workbook-runtime/runtime-manager.js";
import { createZeroDbProvider, createZeroPool, resolveZeroDatabaseUrl } from "./db.js";
import { handleServerMutator } from "./server-mutators.js";
import { ZeroRecalcWorker } from "./recalc-worker.js";
import {
  backfillAuthoritativeCellEval,
  dropLegacyZeroSyncSchemaObjects,
  ensureZeroSyncSchema,
  loadWorkbookEventRecordsAfter,
  loadWorkbookRuntimeMetadata,
} from "./store.js";
import { ensureWorkbookPresenceSchema } from "./presence-store.js";
import { backfillWorkbookChanges, ensureWorkbookChangeSchema } from "./workbook-change-store.js";
import { ensureWorkbookSheetViewSchema } from "./sheet-view-store.js";

export interface ZeroSyncService {
  readonly enabled: boolean;
  initialize(): Promise<void>;
  close(): Promise<void>;
  handleQuery(request: FastifyRequest): Promise<unknown>;
  handleMutate(request: FastifyRequest): Promise<unknown>;
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
    await ensureWorkbookSheetViewSchema(this.pool);
    await ensureWorkbookChangeSchema(this.pool);
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
      "sheetView.byWorkbook": {
        query: queries.sheetView.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "sheetViews.byWorkbook": {
        query: queries.sheetViews.byWorkbook,
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
