import { handleMutateRequest, handleQueryRequest } from "@rocicorp/zero/server";
import { schema } from "@bilig/zero-sync";
import {
  queries,
  workbookCellArgsSchema,
  workbookColumnTileArgsSchema,
  workbookQueryArgsSchema,
  workbookRowTileArgsSchema,
  workbookTileArgsSchema,
} from "@bilig/zero-sync";
import type { FastifyRequest } from "fastify";
import { resolveSessionIdentity } from "../session.js";
import { createZeroDbProvider, createZeroPool, resolveZeroDatabaseUrl } from "./db.js";
import { handleServerMutator } from "./server-mutators.js";
import { ZeroRecalcWorker } from "./recalc-worker.js";
import { WorkbookRuntimeManager } from "./runtime-manager.js";
import { ensureZeroSyncSchema } from "./store.js";

export interface ZeroSyncService {
  readonly enabled: boolean;
  initialize(): Promise<void>;
  close(): Promise<void>;
  handleQuery(request: FastifyRequest): Promise<unknown>;
  handleMutate(request: FastifyRequest): Promise<unknown>;
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
      "workbooks.get": {
        query: queries.workbooks.get,
        schema: workbookQueryArgsSchema,
      },
      "sheets.byWorkbook": {
        query: queries.sheets.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "cells.one": {
        query: queries.cells.one,
        schema: workbookCellArgsSchema,
      },
      "cells.tile": {
        query: queries.cells.tile,
        schema: workbookTileArgsSchema,
      },
      "cellEval.tile": {
        query: queries.cellEval.tile,
        schema: workbookTileArgsSchema,
      },
      "computedCells.tile": {
        query: queries.cellEval.tile,
        schema: workbookTileArgsSchema,
      },
      "rowMetadata.tile": {
        query: queries.rowMetadata.tile,
        schema: workbookRowTileArgsSchema,
      },
      "columnMetadata.tile": {
        query: queries.columnMetadata.tile,
        schema: workbookColumnTileArgsSchema,
      },
      "styleRanges.intersectTile": {
        query: queries.styleRanges.intersectTile,
        schema: workbookTileArgsSchema,
      },
      "formatRanges.intersectTile": {
        query: queries.formatRanges.intersectTile,
        schema: workbookTileArgsSchema,
      },
      "styles.byWorkbook": {
        query: queries.styles.byWorkbook,
        schema: workbookQueryArgsSchema,
      },
      "numberFormats.byWorkbook": {
        query: queries.numberFormats.byWorkbook,
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
