import { handleMutateRequest, handleQueryRequest } from "@rocicorp/zero/server";
import { schema } from "@bilig/zero-sync";
import { queries, workbookQueryArgsSchema } from "@bilig/zero-sync";
import type { FastifyRequest } from "fastify";
import type { Pool } from "pg";
import { createZeroDbProvider, createZeroPool, resolveZeroDatabaseUrl } from "./db.js";
import { handleServerMutator } from "./server-mutators.js";
import { ensureZeroSyncSchema } from "./store.js";

export interface ZeroSyncService {
  readonly enabled: boolean;
  readonly pool: Pool | null;
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
  readonly pool = null;

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
  readonly pool: Pool;
  private readonly dbProvider;

  constructor(connectionString: string) {
    this.pool = createZeroPool(connectionString);
    this.dbProvider = createZeroDbProvider(connectionString);
  }

  async initialize(): Promise<void> {
    await ensureZeroSyncSchema(this.pool);
  }

  async close(): Promise<void> {
    await this.pool.end();
  }

  async handleQuery(request: FastifyRequest): Promise<unknown> {
    const queryLookup = {
      "workbooks.byId": queries.workbooks.byId,
    } as const;

    return await handleQueryRequest(
      (name, args) => {
        if (!hasOwn(queryLookup, name)) {
          throw new Error(`Unknown Zero query: ${name}`);
        }
        const query = queryLookup[name];
        return query.fn({ args: workbookQueryArgsSchema.parse(args), ctx: {} });
      },
      schema,
      fastifyRequestToWebRequest(request),
    );
  }

  async handleMutate(request: FastifyRequest): Promise<unknown> {
    return await handleMutateRequest(
      this.dbProvider,
      (transact) =>
        transact(async (tx, name, args) => {
          return await handleServerMutator(tx, name, args);
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
