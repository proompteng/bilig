import { zeroNodePg } from "@rocicorp/zero/server/adapters/pg";
import { Pool } from "pg";
import { schema } from "@bilig/zero-sync";

export function resolveZeroDatabaseUrl(): string | null {
  return (
    process.env["ZERO_UPSTREAM_DB"] ??
    process.env["DATABASE_URL"] ??
    process.env["BILIG_DATABASE_URL"] ??
    null
  );
}

export function createZeroPool(connectionString: string): Pool {
  const pool = new Pool({
    connectionString,
  });
  pool.on("error", (error) => {
    console.error("Zero Postgres pool error", error);
  });
  return pool;
}

export function createZeroDbProvider(connectionString: string) {
  return zeroNodePg(schema, connectionString);
}

export type BiligDbProvider = ReturnType<typeof createZeroDbProvider>;

declare module "@rocicorp/zero" {
  interface DefaultTypes {
    dbProvider: BiligDbProvider;
  }
}
