import sqlite3InitModule, {
  type Database,
  type SAHPoolUtil,
  type SqlValue,
  type Sqlite3Static,
} from "@sqlite.org/sqlite-wasm";
import type {
  WorkbookLocalAuthoritativeDelta,
  WorkbookLocalAuthoritativeBase,
  WorkbookLocalProjectionOverlay,
  WorkbookLocalViewportBase,
} from "./workbook-local-base.js";
import {
  readWorkbookViewportProjection,
  writeWorkbookAuthoritativeBase,
  writeWorkbookAuthoritativeDelta,
  writeWorkbookProjectionOverlay,
} from "./workbook-local-store-projection.js";
import { initializeWorkbookLocalStoreSchema } from "./workbook-local-store-schema.js";

const WORKBOOK_VFS_NAME = "bilig-opfs-sahpool";
const WORKBOOK_VFS_DIRECTORY = "/bilig/workbooks";
const WORKBOOK_VFS_INITIAL_CAPACITY = 12;

let sqliteRuntimePromise: Promise<{ sqlite3: Sqlite3Static; poolUtil: SAHPoolUtil }> | null = null;
let memorySqliteRuntimePromise: Promise<Sqlite3Static> | null = null;

export class WorkbookLocalStoreLockedError extends Error {
  override readonly name = "WorkbookLocalStoreLockedError";
}

export interface WorkbookStoredState {
  readonly snapshot: unknown;
  readonly replica: unknown;
  readonly authoritativeRevision: number;
  readonly appliedPendingLocalSeq: number;
}

export interface WorkbookBootstrapState {
  readonly workbookName: string;
  readonly sheetNames: readonly string[];
  readonly materializedCellCount: number;
  readonly authoritativeRevision: number;
  readonly appliedPendingLocalSeq: number;
}

export interface WorkbookLocalMutationRecord {
  readonly id: string;
  readonly localSeq: number;
  readonly baseRevision: number;
  readonly method: string;
  readonly args: unknown[];
  readonly enqueuedAtUnixMs: number;
  readonly submittedAtUnixMs: number | null;
  readonly status: "pending" | "submitted";
}

export interface WorkbookLocalStore {
  loadBootstrapState(): Promise<WorkbookBootstrapState | null>;
  loadState(): Promise<WorkbookStoredState | null>;
  persistProjectionState(input: {
    readonly state: WorkbookStoredState;
    readonly authoritativeBase: WorkbookLocalAuthoritativeBase;
    readonly projectionOverlay: WorkbookLocalProjectionOverlay;
  }): Promise<void>;
  ingestAuthoritativeDelta(input: {
    readonly state: WorkbookStoredState;
    readonly authoritativeDelta: WorkbookLocalAuthoritativeDelta;
    readonly projectionOverlay: WorkbookLocalProjectionOverlay;
    readonly removePendingMutationIds?: readonly string[];
  }): Promise<void>;
  listPendingMutations(): Promise<WorkbookLocalMutationRecord[]>;
  appendPendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void>;
  updatePendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void>;
  removePendingMutation(id: string): Promise<void>;
  readViewportProjection(
    sheetName: string,
    viewport: {
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    },
  ): WorkbookLocalViewportBase | null;
  close(): void;
}

export interface WorkbookLocalStoreFactory {
  open(documentId: string): Promise<WorkbookLocalStore>;
}

export interface OpfsWorkbookLocalStoreFactoryOptions {
  vfsName?: string;
  directory?: string;
  initialCapacity?: number;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function supportsWorkerOpfs(): boolean {
  const scope = globalThis as typeof globalThis & {
    navigator?: Navigator;
    document?: Document;
  };
  if (typeof scope.document !== "undefined") {
    return false;
  }
  return typeof scope.navigator?.storage?.getDirectory === "function";
}

function sanitizeDocumentId(documentId: string): string {
  return encodeURIComponent(documentId).replaceAll("%", "_");
}

function isAccessHandleConflict(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }
  return (
    (error.message.includes("createSyncAccessHandle") &&
      error.message.includes("Access Handles cannot be created")) ||
    (error.name === "NoModificationAllowedError" &&
      error.message.includes("Access Handles cannot be created"))
  );
}

function toWorkbookLocalStoreLockedError(
  documentId: string | null,
  cause?: unknown,
): WorkbookLocalStoreLockedError {
  const suffix = documentId ? ` for ${documentId}` : "";
  const error = new WorkbookLocalStoreLockedError(
    `Workbook local store is locked by another tab${suffix}`,
  );
  if (cause !== undefined) {
    Object.defineProperty(error, "cause", {
      configurable: true,
      enumerable: false,
      value: cause,
      writable: true,
    });
  }
  return error;
}

function parseWorkbookStoredState(value: unknown): WorkbookStoredState | null {
  if (
    !isRecord(value) ||
    typeof value["authoritativeRevision"] !== "number" ||
    typeof value["appliedPendingLocalSeq"] !== "number" ||
    !isRecord(value["snapshot"]) ||
    !isRecord(value["replica"])
  ) {
    return null;
  }
  return {
    snapshot: value["snapshot"],
    replica: value["replica"],
    authoritativeRevision: value["authoritativeRevision"],
    appliedPendingLocalSeq: value["appliedPendingLocalSeq"],
  };
}

function parseWorkbookBootstrapState(value: unknown): WorkbookBootstrapState | null {
  if (
    !isRecord(value) ||
    typeof value["workbookName"] !== "string" ||
    !Array.isArray(value["sheetNames"]) ||
    !value["sheetNames"].every((sheetName) => typeof sheetName === "string") ||
    typeof value["materializedCellCount"] !== "number" ||
    typeof value["authoritativeRevision"] !== "number" ||
    typeof value["appliedPendingLocalSeq"] !== "number"
  ) {
    return null;
  }
  return {
    workbookName: value["workbookName"],
    sheetNames: [...value["sheetNames"]],
    materializedCellCount: value["materializedCellCount"],
    authoritativeRevision: value["authoritativeRevision"],
    appliedPendingLocalSeq: value["appliedPendingLocalSeq"],
  };
}

function parseWorkbookLocalMutationRecord(value: unknown): WorkbookLocalMutationRecord | null {
  if (
    !isRecord(value) ||
    typeof value["id"] !== "string" ||
    typeof value["localSeq"] !== "number" ||
    typeof value["baseRevision"] !== "number" ||
    typeof value["method"] !== "string" ||
    !Array.isArray(value["args"]) ||
    typeof value["enqueuedAtUnixMs"] !== "number" ||
    (value["submittedAtUnixMs"] !== null && typeof value["submittedAtUnixMs"] !== "number") ||
    (value["status"] !== "pending" && value["status"] !== "submitted")
  ) {
    return null;
  }
  return {
    id: value["id"],
    localSeq: value["localSeq"],
    baseRevision: value["baseRevision"],
    method: value["method"],
    args: [...value["args"]],
    enqueuedAtUnixMs: value["enqueuedAtUnixMs"],
    submittedAtUnixMs: value["submittedAtUnixMs"] ?? null,
    status: value["status"],
  };
}

function readSingleObjectRow(
  db: Database,
  sql: string,
  bind?: readonly SqlValue[],
): Record<string, SqlValue> | null {
  const statement = db.prepare(sql);
  try {
    if (bind) {
      statement.bind([...bind]);
    }
    if (!statement.step()) {
      return null;
    }
    return statement.get({});
  } finally {
    statement.finalize();
  }
}

async function getSqliteRuntime(
  options: Required<OpfsWorkbookLocalStoreFactoryOptions>,
): Promise<{ sqlite3: Sqlite3Static; poolUtil: SAHPoolUtil }> {
  if (!supportsWorkerOpfs()) {
    throw new Error("Workbook local storage requires a worker with OPFS support");
  }
  if (!sqliteRuntimePromise) {
    sqliteRuntimePromise = (async () => {
      const sqlite3 = await sqlite3InitModule();
      try {
        const poolUtil = await sqlite3.installOpfsSAHPoolVfs({
          name: options.vfsName,
          directory: options.directory,
          initialCapacity: options.initialCapacity,
        });
        return { sqlite3, poolUtil };
      } catch (error) {
        if (isAccessHandleConflict(error)) {
          throw toWorkbookLocalStoreLockedError(null, error);
        }
        throw error;
      }
    })();
  }
  try {
    return await sqliteRuntimePromise;
  } catch (error) {
    if (!(error instanceof WorkbookLocalStoreLockedError)) {
      sqliteRuntimePromise = null;
    }
    throw error;
  }
}

function extractWorkbookName(snapshot: unknown): string | null {
  if (
    isRecord(snapshot) &&
    isRecord(snapshot["workbook"]) &&
    typeof snapshot["workbook"]["name"] === "string"
  ) {
    return snapshot["workbook"]["name"];
  }
  return null;
}

class SqliteWorkbookLocalStore implements WorkbookLocalStore {
  constructor(
    private readonly db: Database,
    private readonly defaultWorkbookName: string,
    private readonly closeDbOnClose = true,
  ) {}

  async loadBootstrapState(): Promise<WorkbookBootstrapState | null> {
    const row = readSingleObjectRow(
      this.db,
      `
        SELECT workbook_name AS workbookName,
               (SELECT COUNT(*) FROM authoritative_cell_render) AS materializedCellCount,
               authoritative_revision AS authoritativeRevision,
               applied_pending_local_seq AS appliedPendingLocalSeq
          FROM runtime_state
         WHERE id = 1
      `,
    );
    if (!row) {
      return null;
    }

    const sheetNames: string[] = [];
    const statement = this.db.prepare(
      `
        SELECT name
          FROM authoritative_sheet
         ORDER BY sort_order ASC, name ASC
      `,
    );
    try {
      while (statement.step()) {
        const sheet = statement.get({});
        if (typeof sheet["name"] === "string") {
          sheetNames.push(sheet["name"]);
        }
      }
    } finally {
      statement.finalize();
    }

    return parseWorkbookBootstrapState({
      workbookName:
        typeof row["workbookName"] === "string" && row["workbookName"].length > 0
          ? row["workbookName"]
          : this.defaultWorkbookName,
      sheetNames,
      materializedCellCount:
        typeof row["materializedCellCount"] === "number" ? row["materializedCellCount"] : 0,
      authoritativeRevision: row["authoritativeRevision"],
      appliedPendingLocalSeq: row["appliedPendingLocalSeq"],
    });
  }

  async loadState(): Promise<WorkbookStoredState | null> {
    const row = readSingleObjectRow(
      this.db,
      `
        SELECT snapshot_json AS snapshotJson,
               replica_json AS replicaJson,
               authoritative_revision AS authoritativeRevision,
               applied_pending_local_seq AS appliedPendingLocalSeq
          FROM runtime_state
         WHERE id = 1
      `,
    );
    if (!row) {
      return null;
    }
    const snapshotJson = row["snapshotJson"];
    const replicaJson = row["replicaJson"];
    const authoritativeRevision = row["authoritativeRevision"];
    const appliedPendingLocalSeq = row["appliedPendingLocalSeq"];
    if (
      typeof snapshotJson !== "string" ||
      typeof replicaJson !== "string" ||
      typeof authoritativeRevision !== "number" ||
      typeof appliedPendingLocalSeq !== "number"
    ) {
      return null;
    }
    try {
      return parseWorkbookStoredState({
        snapshot: JSON.parse(snapshotJson) as unknown,
        replica: JSON.parse(replicaJson) as unknown,
        authoritativeRevision,
        appliedPendingLocalSeq,
      });
    } catch {
      return null;
    }
  }

  async persistProjectionState(input: {
    readonly state: WorkbookStoredState;
    readonly authoritativeBase: WorkbookLocalAuthoritativeBase;
    readonly projectionOverlay: WorkbookLocalProjectionOverlay;
  }): Promise<void> {
    this.db.transaction((db) => {
      writeWorkbookAuthoritativeBase(db, input.authoritativeBase);
      writeWorkbookProjectionOverlay(db, input.projectionOverlay);
      db.exec(
        `
          INSERT INTO runtime_state (
            id,
            workbook_name,
            snapshot_json,
            replica_json,
            authoritative_revision,
            applied_pending_local_seq,
            updated_at_ms
          )
          VALUES (1, ?, ?, ?, ?, ?, ?)
          ON CONFLICT(id) DO UPDATE SET
            workbook_name = excluded.workbook_name,
            snapshot_json = excluded.snapshot_json,
            replica_json = excluded.replica_json,
            authoritative_revision = excluded.authoritative_revision,
            applied_pending_local_seq = excluded.applied_pending_local_seq,
            updated_at_ms = excluded.updated_at_ms
        `,
        {
          bind: [
            extractWorkbookName(input.state.snapshot) ?? this.defaultWorkbookName,
            JSON.stringify(input.state.snapshot),
            JSON.stringify(input.state.replica),
            input.state.authoritativeRevision,
            input.state.appliedPendingLocalSeq,
            Date.now(),
          ],
        },
      );
    });
  }

  async ingestAuthoritativeDelta(input: {
    readonly state: WorkbookStoredState;
    readonly authoritativeDelta: WorkbookLocalAuthoritativeDelta;
    readonly projectionOverlay: WorkbookLocalProjectionOverlay;
    readonly removePendingMutationIds?: readonly string[];
  }): Promise<void> {
    this.db.transaction((db) => {
      if ((input.removePendingMutationIds?.length ?? 0) > 0) {
        const deletePendingMutation = db.prepare("DELETE FROM pending_op WHERE op_id = ?");
        try {
          input.removePendingMutationIds?.forEach((id) => {
            deletePendingMutation.bind([id]);
            deletePendingMutation.step();
            deletePendingMutation.reset();
          });
        } finally {
          deletePendingMutation.finalize();
        }
      }
      writeWorkbookAuthoritativeDelta(db, input.authoritativeDelta);
      writeWorkbookProjectionOverlay(db, input.projectionOverlay);
      db.exec(
        `
          INSERT INTO runtime_state (
            id,
            workbook_name,
            snapshot_json,
            replica_json,
            authoritative_revision,
            applied_pending_local_seq,
            updated_at_ms
          )
          VALUES (1, ?, ?, ?, ?, ?, ?)
          ON CONFLICT(id) DO UPDATE SET
            workbook_name = excluded.workbook_name,
            snapshot_json = excluded.snapshot_json,
            replica_json = excluded.replica_json,
            authoritative_revision = excluded.authoritative_revision,
            applied_pending_local_seq = excluded.applied_pending_local_seq,
            updated_at_ms = excluded.updated_at_ms
        `,
        {
          bind: [
            extractWorkbookName(input.state.snapshot) ?? this.defaultWorkbookName,
            JSON.stringify(input.state.snapshot),
            JSON.stringify(input.state.replica),
            input.state.authoritativeRevision,
            input.state.appliedPendingLocalSeq,
            Date.now(),
          ],
        },
      );
    });
  }

  async listPendingMutations(): Promise<WorkbookLocalMutationRecord[]> {
    const rows: Record<string, SqlValue>[] = [];
    const statement = this.db.prepare(
      `
        SELECT op_id AS id,
               local_seq AS localSeq,
               base_revision AS baseRevision,
               method,
               args_json AS argsJson,
               enqueued_at_ms AS enqueuedAtUnixMs,
               submitted_at_ms AS submittedAtUnixMs,
               status
          FROM pending_op
         ORDER BY local_seq ASC
      `,
    );
    try {
      while (statement.step()) {
        rows.push(statement.get({}));
      }
    } finally {
      statement.finalize();
    }
    return rows.flatMap((row) => {
      const argsJson = row["argsJson"];
      if (typeof argsJson !== "string") {
        return [];
      }
      try {
        const parsed = parseWorkbookLocalMutationRecord({
          ...row,
          args: JSON.parse(argsJson) as unknown,
        });
        return parsed ? [parsed] : [];
      } catch {
        return [];
      }
    });
  }

  async appendPendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void> {
    this.db.transaction((db) => {
      db.exec(
        `
          INSERT INTO pending_op (
            op_id,
            local_seq,
            base_revision,
            method,
            args_json,
            enqueued_at_ms,
            submitted_at_ms,
            status
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        `,
        {
          bind: [
            mutation.id,
            mutation.localSeq,
            mutation.baseRevision,
            mutation.method,
            JSON.stringify(mutation.args),
            mutation.enqueuedAtUnixMs,
            mutation.submittedAtUnixMs,
            mutation.status,
          ],
        },
      );
    });
  }

  async updatePendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void> {
    this.db.exec(
      `
        UPDATE pending_op
           SET base_revision = ?,
               method = ?,
               args_json = ?,
               enqueued_at_ms = ?,
               submitted_at_ms = ?,
               status = ?
         WHERE op_id = ?
      `,
      {
        bind: [
          mutation.baseRevision,
          mutation.method,
          JSON.stringify(mutation.args),
          mutation.enqueuedAtUnixMs,
          mutation.submittedAtUnixMs,
          mutation.status,
          mutation.id,
        ],
      },
    );
  }

  async removePendingMutation(id: string): Promise<void> {
    this.db.exec("DELETE FROM pending_op WHERE op_id = ?", {
      bind: [id],
    });
  }

  readViewportProjection(
    sheetName: string,
    viewport: {
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    },
  ): WorkbookLocalViewportBase | null {
    return readWorkbookViewportProjection(this.db, sheetName, viewport);
  }

  close(): void {
    if (this.closeDbOnClose) {
      this.db.close();
    }
  }
}

async function getMemorySqliteRuntime(): Promise<Sqlite3Static> {
  if (!memorySqliteRuntimePromise) {
    memorySqliteRuntimePromise = (async () => {
      try {
        return await sqlite3InitModule();
      } catch (error) {
        memorySqliteRuntimePromise = null;
        throw error;
      }
    })();
  }
  return await memorySqliteRuntimePromise;
}

export function createOpfsWorkbookLocalStoreFactory(
  options: OpfsWorkbookLocalStoreFactoryOptions = {},
): WorkbookLocalStoreFactory {
  const resolvedOptions: Required<OpfsWorkbookLocalStoreFactoryOptions> = {
    vfsName: options.vfsName ?? WORKBOOK_VFS_NAME,
    directory: options.directory ?? WORKBOOK_VFS_DIRECTORY,
    initialCapacity: options.initialCapacity ?? WORKBOOK_VFS_INITIAL_CAPACITY,
  };

  return {
    async open(documentId: string): Promise<WorkbookLocalStore> {
      try {
        const { poolUtil } = await getSqliteRuntime(resolvedOptions);
        const path = `/workbooks/${sanitizeDocumentId(documentId)}.sqlite`;
        const db = new poolUtil.OpfsSAHPoolDb(path);
        initializeWorkbookLocalStoreSchema(db);
        return new SqliteWorkbookLocalStore(db, documentId);
      } catch (error) {
        if (error instanceof WorkbookLocalStoreLockedError || isAccessHandleConflict(error)) {
          throw toWorkbookLocalStoreLockedError(documentId, error);
        }
        throw error;
      }
    },
  };
}

export function createMemoryWorkbookLocalStoreFactory(): WorkbookLocalStoreFactory {
  const databases = new Map<string, Database>();

  return {
    async open(documentId: string): Promise<WorkbookLocalStore> {
      let db = databases.get(documentId);
      if (!db) {
        const sqlite3 = await getMemorySqliteRuntime();
        db = new sqlite3.oo1.DB(":memory:", "c");
        initializeWorkbookLocalStoreSchema(db);
        databases.set(documentId, db);
      }
      return new SqliteWorkbookLocalStore(db, documentId, false);
    },
  };
}
