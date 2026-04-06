import sqlite3InitModule, {
  type Database,
  type SAHPoolUtil,
  type SqlValue,
  type Sqlite3Static,
} from "@sqlite.org/sqlite-wasm";
import {
  ValueTag,
  type CellSnapshot,
  type CellStyleRecord,
  type WorkbookAxisEntrySnapshot,
} from "@bilig/protocol";
import type {
  WorkbookLocalAuthoritativeBase,
  WorkbookLocalViewportBase,
  WorkbookLocalViewportCell,
} from "./workbook-local-base.js";

const WORKBOOK_VFS_NAME = "bilig-opfs-sahpool";
const WORKBOOK_VFS_DIRECTORY = "/bilig/workbooks";
const WORKBOOK_VFS_INITIAL_CAPACITY = 12;

let sqliteRuntimePromise: Promise<{ sqlite3: Sqlite3Static; poolUtil: SAHPoolUtil }> | null = null;

export class WorkbookLocalStoreLockedError extends Error {
  override readonly name = "WorkbookLocalStoreLockedError";
}

export interface WorkbookStoredState {
  readonly snapshot: unknown;
  readonly replica: unknown;
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
  loadState(): Promise<WorkbookStoredState | null>;
  saveState(state: WorkbookStoredState): Promise<void>;
  listPendingMutations(): Promise<WorkbookLocalMutationRecord[]>;
  appendPendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void>;
  updatePendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void>;
  removePendingMutation(id: string): Promise<void>;
  replaceAuthoritativeBase(base: WorkbookLocalAuthoritativeBase): void;
  readViewportBase(
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
  return (
    error instanceof Error &&
    error.message.includes("createSyncAccessHandle") &&
    error.message.includes("Access Handles cannot be created")
  );
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

function parseCellSnapshotValue(value: unknown): CellSnapshot["value"] | null {
  if (!isRecord(value) || typeof value["tag"] !== "number") {
    return null;
  }
  const tag = value["tag"] as ValueTag;
  switch (tag) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty };
    case ValueTag.Number:
      return typeof value["value"] === "number"
        ? { tag: ValueTag.Number, value: value["value"] }
        : null;
    case ValueTag.Boolean:
      return typeof value["value"] === "boolean"
        ? { tag: ValueTag.Boolean, value: value["value"] }
        : null;
    case ValueTag.String:
      return typeof value["value"] === "string"
        ? {
            tag: ValueTag.String,
            value: value["value"],
            stringId: typeof value["stringId"] === "number" ? value["stringId"] : 0,
          }
        : null;
    case ValueTag.Error:
      return typeof value["code"] === "number"
        ? { tag: ValueTag.Error, code: value["code"] }
        : null;
    default:
      return null;
  }
}

function parseCellSnapshotFromBaseRow(
  row: Record<string, SqlValue>,
): WorkbookLocalViewportCell | null {
  const address = row["address"];
  const sheetName = row["sheetName"];
  const rowNum = row["rowNum"];
  const colNum = row["colNum"];
  const valueJson = row["valueJson"];
  const flags = row["flags"];
  const version = row["version"];
  if (
    typeof address !== "string" ||
    typeof sheetName !== "string" ||
    typeof rowNum !== "number" ||
    typeof colNum !== "number" ||
    typeof valueJson !== "string" ||
    typeof flags !== "number" ||
    typeof version !== "number"
  ) {
    return null;
  }
  try {
    const parsedValue = parseCellSnapshotValue(JSON.parse(valueJson) as unknown);
    if (!parsedValue) {
      return null;
    }
    const inputJson = row["inputJson"];
    const snapshot: CellSnapshot = {
      sheetName,
      address,
      value: parsedValue,
      flags,
      version,
    };
    if (typeof inputJson === "string") {
      const parsedInput = JSON.parse(inputJson) as unknown;
      if (
        parsedInput === null ||
        typeof parsedInput === "boolean" ||
        typeof parsedInput === "number" ||
        typeof parsedInput === "string"
      ) {
        snapshot.input = parsedInput;
      }
    }
    if (typeof row["formula"] === "string") {
      snapshot.formula = row["formula"];
    }
    if (typeof row["format"] === "string") {
      snapshot.format = row["format"];
    }
    if (typeof row["styleId"] === "string") {
      snapshot.styleId = row["styleId"];
    }
    if (typeof row["numberFormatId"] === "string") {
      snapshot.numberFormatId = row["numberFormatId"];
    }
    return {
      row: rowNum,
      col: colNum,
      snapshot,
    };
  } catch {
    return null;
  }
}

function parseAxisEntrySnapshot(row: Record<string, SqlValue>): WorkbookAxisEntrySnapshot | null {
  const id = row["id"];
  const entryIndex = row["entryIndex"];
  if (typeof id !== "string" || typeof entryIndex !== "number") {
    return null;
  }
  const entry: WorkbookAxisEntrySnapshot = {
    id,
    index: entryIndex,
  };
  if (typeof row["size"] === "number") {
    entry.size = row["size"];
  }
  if (typeof row["hidden"] === "number") {
    entry.hidden = row["hidden"] !== 0;
  } else if (typeof row["hidden"] === "boolean") {
    entry.hidden = row["hidden"];
  }
  return entry;
}

function parseCellStyleRecord(row: Record<string, SqlValue>): CellStyleRecord | null {
  const id = row["id"];
  const recordJson = row["recordJson"];
  if (typeof id !== "string" || typeof recordJson !== "string") {
    return null;
  }
  try {
    const parsed = JSON.parse(recordJson) as unknown;
    if (!isRecord(parsed)) {
      return null;
    }
    return {
      ...(parsed as Omit<CellStyleRecord, "id">),
      id,
    };
  } catch {
    return null;
  }
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

function initializeSchema(db: Database): void {
  db.exec(`
    PRAGMA journal_mode = WAL;
    PRAGMA synchronous = NORMAL;
    PRAGMA temp_store = MEMORY;
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS runtime_state (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      snapshot_json TEXT NOT NULL,
      replica_json TEXT NOT NULL,
      authoritative_revision INTEGER NOT NULL,
      applied_pending_local_seq INTEGER NOT NULL DEFAULT 0,
      updated_at_ms INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS pending_op (
      op_id TEXT PRIMARY KEY,
      local_seq INTEGER NOT NULL UNIQUE,
      base_revision INTEGER NOT NULL,
      method TEXT NOT NULL,
      args_json TEXT NOT NULL,
      enqueued_at_ms INTEGER NOT NULL,
      submitted_at_ms INTEGER,
      status TEXT NOT NULL CHECK (status IN ('pending', 'submitted'))
    );

    CREATE INDEX IF NOT EXISTS pending_op_local_seq_idx
      ON pending_op(local_seq);

    CREATE TABLE IF NOT EXISTS authoritative_sheet (
      name TEXT PRIMARY KEY,
      sort_order INTEGER NOT NULL,
      freeze_rows INTEGER NOT NULL DEFAULT 0,
      freeze_cols INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS authoritative_cell_input (
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      address TEXT NOT NULL,
      row_num INTEGER NOT NULL,
      col_num INTEGER NOT NULL,
      input_json TEXT,
      formula TEXT,
      format TEXT,
      PRIMARY KEY (sheet_name, address)
    );

    CREATE TABLE IF NOT EXISTS authoritative_cell_render (
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      address TEXT NOT NULL,
      row_num INTEGER NOT NULL,
      col_num INTEGER NOT NULL,
      value_json TEXT NOT NULL,
      flags INTEGER NOT NULL,
      version INTEGER NOT NULL,
      style_id TEXT,
      number_format_id TEXT,
      PRIMARY KEY (sheet_name, address)
    );

    CREATE TABLE IF NOT EXISTS authoritative_row_axis (
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE IF NOT EXISTS authoritative_column_axis (
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE IF NOT EXISTS authoritative_style (
      style_id TEXT PRIMARY KEY,
      record_json TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS authoritative_cell_render_viewport_idx
      ON authoritative_cell_render(sheet_name, row_num, col_num);

    CREATE INDEX IF NOT EXISTS authoritative_row_axis_viewport_idx
      ON authoritative_row_axis(sheet_name, axis_index);

    CREATE INDEX IF NOT EXISTS authoritative_column_axis_viewport_idx
      ON authoritative_column_axis(sheet_name, axis_index);
  `);
  const appliedPendingColumn = readSingleObjectRow(
    db,
    `
      SELECT 1 AS present
        FROM pragma_table_info('runtime_state')
       WHERE name = 'applied_pending_local_seq'
    `,
  );
  if (!appliedPendingColumn) {
    db.exec(`
      ALTER TABLE runtime_state
      ADD COLUMN applied_pending_local_seq INTEGER NOT NULL DEFAULT 0
    `);
  }
  const submittedAtColumn = readSingleObjectRow(
    db,
    `
      SELECT 1 AS present
        FROM pragma_table_info('pending_op')
       WHERE name = 'submitted_at_ms'
    `,
  );
  if (!submittedAtColumn) {
    db.exec(`
      ALTER TABLE pending_op
      ADD COLUMN submitted_at_ms INTEGER
    `);
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
      try {
        const sqlite3 = await sqlite3InitModule();
        const poolUtil = await sqlite3.installOpfsSAHPoolVfs({
          name: options.vfsName,
          directory: options.directory,
          initialCapacity: options.initialCapacity,
        });
        return { sqlite3, poolUtil };
      } catch (error) {
        sqliteRuntimePromise = null;
        throw error;
      }
    })();
  }
  return await sqliteRuntimePromise;
}

class OpfsWorkbookLocalStore implements WorkbookLocalStore {
  constructor(private readonly db: Database) {}

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

  async saveState(state: WorkbookStoredState): Promise<void> {
    this.db.transaction((db) => {
      db.exec(
        `
          INSERT INTO runtime_state (
            id,
            snapshot_json,
            replica_json,
            authoritative_revision,
            applied_pending_local_seq,
            updated_at_ms
          )
          VALUES (1, ?, ?, ?, ?, ?)
          ON CONFLICT(id) DO UPDATE SET
            snapshot_json = excluded.snapshot_json,
            replica_json = excluded.replica_json,
            authoritative_revision = excluded.authoritative_revision,
            applied_pending_local_seq = excluded.applied_pending_local_seq,
            updated_at_ms = excluded.updated_at_ms
        `,
        {
          bind: [
            JSON.stringify(state.snapshot),
            JSON.stringify(state.replica),
            state.authoritativeRevision,
            state.appliedPendingLocalSeq,
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

  replaceAuthoritativeBase(base: WorkbookLocalAuthoritativeBase): void {
    this.db.transaction((db) => {
      db.exec("DELETE FROM authoritative_cell_input");
      db.exec("DELETE FROM authoritative_cell_render");
      db.exec("DELETE FROM authoritative_row_axis");
      db.exec("DELETE FROM authoritative_column_axis");
      db.exec("DELETE FROM authoritative_style");
      db.exec("DELETE FROM authoritative_sheet");

      const insertSheet = db.prepare(
        `
          INSERT INTO authoritative_sheet (name, sort_order, freeze_rows, freeze_cols)
          VALUES (?, ?, ?, ?)
        `,
      );
      const insertInput = db.prepare(
        `
          INSERT INTO authoritative_cell_input (
            sheet_name,
            address,
            row_num,
            col_num,
            input_json,
            formula,
            format
          )
          VALUES (?, ?, ?, ?, ?, ?, ?)
        `,
      );
      const insertRender = db.prepare(
        `
          INSERT INTO authoritative_cell_render (
            sheet_name,
            address,
            row_num,
            col_num,
            value_json,
            flags,
            version,
            style_id,
            number_format_id
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        `,
      );
      const insertAxis = (tableName: "authoritative_row_axis" | "authoritative_column_axis") =>
        db.prepare(
          `
            INSERT INTO ${tableName} (
              sheet_name,
              axis_index,
              axis_id,
              size,
              hidden
            )
            VALUES (?, ?, ?, ?, ?)
          `,
        );
      const insertStyle = db.prepare(
        `
          INSERT INTO authoritative_style (style_id, record_json)
          VALUES (?, ?)
        `,
      );
      const insertRowAxis = insertAxis("authoritative_row_axis");
      const insertColumnAxis = insertAxis("authoritative_column_axis");
      try {
        for (const sheet of base.sheets) {
          insertSheet.bind([sheet.name, sheet.sortOrder, sheet.freezeRows, sheet.freezeCols]);
          insertSheet.step();
          insertSheet.reset();
        }
        for (const cell of base.cellInputs) {
          insertInput.bind([
            cell.sheetName,
            cell.address,
            cell.rowNum,
            cell.colNum,
            cell.input === undefined ? null : JSON.stringify(cell.input),
            cell.formula ?? null,
            cell.format ?? null,
          ]);
          insertInput.step();
          insertInput.reset();
        }
        for (const cell of base.cellRenders) {
          insertRender.bind([
            cell.sheetName,
            cell.address,
            cell.rowNum,
            cell.colNum,
            JSON.stringify(cell.value),
            cell.flags,
            cell.version,
            cell.styleId ?? null,
            cell.numberFormatId ?? null,
          ]);
          insertRender.step();
          insertRender.reset();
        }
        for (const axis of base.rowAxisEntries) {
          insertRowAxis.bind([
            axis.sheetName,
            axis.entry.index,
            axis.entry.id,
            axis.entry.size ?? null,
            axis.entry.hidden ?? null,
          ]);
          insertRowAxis.step();
          insertRowAxis.reset();
        }
        for (const axis of base.columnAxisEntries) {
          insertColumnAxis.bind([
            axis.sheetName,
            axis.entry.index,
            axis.entry.id,
            axis.entry.size ?? null,
            axis.entry.hidden ?? null,
          ]);
          insertColumnAxis.step();
          insertColumnAxis.reset();
        }
        for (const style of base.styles) {
          insertStyle.bind([style.id, JSON.stringify(style)]);
          insertStyle.step();
          insertStyle.reset();
        }
      } finally {
        insertSheet.finalize();
        insertInput.finalize();
        insertRender.finalize();
        insertRowAxis.finalize();
        insertColumnAxis.finalize();
        insertStyle.finalize();
      }
    });
  }

  readViewportBase(
    sheetName: string,
    viewport: {
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    },
  ): WorkbookLocalViewportBase | null {
    const sheetRecord = readSingleObjectRow(
      this.db,
      `
        SELECT name,
               sort_order AS sortOrder,
               freeze_rows AS freezeRows,
               freeze_cols AS freezeCols
          FROM authoritative_sheet
         WHERE name = ?
      `,
      [sheetName],
    );
    if (!sheetRecord) {
      return null;
    }

    const cells: WorkbookLocalViewportCell[] = [];
    const cellStatement = this.db.prepare(
      `
        SELECT render.sheet_name AS sheetName,
               render.address AS address,
               render.row_num AS rowNum,
               render.col_num AS colNum,
               render.value_json AS valueJson,
               render.flags AS flags,
               render.version AS version,
               render.style_id AS styleId,
               render.number_format_id AS numberFormatId,
               input.input_json AS inputJson,
               input.formula AS formula,
               input.format AS format
          FROM authoritative_cell_render AS render
          LEFT JOIN authoritative_cell_input AS input
            ON input.sheet_name = render.sheet_name
           AND input.address = render.address
         WHERE render.sheet_name = ?
           AND render.row_num >= ?
           AND render.row_num <= ?
           AND render.col_num >= ?
           AND render.col_num <= ?
         ORDER BY render.row_num ASC, render.col_num ASC
      `,
    );
    const styleIds = new Set<string>(["style-0"]);
    try {
      cellStatement.bind([
        sheetName,
        viewport.rowStart,
        viewport.rowEnd,
        viewport.colStart,
        viewport.colEnd,
      ]);
      while (cellStatement.step()) {
        const parsed = parseCellSnapshotFromBaseRow(cellStatement.get({}));
        if (!parsed) {
          continue;
        }
        cells.push(parsed);
        if (parsed.snapshot.styleId) {
          styleIds.add(parsed.snapshot.styleId);
        }
      }
    } finally {
      cellStatement.finalize();
    }

    const readAxisEntries = (
      tableName: "authoritative_row_axis" | "authoritative_column_axis",
      start: number,
      end: number,
    ): WorkbookAxisEntrySnapshot[] => {
      const rows: WorkbookAxisEntrySnapshot[] = [];
      const statement = this.db.prepare(
        `
          SELECT axis_id AS id,
                 axis_index AS entryIndex,
                 size,
                 hidden
            FROM ${tableName}
           WHERE sheet_name = ?
             AND axis_index >= ?
             AND axis_index <= ?
           ORDER BY axis_index ASC
        `,
      );
      try {
        statement.bind([sheetName, start, end]);
        while (statement.step()) {
          const entry = parseAxisEntrySnapshot(statement.get({}));
          if (entry) {
            rows.push(entry);
          }
        }
      } finally {
        statement.finalize();
      }
      return rows;
    };

    const readStyles = (): CellStyleRecord[] => {
      const styles: CellStyleRecord[] = [];
      const statement = this.db.prepare(
        `
          SELECT style_id AS id,
                 record_json AS recordJson
            FROM authoritative_style
           WHERE style_id = ?
        `,
      );
      try {
        for (const styleId of styleIds) {
          statement.bind([styleId]);
          if (statement.step()) {
            const style = parseCellStyleRecord(statement.get({}));
            if (style) {
              styles.push(style);
            }
          } else if (styleId === "style-0") {
            styles.push({ id: "style-0" });
          }
          statement.reset();
        }
      } finally {
        statement.finalize();
      }
      return styles;
    };

    return {
      sheetName,
      cells,
      rowAxisEntries: readAxisEntries("authoritative_row_axis", viewport.rowStart, viewport.rowEnd),
      columnAxisEntries: readAxisEntries(
        "authoritative_column_axis",
        viewport.colStart,
        viewport.colEnd,
      ),
      styles: readStyles(),
    };
  }

  close(): void {
    this.db.close();
  }
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
        initializeSchema(db);
        return new OpfsWorkbookLocalStore(db);
      } catch (error) {
        if (isAccessHandleConflict(error)) {
          throw new WorkbookLocalStoreLockedError(
            `Workbook local store is locked by another tab for ${documentId}`,
          );
        }
        throw error;
      }
    },
  };
}
