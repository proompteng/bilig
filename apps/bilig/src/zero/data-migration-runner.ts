import type { QueryResultRow, Queryable } from './store.js'
import { backfillWorkbookSnapshotsFromInlineState } from './workbook-calculation-store.js'
import { backfillWorkbookChanges } from './workbook-change-store.js'
import {
  backfillCellEvalStyleJson,
  backfillWorkbookSourceProjectionVersion,
  dropLegacyZeroSyncSchemaObjects,
  repairWorkbookSheetIdsForMigration,
} from './workbook-migration-store.js'

export type ZeroDataMigrationClassification = 'required' | 'cleanup'

export interface ZeroDataMigrationDefinition {
  readonly name: string
  readonly classification: ZeroDataMigrationClassification
  run(db: Queryable): Promise<void>
}

export interface ZeroDataMigrationStatus {
  readonly applied: readonly string[]
  readonly pendingRequired: readonly string[]
  readonly pendingCleanup: readonly string[]
}

export interface RunZeroDataMigrationsResult extends ZeroDataMigrationStatus {
  readonly appliedThisRun: readonly string[]
}

export interface ZeroDataMigrationConnection extends Queryable {
  release(): void
}

export interface ZeroDataMigrationPool extends Queryable {
  connect(): Promise<ZeroDataMigrationConnection>
}

interface ZeroDataMigrationLedgerRow extends QueryResultRow {
  readonly name?: unknown
}

const DATA_MIGRATION_LOCK_KEY = 'bilig-zero-data-migrations'

const zeroDataMigrations = [
  {
    name: 'sheet-id-repair',
    classification: 'required',
    run: repairWorkbookSheetIdsForMigration,
  },
  {
    name: 'workbook-source-projection-v2-backfill',
    classification: 'required',
    run: backfillWorkbookSourceProjectionVersion,
  },
  {
    name: 'cell-eval-style-json-backfill',
    classification: 'required',
    run: backfillCellEvalStyleJson,
  },
  {
    name: 'workbook-change-backfill',
    classification: 'required',
    run: backfillWorkbookChanges,
  },
  {
    name: 'workbook-snapshot-json-v1-backfill',
    classification: 'cleanup',
    run: backfillWorkbookSnapshotsFromInlineState,
  },
  {
    name: 'legacy-zero-style-format-table-retirement',
    classification: 'cleanup',
    run: dropLegacyZeroSyncSchemaObjects,
  },
] as const satisfies readonly ZeroDataMigrationDefinition[]

export class PendingZeroDataMigrationsError extends Error {
  readonly pendingRequired: readonly string[]
  readonly pendingCleanup: readonly string[]

  constructor(status: ZeroDataMigrationStatus) {
    super(formatPendingZeroDataMigrationMessage(status))
    this.name = 'PendingZeroDataMigrationsError'
    this.pendingRequired = status.pendingRequired
    this.pendingCleanup = status.pendingCleanup
  }
}

export async function ensureZeroDataMigrationSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS bilig_data_migration (
      name TEXT PRIMARY KEY,
      applied_at TIMESTAMPTZ NOT NULL,
      code_version TEXT NOT NULL,
      details JSONB NOT NULL DEFAULT '{}'::jsonb
    );
  `)
}

async function getZeroDataMigrationStatus(
  db: Queryable,
  migrations: readonly ZeroDataMigrationDefinition[] = zeroDataMigrations,
): Promise<ZeroDataMigrationStatus> {
  await ensureZeroDataMigrationSchema(db)
  const appliedNames = await loadAppliedMigrationNames(db)
  const pendingRequired: string[] = []
  const pendingCleanup: string[] = []
  for (const migration of migrations) {
    if (appliedNames.has(migration.name)) {
      continue
    }
    if (migration.classification === 'required') {
      pendingRequired.push(migration.name)
      continue
    }
    pendingCleanup.push(migration.name)
  }
  return {
    applied: [...appliedNames].toSorted((left, right) => left.localeCompare(right)),
    pendingRequired,
    pendingCleanup,
  }
}

export async function assertZeroDataMigrationsReady(
  db: Queryable,
  options: {
    readonly allowPendingCleanup?: boolean
    readonly migrations?: readonly ZeroDataMigrationDefinition[]
  } = {},
): Promise<ZeroDataMigrationStatus> {
  const status = await getZeroDataMigrationStatus(db, options.migrations)
  const cleanupPending = !options.allowPendingCleanup && status.pendingCleanup.length > 0
  if (status.pendingRequired.length > 0 || cleanupPending) {
    throw new PendingZeroDataMigrationsError(status)
  }
  return status
}

export async function runPendingZeroDataMigrations(
  pool: ZeroDataMigrationPool,
  options: {
    readonly codeVersion?: string
    readonly migrations?: readonly ZeroDataMigrationDefinition[]
  } = {},
): Promise<RunZeroDataMigrationsResult> {
  const migrations = options.migrations ?? zeroDataMigrations
  const codeVersion = options.codeVersion ?? resolveZeroDataMigrationCodeVersion()
  const client = await pool.connect()
  try {
    await ensureZeroDataMigrationSchema(client)
    await client.query(`SELECT pg_advisory_lock(hashtext($1))`, [DATA_MIGRATION_LOCK_KEY])

    const appliedThisRun: string[] = []
    const appliedNames = await loadAppliedMigrationNames(client)
    await runPendingMigrationBatch(client, {
      migrations: migrations.filter((migration) => !appliedNames.has(migration.name)),
      codeVersion,
      appliedNames,
      appliedThisRun,
    })

    const status = await getZeroDataMigrationStatus(client, migrations)
    return {
      ...status,
      appliedThisRun,
    }
  } finally {
    await client.query(`SELECT pg_advisory_unlock(hashtext($1))`, [DATA_MIGRATION_LOCK_KEY]).catch(() => undefined)
    client.release()
  }
}

export function resolveRunDataMigrationsOnBoot(env: Record<string, string | undefined> = process.env): boolean {
  return parseBooleanEnv(env['BILIG_RUN_DATA_MIGRATIONS_ON_BOOT'])
}

export function resolveAllowPendingCleanupMigrations(env: Record<string, string | undefined> = process.env): boolean {
  return parseBooleanEnv(env['BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS'])
}

function resolveZeroDataMigrationCodeVersion(env: Record<string, string | undefined> = process.env): string {
  return (
    env['BILIG_RELEASE_VERSION']?.trim() ||
    env['BILIG_IMAGE_TAG']?.trim() ||
    env['GIT_COMMIT_SHA']?.trim() ||
    env['VERCEL_GIT_COMMIT_SHA']?.trim() ||
    'dev'
  )
}

async function loadAppliedMigrationNames(db: Queryable): Promise<Set<string>> {
  const result = await db.query<ZeroDataMigrationLedgerRow>(
    `
      SELECT name
      FROM bilig_data_migration
      ORDER BY name ASC
    `,
  )
  return new Set(result.rows.flatMap((row) => (typeof row.name === 'string' && row.name.length > 0 ? [row.name] : [])))
}

async function runPendingMigrationBatch(
  client: ZeroDataMigrationConnection,
  state: {
    readonly migrations: readonly ZeroDataMigrationDefinition[]
    readonly codeVersion: string
    readonly appliedNames: Set<string>
    readonly appliedThisRun: string[]
  },
  index = 0,
): Promise<void> {
  const migration = state.migrations[index]
  if (!migration) {
    return
  }
  await client.query('BEGIN')
  try {
    await migration.run(client)
    await client.query(
      `
        INSERT INTO bilig_data_migration (
          name,
          applied_at,
          code_version,
          details
        )
        VALUES ($1, NOW(), $2, $3::jsonb)
        ON CONFLICT (name)
        DO NOTHING
      `,
      [
        migration.name,
        state.codeVersion,
        JSON.stringify({
          classification: migration.classification,
        }),
      ],
    )
    await client.query('COMMIT')
    state.appliedNames.add(migration.name)
    state.appliedThisRun.push(migration.name)
  } catch (error) {
    await client.query('ROLLBACK').catch(() => undefined)
    throw error
  }
  await runPendingMigrationBatch(client, state, index + 1)
}

function formatPendingZeroDataMigrationMessage(status: ZeroDataMigrationStatus): string {
  const parts: string[] = []
  if (status.pendingRequired.length > 0) {
    parts.push(`required: ${status.pendingRequired.join(', ')}`)
  }
  if (status.pendingCleanup.length > 0) {
    parts.push(`cleanup: ${status.pendingCleanup.join(', ')}`)
  }
  const detail = parts.length > 0 ? ` (${parts.join('; ')})` : ''
  return `Pending Zero data migrations. Run bun scripts/run-zero-data-migrations.ts before starting bilig${detail}.`
}

function parseBooleanEnv(value: string | undefined): boolean {
  if (!value) {
    return false
  }
  const normalized = value.trim().toLowerCase()
  return normalized === '1' || normalized === 'true' || normalized === 'yes'
}
