import { describe, expect, it, vi } from 'vitest'
import {
  assertZeroDataMigrationsReady,
  resolveAllowPendingCleanupMigrations,
  resolveRunDataMigrationsOnBoot,
  runPendingZeroDataMigrations,
  zeroDataMigrations,
  type PendingZeroDataMigrationsError,
  type ZeroDataMigrationConnection,
  type ZeroDataMigrationDefinition,
  type ZeroDataMigrationPool,
} from '../data-migration-runner.js'
import type { QueryResultRow } from '../store.js'
import { runQueryableTransaction } from '../transaction-support.js'

interface RecordedQuery {
  readonly text: string
  readonly values: readonly unknown[] | undefined
}

class FakeMigrationClient implements ZeroDataMigrationConnection {
  readonly calls: RecordedQuery[] = []
  readonly release = vi.fn()

  constructor(
    private readonly responders: readonly ((text: string, values: readonly unknown[] | undefined) => QueryResultRow[] | null)[] = [],
  ) {}

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    for (const responder of this.responders) {
      const rows = responder(text, values)
      if (rows) {
        return { rows: rows.filter((row): row is T => row !== null) }
      }
    }
    return { rows: [] }
  }
}

class FakeMigrationPool implements ZeroDataMigrationPool {
  constructor(readonly client: FakeMigrationClient) {}

  async connect(): Promise<ZeroDataMigrationConnection> {
    return this.client
  }

  query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    return this.client.query(text, values)
  }
}

class FakeConnectedMigrationClient extends FakeMigrationClient {
  readonly connect = vi.fn(async () => this)
}

describe('data migration runner', () => {
  it('runs pending migrations in order and records them in the ledger', async () => {
    const executed: string[] = []
    const migrations: readonly ZeroDataMigrationDefinition[] = [
      {
        name: 'required-a',
        classification: 'required',
        async run() {
          executed.push('required-a')
        },
      },
      {
        name: 'cleanup-b',
        classification: 'cleanup',
        async run() {
          executed.push('cleanup-b')
        },
      },
    ]
    const client = new FakeMigrationClient()
    const pool = new FakeMigrationPool(client)

    const result = await runPendingZeroDataMigrations(pool, {
      codeVersion: 'test-sha',
      migrations,
    })

    expect(executed).toEqual(['required-a', 'cleanup-b'])
    expect(result.appliedThisRun).toEqual(['required-a', 'cleanup-b'])
    expect(client.calls.some((call) => call.text.includes('SELECT pg_advisory_lock(hashtext($1))'))).toBe(true)
    expect(client.calls.some((call) => call.text.includes('SELECT pg_advisory_unlock(hashtext($1))'))).toBe(true)
    expect(client.calls.filter((call) => call.text.includes('INSERT INTO bilig_data_migration')).map((call) => call.values?.[0])).toEqual([
      'required-a',
      'cleanup-b',
    ])
    expect(client.release).toHaveBeenCalledOnce()
  })

  it('does not reconnect an already connected migration client inside helper transactions', async () => {
    const client = new FakeConnectedMigrationClient()
    const pool = new FakeMigrationPool(client)
    const migrations: readonly ZeroDataMigrationDefinition[] = [
      {
        name: 'nested-helper-transaction',
        classification: 'required',
        async run(db) {
          await runQueryableTransaction(db, async (transactionDb) => {
            await transactionDb.query('SELECT 1')
          })
        },
      },
    ]

    const result = await runPendingZeroDataMigrations(pool, {
      codeVersion: 'test-sha',
      migrations,
    })

    expect(result.appliedThisRun).toEqual(['nested-helper-transaction'])
    expect(client.connect).not.toHaveBeenCalled()
    expect(client.calls.filter((call) => call.text === 'BEGIN')).toHaveLength(1)
    expect(client.calls.filter((call) => call.text === 'COMMIT')).toHaveLength(1)
    expect(client.release).toHaveBeenCalledOnce()
  })

  it('fails startup readiness when required migrations are still pending', async () => {
    const client = new FakeMigrationClient([
      (text) => (text.includes('FROM bilig_data_migration') ? [{ name: 'sheet-id-repair' } satisfies QueryResultRow] : null),
    ])

    await expect(
      assertZeroDataMigrationsReady(client, {
        migrations: [
          {
            name: 'sheet-id-repair',
            classification: 'required',
            async run() {},
          },
          {
            name: 'workbook-change-backfill',
            classification: 'required',
            async run() {},
          },
        ],
      }),
    ).rejects.toMatchObject<PendingZeroDataMigrationsError>({
      pendingRequired: ['workbook-change-backfill'],
      pendingCleanup: [],
    })
  })

  it('allows pending cleanup migrations when explicitly configured', async () => {
    const client = new FakeMigrationClient([
      (text) => (text.includes('FROM bilig_data_migration') ? [{ name: 'sheet-id-repair' } satisfies QueryResultRow] : null),
    ])

    await expect(
      assertZeroDataMigrationsReady(client, {
        allowPendingCleanup: true,
        migrations: [
          {
            name: 'sheet-id-repair',
            classification: 'required',
            async run() {},
          },
          {
            name: 'legacy-zero-style-format-table-retirement',
            classification: 'cleanup',
            async run() {},
          },
        ],
      }),
    ).resolves.toEqual({
      applied: ['sheet-id-repair'],
      pendingRequired: [],
      pendingCleanup: ['legacy-zero-style-format-table-retirement'],
    })
  })

  it('resolves migration startup flags from explicit boolean env values', () => {
    expect(resolveRunDataMigrationsOnBoot({})).toBe(false)
    expect(resolveRunDataMigrationsOnBoot({ BILIG_RUN_DATA_MIGRATIONS_ON_BOOT: 'true' })).toBe(true)
    expect(resolveRunDataMigrationsOnBoot({ BILIG_RUN_DATA_MIGRATIONS_ON_BOOT: '1' })).toBe(true)
    expect(resolveRunDataMigrationsOnBoot({ BILIG_RUN_DATA_MIGRATIONS_ON_BOOT: 'false' })).toBe(false)
    expect(resolveRunDataMigrationsOnBoot({ BILIG_RUN_DATA_MIGRATIONS_ON_BOOT: '0' })).toBe(false)
    expect(resolveAllowPendingCleanupMigrations({ BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS: 'true' })).toBe(true)
  })

  it('rejects malformed migration startup flags instead of silently disabling migrations', () => {
    expect(() => resolveRunDataMigrationsOnBoot({ BILIG_RUN_DATA_MIGRATIONS_ON_BOOT: 'yes' })).toThrow(
      'BILIG_RUN_DATA_MIGRATIONS_ON_BOOT must be "1", "true", "0", or "false" when set, got yes',
    )
    expect(() => resolveAllowPendingCleanupMigrations({ BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS: 'maybe' })).toThrow(
      'BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS must be "1", "true", "0", or "false" when set, got maybe',
    )
  })

  it('tracks workbook event client mutation uniqueness as a required migration', () => {
    expect(zeroDataMigrations.map((migration) => [migration.name, migration.classification])).toContainEqual([
      'workbook-event-client-mutation-id-uniqueness',
      'required',
    ])
  })

  it('enforces sheet id invariants immediately after repairing sheet ids', () => {
    expect(zeroDataMigrations.map((migration) => migration.name).slice(0, 2)).toEqual(['sheet-id-repair', 'sheet-id-invariant-enforcement'])
    expect(zeroDataMigrations.find((migration) => migration.name === 'sheet-id-invariant-enforcement')?.classification).toBe('required')
  })
})
