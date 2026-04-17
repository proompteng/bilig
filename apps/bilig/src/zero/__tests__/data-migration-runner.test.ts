import { describe, expect, it, vi } from 'vitest'
import {
  assertZeroDataMigrationsReady,
  runPendingZeroDataMigrations,
  type PendingZeroDataMigrationsError,
  type ZeroDataMigrationConnection,
  type ZeroDataMigrationDefinition,
  type ZeroDataMigrationPool,
} from '../data-migration-runner.js'
import type { QueryResultRow } from '../store.js'

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
})
