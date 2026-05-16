import { describe, expect, it, vi } from 'vitest'
import { zeroSchemaServerColumnNamesByTable } from '@bilig/zero-sync'
import { ensureWorkbookChangeSchema } from '../workbook-change-store.js'
import { ensureWorkbookChatThreadSchema } from '../workbook-chat-thread-store.js'
import { ensureWorkbookPresenceSchema } from '../presence-store.js'
import { ensureWorkbookWorkflowRunSchema } from '../workbook-workflow-run-store.js'
import { ensureZeroSyncSchema } from '../zero-schema-store.js'
import type { Queryable } from '../store.js'

function collectBootstrappedColumns(calls: readonly string[]): Map<string, Set<string>> {
  const columnsByTable = new Map<string, Set<string>>()
  const ensureTable = (tableName: string) => {
    const columns = columnsByTable.get(tableName) ?? new Set<string>()
    columnsByTable.set(tableName, columns)
    return columns
  }
  for (const text of calls) {
    const createMatch = /CREATE TABLE IF NOT EXISTS\s+([a-z_]+)\s*\(([\s\S]*)\)\s*;?\s*$/iu.exec(text.trim())
    if (createMatch) {
      const [, tableName, body] = createMatch
      const columns = ensureTable(tableName)
      for (const line of body.split('\n')) {
        const columnMatch = /^\s*([a-z_]+)\s+/iu.exec(line.trim())
        if (columnMatch && columnMatch[1] !== 'PRIMARY' && columnMatch[1] !== 'FOREIGN' && columnMatch[1] !== 'UNIQUE') {
          columns.add(columnMatch[1])
        }
      }
    }
    for (const [, tableName, columnName] of text.matchAll(/ALTER TABLE\s+([a-z_]+)[\s\S]*?ADD COLUMN IF NOT EXISTS\s+([a-z_]+)/giu)) {
      ensureTable(tableName).add(columnName)
    }
  }
  return columnsByTable
}

describe('zero schema store', () => {
  it('backfills and enforces workbook authority columns on legacy schemas', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await ensureZeroSyncSchema(db)

    const calls = query.mock.calls.map(([text]) => String(text))
    const ownerBackfillIndex = calls.findIndex((text) => text.includes("SET owner_user_id = 'system'"))
    const ownerNotNullIndex = calls.findIndex((text) => text.includes('ALTER COLUMN owner_user_id SET NOT NULL'))
    const headRevisionBackfillIndex = calls.findIndex((text) => text.includes('SET head_revision = 0'))
    const headRevisionNotNullIndex = calls.findIndex((text) => text.includes('ALTER COLUMN head_revision SET NOT NULL'))
    expect(ownerBackfillIndex).toBeGreaterThan(-1)
    expect(ownerNotNullIndex).toBeGreaterThan(ownerBackfillIndex)
    expect(headRevisionBackfillIndex).toBeGreaterThan(-1)
    expect(headRevisionNotNullIndex).toBeGreaterThan(headRevisionBackfillIndex)
  })

  it('backfills and enforces projection timestamp and revision columns on legacy schemas', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await ensureZeroSyncSchema(db)

    const calls = query.mock.calls.map(([text]) => String(text))
    const cellRevisionBackfillIndex = calls.findIndex((text) => text.includes('UPDATE cells') && text.includes('SET source_revision = 0'))
    const cellRevisionNotNullIndex = calls.findIndex(
      (text) => text.includes('ALTER TABLE cells') && text.includes('ALTER COLUMN source_revision SET NOT NULL'),
    )
    const cellEvalUpdatedAtBackfillIndex = calls.findIndex(
      (text) => text.includes('UPDATE cell_eval') && text.includes('SET updated_at = NOW()'),
    )
    const cellEvalUpdatedAtNotNullIndex = calls.findIndex(
      (text) => text.includes('ALTER TABLE cell_eval') && text.includes('ALTER COLUMN updated_at SET NOT NULL'),
    )
    expect(cellRevisionBackfillIndex).toBeGreaterThan(-1)
    expect(cellRevisionNotNullIndex).toBeGreaterThan(cellRevisionBackfillIndex)
    expect(cellEvalUpdatedAtBackfillIndex).toBeGreaterThan(-1)
    expect(cellEvalUpdatedAtNotNullIndex).toBeGreaterThan(cellEvalUpdatedAtBackfillIndex)
  })

  it('backfills and enforces axis metadata revision columns before indexing', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await ensureZeroSyncSchema(db)

    const calls = query.mock.calls.map(([text]) => String(text))
    const rowBackfillIndex = calls.findIndex((text) => text.includes('UPDATE row_metadata') && text.includes('SET source_revision = 0'))
    const rowNotNullIndex = calls.findIndex(
      (text) => text.includes('ALTER TABLE row_metadata') && text.includes('ALTER COLUMN source_revision SET NOT NULL'),
    )
    const rowIndex = calls.findIndex((text) => text.includes('row_metadata_workbook_sheet_idx'))
    const columnBackfillIndex = calls.findIndex(
      (text) => text.includes('UPDATE column_metadata') && text.includes('SET source_revision = 0'),
    )
    const columnNotNullIndex = calls.findIndex(
      (text) => text.includes('ALTER TABLE column_metadata') && text.includes('ALTER COLUMN source_revision SET NOT NULL'),
    )
    const columnIndex = calls.findIndex((text) => text.includes('column_metadata_workbook_sheet_idx'))
    expect(rowBackfillIndex).toBeGreaterThan(-1)
    expect(rowNotNullIndex).toBeGreaterThan(rowBackfillIndex)
    expect(rowIndex).toBeGreaterThan(rowNotNullIndex)
    expect(columnBackfillIndex).toBeGreaterThan(-1)
    expect(columnNotNullIndex).toBeGreaterThan(columnBackfillIndex)
    expect(columnIndex).toBeGreaterThan(columnNotNullIndex)
  })

  it('creates schema tables without running data-repair migrations', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await ensureZeroSyncSchema(db)

    expect(query.mock.calls.some(([text]) => String(text).includes('CREATE TABLE IF NOT EXISTS workbook_snapshot'))).toBe(true)
    expect(query.mock.calls.some(([text]) => String(text).includes('workbook_event_workbook_client_mutation_idx'))).toBe(false)
    expect(
      query.mock.calls.some(([text]) => String(text).includes('UPDATE sheets SET sheet_id = sort_order + 1 WHERE sheet_id IS NULL')),
    ).toBe(false)
    expect(query.mock.calls.some(([text]) => String(text).includes('INSERT INTO workbook_snapshot'))).toBe(false)
  })

  it('bootstraps every shared Zero schema table and column', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await ensureZeroSyncSchema(db)
    await ensureWorkbookPresenceSchema(db)
    await ensureWorkbookChangeSchema(db)
    await ensureWorkbookChatThreadSchema(db)
    await ensureWorkbookWorkflowRunSchema(db)

    const columnsByTable = collectBootstrappedColumns(query.mock.calls.map(([text]) => String(text)))
    expect([...columnsByTable.keys()].toSorted()).toEqual(expect.arrayContaining(Object.keys(zeroSchemaServerColumnNamesByTable)))
    for (const [tableName, serverColumnNames] of Object.entries(zeroSchemaServerColumnNamesByTable)) {
      const bootstrappedColumns = columnsByTable.get(tableName)
      expect(bootstrappedColumns, `${tableName} is missing from schema bootstrap`).toBeDefined()
      expect([...(bootstrappedColumns ?? [])].toSorted()).toEqual(expect.arrayContaining([...serverColumnNames]))
    }
  })
})
