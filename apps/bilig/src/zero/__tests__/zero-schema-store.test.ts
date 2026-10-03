import { describe, expect, it, vi } from 'vitest'
import { zeroSchemaServerColumnNamesByTable } from '@bilig/zero-sync'
import { ensureZeroServiceSchema } from '../schema-bootstrap.js'
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
  it('starts with current authority columns and invariants without historical upgrades', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    await ensureZeroSyncSchema({ query })
    const sql = query.mock.calls.map(([text]) => String(text)).join('\n')
    expect(sql).toContain("owner_user_id TEXT NOT NULL DEFAULT 'system'")
    expect(sql).toContain('source_projection_version BIGINT NOT NULL DEFAULT 2')
    expect(sql).toContain('sheet_id INTEGER NOT NULL CONSTRAINT sheets_sheet_id_positive_chk CHECK (sheet_id > 0)')
    expect(sql).toContain('CREATE UNIQUE INDEX IF NOT EXISTS workbook_event_workbook_client_mutation_idx')
    expect(sql).not.toContain('ALTER TABLE')
    expect(sql).not.toContain('UPDATE ')
    expect(sql).not.toContain('computed_cells')
  })

  it('bootstraps every shared Zero schema table and column', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await ensureZeroServiceSchema(db)

    const columnsByTable = collectBootstrappedColumns(query.mock.calls.map(([text]) => String(text)))
    expect([...columnsByTable.keys()].toSorted()).toEqual(expect.arrayContaining(Object.keys(zeroSchemaServerColumnNamesByTable)))
    for (const [tableName, serverColumnNames] of Object.entries(zeroSchemaServerColumnNamesByTable)) {
      const bootstrappedColumns = columnsByTable.get(tableName)
      expect(bootstrappedColumns, `${tableName} is missing from schema bootstrap`).toBeDefined()
      expect([...(bootstrappedColumns ?? [])].toSorted()).toEqual(expect.arrayContaining([...serverColumnNames]))
    }
  })
})
