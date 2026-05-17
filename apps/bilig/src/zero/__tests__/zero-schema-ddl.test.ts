import { describe, expect, it } from 'vitest'
import { schema, zeroSchemaServerColumnNamesByTable } from '@bilig/zero-sync'
import { createZeroSchemaTableSql } from '../zero-schema-ddl.js'

function normalizedSql(sql: string): string {
  return sql.replace(/\s+/gu, ' ').trim()
}

describe('zero schema DDL', () => {
  it('creates app-side tables from the shared Zero schema columns and primary keys', () => {
    for (const [tableName, tableSchema] of Object.entries(schema.tables)) {
      const sql = normalizedSql(createZeroSchemaTableSql(tableName))
      const serverColumnNames = zeroSchemaServerColumnNamesByTable[tableName] ?? []

      expect(sql).toContain(`CREATE TABLE IF NOT EXISTS ${tableName} (`)
      for (const serverColumnName of serverColumnNames) {
        expect(sql, `${tableName} missing ${serverColumnName}`).toMatch(
          new RegExp(`(?:^| )${serverColumnName} (?:TEXT|BIGINT|BOOLEAN|JSONB)`, 'u'),
        )
      }
      const primaryKeyColumns = tableSchema.primaryKey.map((columnName) => {
        const column = tableSchema.columns[columnName]
        return column.serverName ?? columnName
      })
      expect(sql).toContain(`PRIMARY KEY (${primaryKeyColumns.join(', ')})`)
    }
  })

  it('keeps store-specific defaults and numeric compatibility as explicit overrides', () => {
    const sql = normalizedSql(
      createZeroSchemaTableSql('workbook_chat_thread', {
        columnOverrides: {
          scope: { defaultSql: "'private'" },
          executionPolicy: { defaultSql: "'autoApplyAll'" },
          entryCount: { defaultSql: '0' },
          reviewQueueItemCount: { defaultSql: '0' },
        },
      }),
    )

    expect(sql).toContain("scope TEXT NOT NULL DEFAULT 'private'")
    expect(sql).toContain("execution_policy TEXT NOT NULL DEFAULT 'autoApplyAll'")
    expect(sql).toContain('entry_count BIGINT NOT NULL DEFAULT 0')
    expect(sql).toContain('review_queue_item_count BIGINT NOT NULL DEFAULT 0')
  })

  it('rejects stale table and override names instead of silently generating drifted DDL', () => {
    expect(() => createZeroSchemaTableSql('missing_table')).toThrow('Unknown Zero schema table: missing_table')
    expect(() =>
      createZeroSchemaTableSql('workbook_chat_thread', {
        columnOverrides: {
          staleColumn: { defaultSql: '0' },
        },
      }),
    ).toThrow('Zero table workbook_chat_thread does not define column staleColumn')
  })
})
