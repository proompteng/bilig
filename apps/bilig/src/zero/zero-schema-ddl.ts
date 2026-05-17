import { schema } from '@bilig/zero-sync'
import type { Queryable } from './store.js'

type ZeroSchemaColumnType = 'string' | 'number' | 'boolean' | 'json'

interface ZeroSchemaColumn {
  readonly type: ZeroSchemaColumnType
  readonly optional: boolean
  readonly serverName?: string | undefined
}

interface ZeroSchemaTable {
  readonly columns: Readonly<Record<string, ZeroSchemaColumn>>
  readonly primaryKey: readonly string[]
}

interface ZeroColumnDdlOverride {
  readonly dataType?: string
  readonly defaultSql?: string
}

interface ZeroTableDdlOptions {
  readonly columnOverrides?: Readonly<Record<string, ZeroColumnDdlOverride>>
}

function quoteSqlIdentifier(identifier: string): string {
  return identifier
}

function columnDataType(type: ZeroSchemaColumnType): string {
  switch (type) {
    case 'string':
      return 'TEXT'
    case 'number':
      return 'BIGINT'
    case 'boolean':
      return 'BOOLEAN'
    case 'json':
      return 'JSONB'
    default:
      throw new Error(`Unsupported Zero column type: ${String(type)}`)
  }
}

function serverColumnName(columnName: string, column: { readonly serverName?: string | undefined }): string {
  return column.serverName ?? columnName
}

function readZeroSchemaTable(tableName: string): ZeroSchemaTable {
  for (const [candidateName, table] of Object.entries(schema.tables)) {
    if (candidateName !== tableName) {
      continue
    }
    const columns: Record<string, ZeroSchemaColumn> = {}
    for (const [columnName, column] of Object.entries(table.columns)) {
      columns[columnName] =
        column.serverName === undefined
          ? { type: column.type, optional: column.optional }
          : { type: column.type, optional: column.optional, serverName: column.serverName }
    }
    return {
      columns,
      primaryKey: table.primaryKey,
    }
  }
  throw new Error(`Unknown Zero schema table: ${tableName}`)
}

function requireZeroColumn(tableName: string, columnName: string, columns: Readonly<Record<string, ZeroSchemaColumn>>): ZeroSchemaColumn {
  const column = columns[columnName]
  if (!column) {
    throw new Error(`Zero table ${tableName} does not define column ${columnName}`)
  }
  return column
}

export function createZeroSchemaTableSql(tableName: string, options: ZeroTableDdlOptions = {}): string {
  const tableSchema = readZeroSchemaTable(tableName)
  for (const columnName of Object.keys(options.columnOverrides ?? {})) {
    requireZeroColumn(tableName, columnName, tableSchema.columns)
  }
  const primaryKeyColumns = new Set(tableSchema.primaryKey)
  const columnLines = Object.entries(tableSchema.columns).map(([columnName, column]) => {
    const override = options.columnOverrides?.[columnName]
    const sqlName = quoteSqlIdentifier(serverColumnName(columnName, column))
    const sqlType = override?.dataType ?? columnDataType(column.type)
    const defaultSql = override?.defaultSql === undefined ? '' : ` DEFAULT ${override.defaultSql}`
    const notNull = column.optional || primaryKeyColumns.has(columnName) ? '' : ' NOT NULL'
    return `      ${sqlName} ${sqlType}${notNull}${defaultSql}`
  })
  const primaryKeyLine = `      PRIMARY KEY (${tableSchema.primaryKey.map((columnName) => quoteSqlIdentifier(serverColumnName(columnName, requireZeroColumn(tableName, columnName, tableSchema.columns)))).join(', ')})`
  return `
    CREATE TABLE IF NOT EXISTS ${quoteSqlIdentifier(tableName)} (
${[...columnLines, primaryKeyLine].join(',\n')}
    )
  `
}

export async function ensureZeroSchemaTable(db: Queryable, tableName: string, options: ZeroTableDdlOptions = {}): Promise<void> {
  await db.query(createZeroSchemaTableSql(tableName, options))
}
