import type { Queryable } from './store.js'

export const DEFAULT_ZERO_PUBLICATION = 'zero_data_v2'

export const ZERO_PUBLICATION_TABLES = [
  'workbooks',
  'sheets',
  'cell_styles',
  'cell_number_formats',
  'cells',
  'row_metadata',
  'column_metadata',
  'cell_eval',
  'defined_names',
  'presence_coarse',
  'workbook_change',
  'workbook_chat_thread',
  'workbook_workflow_run',
] as const

const POSTGRES_IDENTIFIER_PATTERN = /^[A-Za-z_][A-Za-z0-9_]*$/

function quoteIdentifier(identifier: string): string {
  return `"${identifier.replaceAll('"', '""')}"`
}

function formatQualifiedTable(tableName: string): string {
  return `public.${quoteIdentifier(tableName)}`
}

function formatQualifiedTableList(tableNames: readonly string[]): string {
  return tableNames.map((tableName) => formatQualifiedTable(tableName)).join(', ')
}

function parsePublicationTableRows(rows: readonly { tableName?: unknown }[]): ReadonlySet<string> {
  return new Set(rows.flatMap((row) => (typeof row.tableName === 'string' && row.tableName.length > 0 ? [row.tableName] : [])))
}

export function resolveZeroPublicationName(env: Record<string, string | undefined> = process.env): string {
  const publication = env['BILIG_ZERO_PUBLICATION']?.trim() || DEFAULT_ZERO_PUBLICATION
  if (!POSTGRES_IDENTIFIER_PATTERN.test(publication)) {
    throw new Error(`Invalid Zero publication name: ${publication}`)
  }
  return publication
}

async function publicationExists(db: Queryable, publicationName: string): Promise<boolean> {
  const result = await db.query<{ present?: unknown }>(
    `
      SELECT 1 AS present
      FROM pg_publication
      WHERE pubname = $1
      LIMIT 1
    `,
    [publicationName],
  )
  return result.rows.length > 0
}

async function loadPublicationTables(db: Queryable, publicationName: string): Promise<ReadonlySet<string>> {
  const result = await db.query<{ tableName?: unknown }>(
    `
      SELECT tablename AS "tableName"
      FROM pg_publication_tables
      WHERE pubname = $1
        AND schemaname = 'public'
    `,
    [publicationName],
  )
  return parsePublicationTableRows(result.rows)
}

export async function ensureZeroPublication(db: Queryable, publicationName = resolveZeroPublicationName()): Promise<void> {
  const quotedPublicationName = quoteIdentifier(publicationName)
  if (!(await publicationExists(db, publicationName))) {
    await db.query(`CREATE PUBLICATION ${quotedPublicationName} FOR TABLE ${formatQualifiedTableList(ZERO_PUBLICATION_TABLES)}`)
    return
  }

  const existingTables = await loadPublicationTables(db, publicationName)
  const missingTables = ZERO_PUBLICATION_TABLES.filter((tableName) => !existingTables.has(tableName))
  if (missingTables.length === 0) {
    return
  }

  await db.query(`ALTER PUBLICATION ${quotedPublicationName} ADD TABLE ${formatQualifiedTableList(missingTables)}`)
}
