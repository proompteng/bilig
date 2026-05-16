import { pathToFileURL } from 'node:url'

import { parseWorkPaperDocument } from '@bilig/headless'

import { createInMemoryWorkbookStorage, createWorkPaperRequestHandler, type WorkPaperJsonStorage } from './route.ts'

type Summary = {
  largestDeal: number
  totalRevenue: number
  westCustomers: number
}

type EditResponse = {
  after: Summary
  checks: {
    formulasPersisted: boolean
    serializedBytes: number
    totalRevenueChanged: boolean
  }
  records: number
}

type SummaryResponse = {
  summary: Summary
}

type InitialWorkbookFactory = () => Promise<string> | string

type DurableStorageOptions = {
  createInitialWorkbookJson?: InitialWorkbookFactory
}

type PostgresStorageOptions = DurableStorageOptions & {
  documentId: string
}

type SQLiteStorageOptions = DurableStorageOptions & {
  documentId: string
}

type KeyedStorageOptions = DurableStorageOptions & {
  key: string
}

type WorkPaperRequestHandler = (request: Request) => Promise<Response>

export type PostgresQueryResult<Row> = {
  rows: Row[]
}

export type PostgresWorkbookRow = {
  workbook_json: string
}

export type PostgresJsonbClient = {
  query(sql: string, values: readonly unknown[]): Promise<PostgresQueryResult<PostgresWorkbookRow>>
}

export type SQLiteWorkbookRow = {
  workbook_json: string
}

export type SQLiteClient = {
  get(sql: string, values: readonly unknown[]): Promise<SQLiteWorkbookRow | undefined> | SQLiteWorkbookRow | undefined
  run(sql: string, values: readonly unknown[]): Promise<void> | void
}

export type RedisTextClient = {
  get(key: string): Promise<null | string> | null | string
  set(key: string, value: string): Promise<unknown> | void
}

export type ObjectTextStore = {
  getText(key: string): Promise<null | string> | null | string
  putText(key: string, value: string, options: { contentType: string }): Promise<void> | void
}

export function createPostgresJsonbWorkPaperStorage(db: PostgresJsonbClient, options: PostgresStorageOptions): WorkPaperJsonStorage {
  return {
    async loadWorkbookJson() {
      const result = await db.query(
        `
          select workbook_json::text as workbook_json
          from workpaper_documents
          where id = $1
        `,
        [options.documentId],
      )

      const stored = result.rows[0]?.workbook_json
      if (stored === undefined) {
        return readInitialWorkbookJson(options.createInitialWorkbookJson)
      }
      return validateWorkbookJson(stored)
    },
    async saveWorkbookJson(workbookJson: string) {
      await db.query(
        `
          insert into workpaper_documents (id, workbook_json, updated_at)
          values ($1, $2::jsonb, now())
          on conflict (id) do update
            set workbook_json = excluded.workbook_json,
                updated_at = now()
        `,
        [options.documentId, validateWorkbookJson(workbookJson)],
      )
    },
  }
}

export function createSQLiteWorkPaperStorage(db: SQLiteClient, options: SQLiteStorageOptions): WorkPaperJsonStorage {
  return {
    async loadWorkbookJson() {
      const row = await db.get(
        `
          select workbook_json
          from workpaper_documents
          where id = ?
        `,
        [options.documentId],
      )

      if (row === undefined) {
        return readInitialWorkbookJson(options.createInitialWorkbookJson)
      }
      return validateWorkbookJson(row.workbook_json)
    },
    async saveWorkbookJson(workbookJson: string) {
      await db.run(
        `
          insert into workpaper_documents (id, workbook_json, updated_at)
          values (?, ?, current_timestamp)
          on conflict(id) do update
            set workbook_json = excluded.workbook_json,
                updated_at = current_timestamp
        `,
        [options.documentId, validateWorkbookJson(workbookJson)],
      )
    },
  }
}

export function createRedisWorkPaperStorage(redis: RedisTextClient, options: KeyedStorageOptions): WorkPaperJsonStorage {
  return {
    async loadWorkbookJson() {
      const stored = await redis.get(options.key)
      if (stored === null) {
        return readInitialWorkbookJson(options.createInitialWorkbookJson)
      }
      return validateWorkbookJson(stored)
    },
    async saveWorkbookJson(workbookJson: string) {
      await redis.set(options.key, validateWorkbookJson(workbookJson))
    },
  }
}

export function createObjectStoreWorkPaperStorage(store: ObjectTextStore, options: KeyedStorageOptions): WorkPaperJsonStorage {
  return {
    async loadWorkbookJson() {
      const stored = await store.getText(options.key)
      if (stored === null) {
        return readInitialWorkbookJson(options.createInitialWorkbookJson)
      }
      return validateWorkbookJson(stored)
    },
    async saveWorkbookJson(workbookJson: string) {
      await store.putText(options.key, validateWorkbookJson(workbookJson), {
        contentType: 'application/json; charset=utf-8',
      })
    },
  }
}

export async function createPersistenceAdapterDemoOutput() {
  const documentId = 'revenue-plan'
  const redisKey = 'workpapers:revenue-plan'
  const objectKey = 'workpapers/revenue-plan.json'

  const postgres = createInMemoryPostgresJsonbClient()
  const sqlite = createInMemorySQLiteClient()
  const redis = createInMemoryRedisTextClient()
  const objectStore = createInMemoryObjectTextStore()

  const postgresRun = await exerciseStorage('postgres-jsonb', () =>
    createPostgresJsonbWorkPaperStorage(postgres, {
      documentId,
    }),
  )
  const sqliteRun = await exerciseStorage('sqlite', () =>
    createSQLiteWorkPaperStorage(sqlite, {
      documentId,
    }),
  )
  const redisRun = await exerciseStorage('redis', () =>
    createRedisWorkPaperStorage(redis, {
      key: redisKey,
    }),
  )
  const objectStoreRun = await exerciseStorage('object-storage', () =>
    createObjectStoreWorkPaperStorage(objectStore, {
      key: objectKey,
    }),
  )

  const output = {
    adapters: ['postgres-jsonb', 'sqlite', 'redis', 'object-storage'],
    postgres: postgresRun,
    sqlite: sqliteRun,
    redis: redisRun,
    objectStorage: objectStoreRun,
    persistedBytes: {
      objectStorage: readByteLength(objectStore.getSavedText(objectKey), 'object store saved document'),
      postgres: readByteLength(postgres.getSavedWorkbookJson(documentId), 'postgres saved document'),
      redis: readByteLength(redis.getSavedText(redisKey), 'redis saved document'),
      sqlite: readByteLength(sqlite.getSavedWorkbookJson(documentId), 'sqlite saved document'),
    },
    verified: true,
  }

  assertOutput(output)
  return output
}

type PersistenceAdapterDemoOutput = Awaited<ReturnType<typeof createPersistenceAdapterDemoOutput>>
type StorageRun = PersistenceAdapterDemoOutput['postgres']

const updatedRevenueRecords = [
  { region: 'West', customers: 20, arpa: 1200 },
  { region: 'East', customers: 30, arpa: 250 },
  { region: 'Central', customers: 18, arpa: 300 },
  { region: 'North', customers: 65, arpa: 180 },
]

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  console.log(JSON.stringify(await createPersistenceAdapterDemoOutput(), null, 2))
}

async function exerciseStorage(label: string, createStorage: () => WorkPaperJsonStorage) {
  const writeHandler = createWorkPaperRequestHandler(createStorage())

  const before = await requestJson(writeHandler, 'GET', '/api/workpaper/summary', parseSummaryResponse, `${label} summary before`)
  const edit = await requestJson(writeHandler, 'POST', '/api/workpaper/revenue', parseEditResponse, `${label} revenue edit`, {
    records: updatedRevenueRecords,
  })
  const coldReadHandler = createWorkPaperRequestHandler(createStorage())
  const after = await requestJson(
    coldReadHandler,
    'GET',
    '/api/workpaper/summary',
    parseSummaryResponse,
    `${label} summary after cold restore`,
  )

  const output = {
    before: before.summary,
    edit: {
      after: edit.after,
      checks: edit.checks,
      records: edit.records,
    },
    after: after.summary,
    verified: true,
  }

  assertStorageRun(output, label)
  return output
}

async function requestJson<T>(
  handler: WorkPaperRequestHandler,
  method: string,
  path: string,
  parse: (value: unknown) => T,
  label: string,
  body?: unknown,
): Promise<T> {
  const request = new Request(`http://localhost:8787${path}`, {
    method,
    headers: body === undefined ? undefined : { 'content-type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body),
  })
  const response = await handler(request)
  const payload: unknown = await response.json()
  if (!response.ok) {
    throw new Error(`${label} failed: ${response.status} ${JSON.stringify(payload)}`)
  }
  return parse(payload)
}

function parseSummaryResponse(value: unknown): SummaryResponse {
  const record = readJsonRecord(value, 'summary response')
  return {
    summary: readSummary(record.summary, 'summary response summary'),
  }
}

function parseEditResponse(value: unknown): EditResponse {
  const record = readJsonRecord(value, 'edit response')
  const checks = readJsonRecord(record.checks, 'edit response checks')
  return {
    after: readSummary(record.after, 'edit response after'),
    checks: {
      formulasPersisted: readBoolean(checks.formulasPersisted, 'edit response formulasPersisted'),
      serializedBytes: readNumber(checks.serializedBytes, 'edit response serializedBytes'),
      totalRevenueChanged: readBoolean(checks.totalRevenueChanged, 'edit response totalRevenueChanged'),
    },
    records: readNumber(record.records, 'edit response records'),
  }
}

function readSummary(value: unknown, label: string): Summary {
  const record = readJsonRecord(value, label)
  return {
    largestDeal: readNumber(record.largestDeal, `${label} largestDeal`),
    totalRevenue: readNumber(record.totalRevenue, `${label} totalRevenue`),
    westCustomers: readNumber(record.westCustomers, `${label} westCustomers`),
  }
}

function readJsonRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isJsonRecord(value)) {
    throw new Error(`${label} must be an object`)
  }
  return value
}

function isJsonRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`${label} must be a number`)
  }
  return value
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`${label} must be a boolean`)
  }
  return value
}

async function readInitialWorkbookJson(createInitialWorkbookJson: InitialWorkbookFactory | undefined): Promise<string> {
  return validateWorkbookJson(await (createInitialWorkbookJson ?? createDefaultWorkbookJson)())
}

async function createDefaultWorkbookJson(): Promise<string> {
  return createInMemoryWorkbookStorage().loadWorkbookJson()
}

function validateWorkbookJson(workbookJson: string): string {
  parseWorkPaperDocument(workbookJson)
  return workbookJson
}

function readByteLength(value: string | undefined, label: string): number {
  if (value === undefined) {
    throw new Error(`${label} was not saved`)
  }
  validateWorkbookJson(value)
  return Buffer.byteLength(value, 'utf8')
}

function assertOutput(actual: PersistenceAdapterDemoOutput): void {
  if (
    actual.adapters.join(',') !== 'postgres-jsonb,sqlite,redis,object-storage' ||
    !actual.verified ||
    actual.persistedBytes.postgres <= 0 ||
    actual.persistedBytes.redis <= 0 ||
    actual.persistedBytes.sqlite <= 0 ||
    actual.persistedBytes.objectStorage <= 0
  ) {
    throw new Error(`unexpected persistence adapter output: ${JSON.stringify(actual)}`)
  }

  assertStorageRun(actual.postgres, 'postgres-jsonb')
  assertStorageRun(actual.sqlite, 'sqlite')
  assertStorageRun(actual.redis, 'redis')
  assertStorageRun(actual.objectStorage, 'object-storage')
}

function assertStorageRun(actual: StorageRun, label: string): void {
  const expectedBefore = {
    largestDeal: 24000,
    totalRevenue: 36900,
    westCustomers: 20,
  }
  const expectedAfter = {
    largestDeal: 24000,
    totalRevenue: 48600,
    westCustomers: 20,
  }

  if (
    JSON.stringify(actual.before) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(actual.edit.after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(actual.after) !== JSON.stringify(expectedAfter) ||
    actual.edit.records !== 4 ||
    !actual.edit.checks.totalRevenueChanged ||
    !actual.edit.checks.formulasPersisted ||
    actual.edit.checks.serializedBytes <= 0 ||
    !actual.verified
  ) {
    throw new Error(`unexpected ${label} storage run: ${JSON.stringify(actual)}`)
  }
}

type InMemoryPostgresJsonbClient = PostgresJsonbClient & {
  getSavedWorkbookJson(documentId: string): string | undefined
}

function createInMemoryPostgresJsonbClient(): InMemoryPostgresJsonbClient {
  const rows = new Map<string, string>()

  return {
    getSavedWorkbookJson(documentId: string) {
      return rows.get(documentId)
    },
    async query(sql: string, values: readonly unknown[]): Promise<PostgresQueryResult<PostgresWorkbookRow>> {
      const documentId = readString(values[0], 'document id')
      if (/^\s*select\b/i.test(sql)) {
        const workbookJson = rows.get(documentId)
        return {
          rows: workbookJson === undefined ? [] : [{ workbook_json: workbookJson }],
        }
      }

      const workbookJson = readString(values[1], 'workbook json')
      validateWorkbookJson(workbookJson)
      rows.set(documentId, workbookJson)
      return { rows: [] }
    },
  }
}

type InMemorySQLiteClient = SQLiteClient & {
  getSavedWorkbookJson(documentId: string): string | undefined
}

function createInMemorySQLiteClient(): InMemorySQLiteClient {
  const rows = new Map<string, string>()

  return {
    getSavedWorkbookJson(documentId: string) {
      return rows.get(documentId)
    },
    get(sql: string, values: readonly unknown[]): SQLiteWorkbookRow | undefined {
      if (!sql.includes('?')) {
        throw new Error('SQLite statements must use parameter placeholders')
      }
      const documentId = readString(values[0], 'document id')
      const workbookJson = rows.get(documentId)
      return workbookJson === undefined ? undefined : { workbook_json: workbookJson }
    },
    run(sql: string, values: readonly unknown[]) {
      if (!sql.includes('?')) {
        throw new Error('SQLite statements must use parameter placeholders')
      }
      const documentId = readString(values[0], 'document id')
      const workbookJson = readString(values[1], 'workbook json')
      validateWorkbookJson(workbookJson)
      rows.set(documentId, workbookJson)
    },
  }
}

type InMemoryRedisTextClient = RedisTextClient & {
  getSavedText(key: string): string | undefined
}

function createInMemoryRedisTextClient(): InMemoryRedisTextClient {
  const values = new Map<string, string>()
  return {
    get(key: string) {
      return values.get(key) ?? null
    },
    getSavedText(key: string) {
      return values.get(key)
    },
    set(key: string, value: string) {
      validateWorkbookJson(value)
      values.set(key, value)
    },
  }
}

type InMemoryObjectTextStore = ObjectTextStore & {
  getSavedText(key: string): string | undefined
}

function createInMemoryObjectTextStore(): InMemoryObjectTextStore {
  const objects = new Map<string, string>()
  return {
    getSavedText(key: string) {
      return objects.get(key)
    },
    getText(key: string) {
      return objects.get(key) ?? null
    },
    putText(key: string, value: string, options: { contentType: string }) {
      if (options.contentType !== 'application/json; charset=utf-8') {
        throw new Error(`unexpected content type: ${options.contentType}`)
      }
      validateWorkbookJson(value)
      objects.set(key, value)
    },
  }
}

function readString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`${label} must be a string`)
  }
  return value
}
