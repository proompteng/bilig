import { ensureWorkbookDocumentExists } from '../workbook-document-initialization.js'
import { beforeEach, describe, expect, it, vi } from 'vitest'

const storeFns = vi.hoisted(() => ({
  applyAxisMetadataDiff: vi.fn(),
  applyCalculationSettings: vi.fn(),
  applyCellDiff: vi.fn(),
  applyDefinedNameDiff: vi.fn(),
  applyNumberFormatDiff: vi.fn(),
  applySheetDiff: vi.fn(),
  applyStyleDiff: vi.fn(),
  applyWorkbookMetadataDiff: vi.fn(),
  insertWorkbookHeaderIfMissing: vi.fn(),
}))

vi.mock('../store.js', () => ({
  applyAxisMetadataDiff: storeFns.applyAxisMetadataDiff,
  applyCalculationSettings: storeFns.applyCalculationSettings,
  applyCellDiff: storeFns.applyCellDiff,
  applyDefinedNameDiff: storeFns.applyDefinedNameDiff,
  applyNumberFormatDiff: storeFns.applyNumberFormatDiff,
  applySheetDiff: storeFns.applySheetDiff,
  applyStyleDiff: storeFns.applyStyleDiff,
  applyWorkbookMetadataDiff: storeFns.applyWorkbookMetadataDiff,
  insertWorkbookHeaderIfMissing: storeFns.insertWorkbookHeaderIfMissing,
}))
import type { QueryResultRow, Queryable } from '../store.js'

interface RecordedQuery {
  readonly text: string
  readonly values: readonly unknown[] | undefined
}

class FakeTransactionClient implements Queryable {
  readonly calls: RecordedQuery[] = []
  releaseCount = 0

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    return { rows: [] }
  }

  release(): void {
    this.releaseCount += 1
  }
}

class FakeTransactionalQueryable implements Queryable {
  readonly calls: RecordedQuery[] = []
  readonly client = new FakeTransactionClient()
  connectCount = 0

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    return { rows: [] as T[] }
  }

  async connect(): Promise<FakeTransactionClient> {
    this.connectCount += 1
    return this.client
  }
}

describe('workbook document initialization', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('skips projection replacement when the workbook already exists', async () => {
    storeFns.insertWorkbookHeaderIfMissing.mockResolvedValueOnce(false)
    const query = vi.fn()
    const db: Queryable = { query }

    await ensureWorkbookDocumentExists(db, 'book-1', 'owner-1')

    expect(storeFns.insertWorkbookHeaderIfMissing).toHaveBeenCalledOnce()
    expect(storeFns.applySheetDiff).not.toHaveBeenCalled()
    expect(query).not.toHaveBeenCalled()
  })

  it('creates missing workbooks atomically when the queryable supports transactions', async () => {
    storeFns.insertWorkbookHeaderIfMissing.mockResolvedValueOnce(true)
    const db = new FakeTransactionalQueryable()

    await ensureWorkbookDocumentExists(db, 'book-1', 'owner-1')

    expect(db.connectCount).toBe(1)
    expect(db.calls).toEqual([])
    expect(db.client.releaseCount).toBe(1)
    expect(db.client.calls[0]?.text).toBe('BEGIN')
    expect(db.client.calls.at(-1)?.text).toBe('COMMIT')
    expect(storeFns.insertWorkbookHeaderIfMissing).toHaveBeenCalledWith(db.client, 'book-1', expect.any(Object), expect.any(Object), null)
    expect(storeFns.applySheetDiff).toHaveBeenCalledWith(db.client, [], expect.any(Array))
    expect(db.client.calls.some((call) => call.text.includes('DELETE FROM sheets'))).toBe(true)
  })
})
