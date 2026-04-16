import { describe, expect, it } from 'vitest'
import { normalizeWorkbookSheetIdAssignments, repairWorkbookSheetIds, type WorkbookSheetIdRow } from '../sheet-id-repair.js'
import type { QueryResultRow, Queryable } from '../store.js'

interface RecordedQuery {
  readonly text: string
  readonly values: readonly unknown[] | undefined
}

class FakeQueryable implements Queryable {
  readonly calls: RecordedQuery[] = []

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

describe('sheet-id-repair', () => {
  it('preserves unique existing sheet ids and repairs duplicates deterministically', () => {
    const assignments = normalizeWorkbookSheetIdAssignments([
      { workbookId: 'doc-1', name: 'Sheet1', sortOrder: 0, sheetId: 2 },
      { workbookId: 'doc-1', name: 'Sheet2', sortOrder: 1, sheetId: 2 },
      { workbookId: 'doc-1', name: 'Sheet3', sortOrder: 2, sheetId: null },
      { workbookId: 'doc-2', name: 'Sheet1', sortOrder: 0, sheetId: 3 },
    ] satisfies readonly WorkbookSheetIdRow[])

    expect(assignments).toEqual([
      {
        workbookId: 'doc-1',
        name: 'Sheet1',
        previousSheetId: 2,
        nextSheetId: 1,
      },
      {
        workbookId: 'doc-1',
        name: 'Sheet2',
        previousSheetId: 2,
        nextSheetId: 2,
      },
      {
        workbookId: 'doc-1',
        name: 'Sheet3',
        previousSheetId: null,
        nextSheetId: 3,
      },
      {
        workbookId: 'doc-2',
        name: 'Sheet1',
        previousSheetId: 3,
        nextSheetId: 3,
      },
    ])
  })

  it('repairs duplicate sheet ids and updates dependent tables by sheet name', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM sheets')
          ? [
              {
                workbookId: 'doc-1',
                name: 'Sheet1',
                sortOrder: 0,
                sheetId: 2,
              } satisfies QueryResultRow,
              {
                workbookId: 'doc-1',
                name: 'Sheet2',
                sortOrder: 1,
                sheetId: 2,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) => (text.includes('FROM pg_tables') ? [{ tableName: 'presence_coarse' } satisfies QueryResultRow] : null),
    ])

    await repairWorkbookSheetIds(queryable)

    expect(queryable.calls.some((call) => call.text.includes('UPDATE sheets'))).toBe(true)
    expect(
      queryable.calls.some(
        (call) =>
          call.text.includes('UPDATE sheets') && call.values?.[0] === 'doc-1' && call.values?.[1] === 'Sheet1' && call.values?.[2] === 1,
      ),
    ).toBe(true)
    expect(
      queryable.calls.some(
        (call) =>
          call.text.includes('UPDATE presence_coarse') &&
          call.values?.[0] === 'doc-1' &&
          call.values?.[1] === 'Sheet1' &&
          call.values?.[2] === 1,
      ),
    ).toBe(true)
    expect(queryable.calls.some((call) => call.text.includes('UPDATE workbook_change'))).toBe(false)
  })

  it('does nothing when sheet ids are already unique', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM sheets')
          ? [
              {
                workbookId: 'doc-1',
                name: 'Sheet1',
                sortOrder: 0,
                sheetId: 1,
              } satisfies QueryResultRow,
              {
                workbookId: 'doc-1',
                name: 'Sheet2',
                sortOrder: 1,
                sheetId: 2,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await repairWorkbookSheetIds(queryable)

    expect(queryable.calls).toHaveLength(1)
  })
})
