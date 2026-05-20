import { describe, expect, it } from 'vitest'
import {
  appendWorkbookChange,
  backfillWorkbookChanges,
  buildWorkbookChangeDescriptor,
  ensureWorkbookChangeSchema,
  listWorkbookChangesAfterRevision,
  listWorkbookChanges,
  loadLatestRedoableWorkbookChange,
  loadLatestUndoableWorkbookChange,
  loadWorkbookChange,
} from '../workbook-change-store.js'
import type { QueryResultRow, Queryable } from '../store.js'
import type { Row } from '@rocicorp/zero'

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
        return {
          rows: rows.filter((row): row is T => row !== null),
        }
      }
    }
    return { rows: [] }
  }

  async loadWorkbookChangeRow(input: { readonly documentId: string; readonly revision: number }): Promise<ZeroWorkbookChangeRow | null> {
    const result = await this.query(
      `
        FROM workbook_change
        WHERE workbook_id = $1 AND revision = $2
      `,
      [input.documentId, input.revision],
    )
    return result.rows[0] ? toZeroWorkbookChangeRow(result.rows[0]) : null
  }

  async listWorkbookChangesAfterRevisionRows(input: {
    readonly documentId: string
    readonly revision: number
  }): Promise<readonly ZeroWorkbookChangeRow[]> {
    const result = await this.query(
      `
        FROM workbook_change
        WHERE workbook_id = $1 AND revision > $2
        ORDER BY revision ASC
      `,
      [input.documentId, input.revision],
    )
    return result.rows.map(toZeroWorkbookChangeRow)
  }

  async listWorkbookHistoryRows(input: { readonly documentId: string }): Promise<readonly ZeroWorkbookChangeRow[]> {
    const result = await this.query(
      `
        FROM workbook_change
        WHERE workbook_id = $1
        ORDER BY revision ASC
      `,
      [input.documentId],
    )
    return result.rows.map(toZeroWorkbookChangeRow)
  }

  async listRecentWorkbookChangeRows(input: {
    readonly documentId: string
    readonly limit: number
  }): Promise<readonly ZeroWorkbookChangeRow[]> {
    const result = await this.query(
      `
        FROM workbook_change
        WHERE workbook_id = $1
        ORDER BY revision DESC
      `,
      [input.documentId, input.limit],
    )
    return result.rows.map(toZeroWorkbookChangeRow)
  }
}

type ZeroWorkbookChangeRow = Row['workbook_change']

type JsonValue = string | number | boolean | null | readonly JsonValue[] | { readonly [key: string]: JsonValue }

function isJsonValue(value: unknown): value is JsonValue {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return true
  }
  if (Array.isArray(value)) {
    return value.every((entry) => isJsonValue(entry))
  }
  if (typeof value !== 'object') {
    return false
  }
  return Object.values(value).every((entry) => isJsonValue(entry))
}

function stringValue(value: unknown, fallback: string): string {
  return typeof value === 'string' ? value : fallback
}

function optionalStringValue(value: unknown): string | null {
  return typeof value === 'string' ? value : null
}

function numberValue(value: unknown, fallback: number): number {
  return typeof value === 'number' ? value : fallback
}

function optionalNumberValue(value: unknown): number | null {
  return typeof value === 'number' ? value : null
}

function jsonValue(value: unknown): JsonValue | null {
  return isJsonValue(value) ? value : null
}

function toZeroWorkbookChangeRow(row: QueryResultRow): ZeroWorkbookChangeRow {
  return {
    workbookId: stringValue(row['workbookId'], 'doc-1'),
    revision: numberValue(row['revision'], -1),
    actorUserId: stringValue(row['actorUserId'], ''),
    clientMutationId: optionalStringValue(row['clientMutationId']),
    eventKind: stringValue(row['eventKind'], ''),
    summary: stringValue(row['summary'], ''),
    sheetId: optionalNumberValue(row['sheetId']),
    sheetName: optionalStringValue(row['sheetName']),
    anchorAddress: optionalStringValue(row['anchorAddress']),
    rangeJson: jsonValue(row['rangeJson']),
    undoBundleJson: jsonValue(row['undoBundleJson']),
    revertedByRevision: optionalNumberValue(row['revertedByRevision']),
    revertsRevision: optionalNumberValue(row['revertsRevision']),
    createdAt: numberValue(row['createdAt'] ?? row['createdAtUnixMs'], 0),
  }
}

class FakeWorkbookChangeStoreConnection extends FakeQueryable {
  readonly zeroChangeInputs: {
    readonly kind: 'one' | 'afterRevision' | 'history' | 'recent'
    readonly documentId: string
    readonly revision?: number
    readonly limit?: number
  }[] = []

  constructor(private readonly zeroChangeRows: readonly ZeroWorkbookChangeRow[]) {
    super()
  }

  async loadWorkbookChangeRow(input: { readonly documentId: string; readonly revision: number }) {
    this.zeroChangeInputs.push({
      kind: 'one',
      documentId: input.documentId,
      revision: input.revision,
    })
    return this.zeroChangeRows.find((row) => row.workbookId === input.documentId && row.revision === input.revision) ?? null
  }

  async listWorkbookChangesAfterRevisionRows(input: { readonly documentId: string; readonly revision: number }) {
    this.zeroChangeInputs.push({
      kind: 'afterRevision',
      documentId: input.documentId,
      revision: input.revision,
    })
    return this.zeroChangeRows
      .filter((row) => row.workbookId === input.documentId && row.revision > input.revision)
      .toSorted((left, right) => left.revision - right.revision)
  }

  async listWorkbookHistoryRows(input: { readonly documentId: string }) {
    this.zeroChangeInputs.push({
      kind: 'history',
      documentId: input.documentId,
    })
    return this.zeroChangeRows
      .filter((row) => row.workbookId === input.documentId)
      .toSorted((left, right) => left.revision - right.revision)
  }

  async listRecentWorkbookChangeRows(input: { readonly documentId: string; readonly limit: number }) {
    this.zeroChangeInputs.push({
      kind: 'recent',
      documentId: input.documentId,
      limit: input.limit,
    })
    return this.zeroChangeRows
      .filter((row) => row.workbookId === input.documentId)
      .toSorted((left, right) => right.revision - left.revision)
      .slice(0, input.limit)
  }
}

class FakeTransactionClient extends FakeQueryable {
  releaseCount = 0

  release(): void {
    this.releaseCount += 1
  }
}

class FakeTransactionalQueryable implements Queryable {
  readonly calls: RecordedQuery[] = []
  readonly client: FakeTransactionClient
  connectCount = 0

  constructor(responders: readonly ((text: string, values: readonly unknown[] | undefined) => QueryResultRow[] | null)[] = []) {
    this.client = new FakeTransactionClient(responders)
  }

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    return { rows: [] }
  }

  async connect(): Promise<FakeTransactionClient> {
    this.connectCount += 1
    return this.client
  }
}

class ConcurrentHistoryQueryable extends FakeQueryable {
  maxActiveInserts = 0
  private activeInserts = 0

  override async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    const isWorkbookChangeInsert = text.includes('INSERT INTO workbook_change')
    if (isWorkbookChangeInsert) {
      this.activeInserts += 1
      this.maxActiveInserts = Math.max(this.maxActiveInserts, this.activeInserts)
      await new Promise((resolve) => setTimeout(resolve, 0))
    }
    try {
      return await super.query<T>(text, values)
    } finally {
      if (isWorkbookChangeInsert) {
        this.activeInserts -= 1
      }
    }
  }
}

function latestQuery(queryable: FakeQueryable): RecordedQuery {
  const query = queryable.calls.at(-1)
  if (!query) {
    throw new Error('Expected at least one query')
  }
  return query
}

describe('workbook-change-store', () => {
  it('adds the client mutation id column for legacy history tables', async () => {
    const queryable = new FakeQueryable()

    await ensureWorkbookChangeSchema(queryable)

    expect(
      queryable.calls.some(
        (call) => call.text.includes('ALTER TABLE workbook_change') && call.text.includes('ADD COLUMN IF NOT EXISTS client_mutation_id'),
      ),
    ).toBe(true)
  })

  it('summarizes renderCommit cell bundles as authoritative range changes', () => {
    expect(
      buildWorkbookChangeDescriptor({
        kind: 'renderCommit',
        ops: [
          { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B2', value: 1 },
          { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'C4', formula: '=SUM(B2:B3)' },
          { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B4', value: 3 },
        ],
      }),
    ).toEqual({
      eventKind: 'renderCommit',
      summary: 'Updated 3 cells in Sheet1!B2:C4',
      sheetName: 'Sheet1',
      anchorAddress: 'B2',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C4',
      },
    })
  })

  it('marks structural row and column descriptors with Zero range scope', () => {
    expect(
      buildWorkbookChangeDescriptor({
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 2,
        count: 2,
      }),
    ).toMatchObject({
      eventKind: 'insertRows',
      anchorAddress: 'A3',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A3',
        endAddress: 'A4',
        scope: 'rows',
      },
    })

    expect(
      buildWorkbookChangeDescriptor({
        kind: 'deleteColumns',
        sheetName: 'Sheet1',
        start: 1,
        count: 2,
      }),
    ).toMatchObject({
      eventKind: 'deleteColumns',
      anchorAddress: 'B1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B1',
        endAddress: 'C1',
        scope: 'columns',
      },
    })
  })

  it('canonicalizes authored range descriptors through the shared Zero range authority', () => {
    expect(
      buildWorkbookChangeDescriptor({
        kind: 'clearRange',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'D5',
          endAddress: 'B2',
        },
      }),
    ).toEqual({
      eventKind: 'clearRange',
      summary: 'Cleared Sheet1!B2:D5',
      sheetName: 'Sheet1',
      anchorAddress: 'B2',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'D5',
      },
    })
  })

  it('keeps multi-sheet render commits conservative instead of lying about one sheet range', () => {
    expect(
      buildWorkbookChangeDescriptor({
        kind: 'renderCommit',
        ops: [
          { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 1 },
          { kind: 'upsertCell', sheetName: 'Sheet2', addr: 'B2', value: 2 },
        ],
      }),
    ).toEqual({
      eventKind: 'renderCommit',
      summary: 'Updated 2 cells across 2 sheets',
      sheetName: null,
      anchorAddress: null,
      range: null,
    })
  })

  it('preserves persisted structural range scope from the shared Zero model', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 16,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-16',
                eventKind: 'deleteColumns',
                summary: 'Deleted columns B:C on Sheet1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'B1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'C1', scope: 'columns' },
                undoBundleJson: null,
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 124_500,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(loadWorkbookChange(queryable, 'doc-1', 16)).resolves.toMatchObject({
      revision: 16,
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B1',
        endAddress: 'C1',
        scope: 'columns',
      },
    })
  })

  it('drops malformed persisted range scope so history conflict checks stay conservative', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 17,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-17',
                eventKind: 'insertRows',
                summary: 'Inserted rows 3:4 on Sheet1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A3',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'row-band' },
                undoBundleJson: null,
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 124_600,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(loadWorkbookChange(queryable, 'doc-1', 17)).resolves.toMatchObject({
      revision: 17,
      sheetName: 'Sheet1',
      anchorAddress: 'A3',
      range: null,
      rangeInvalid: true,
    })
  })

  it('drops malformed persisted range addresses so history conflict checks stay conservative', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 19,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-19',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A0', endAddress: 'A1' },
                undoBundleJson: null,
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 124_650,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(loadWorkbookChange(queryable, 'doc-1', 19)).resolves.toMatchObject({
      revision: 19,
      sheetName: 'Sheet1',
      anchorAddress: 'A1',
      range: null,
      rangeInvalid: true,
    })
  })

  it('drops workbook change rows with event kinds outside the shared Zero event model', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 18,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-18',
                eventKind: 'legacyPatch',
                summary: 'Legacy patch',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 124_700,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(loadWorkbookChange(queryable, 'doc-1', 18)).resolves.toBeNull()
  })

  it('drops workbook change rows with invalid revision trust metadata', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: '-1',
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-negative',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 124_800,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(loadWorkbookChange(queryable, 'doc-1', -1)).resolves.toBeNull()
  })

  it('records workbook changes with resolved sheet ids and serialized ranges', async () => {
    const queryable = new FakeQueryable([
      (text) => (text.includes('FROM sheets') ? [{ sheetId: 4, sheetName: 'Sheet1' } satisfies QueryResultRow] : null),
    ])

    await appendWorkbookChange(queryable, {
      documentId: 'doc-1',
      revision: 7,
      actorUserId: 'amy@example.com',
      clientMutationId: 'mutation-7',
      payload: {
        kind: 'fillRange',
        source: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A2',
        },
        target: {
          sheetName: 'Sheet1',
          startAddress: 'B1',
          endAddress: 'B2',
        },
      },
      undoBundle: {
        kind: 'engineOps',
        ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
      },
      createdAtUnixMs: 123_456,
    })

    const query = latestQuery(queryable)
    expect(query.text).toContain('INSERT INTO workbook_change')
    expect(query.values).toEqual([
      'doc-1',
      7,
      'amy@example.com',
      'mutation-7',
      'fillRange',
      'Filled Sheet1!B1:B2',
      4,
      'Sheet1',
      'B1',
      JSON.stringify({
        sheetName: 'Sheet1',
        startAddress: 'B1',
        endAddress: 'B2',
      }),
      JSON.stringify({
        kind: 'engineOps',
        ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
      }),
      null,
      null,
      123_456,
    ])
  })

  it('lists recent workbook changes in revision-descending order', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change') && text.includes('ORDER BY revision DESC')
          ? [
              {
                revision: 12,
                actorUserId: 'alex@example.com',
                clientMutationId: null,
                eventKind: 'applyAgentCommandBundle',
                summary: 'Applied workbook change set at revision r12',
                sheetId: 4,
                sheetName: 'Sheet1',
                anchorAddress: 'B2',
                rangeJson: {
                  sheetName: 'Sheet1',
                  startAddress: 'B2',
                  endAddress: 'C4',
                },
                undoBundleJson: null,
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 1_234,
              } satisfies QueryResultRow,
              {
                revision: 11,
                actorUserId: 'casey@example.com',
                clientMutationId: 'mutation-11',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 4,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: {
                  sheetName: 'Sheet1',
                  startAddress: 'A1',
                  endAddress: 'A1',
                },
                undoBundleJson: null,
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 1_111,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(listWorkbookChanges(queryable, { documentId: 'doc-1', limit: 2 })).resolves.toEqual([
      {
        revision: 12,
        actorUserId: 'alex@example.com',
        clientMutationId: null,
        eventKind: 'applyAgentCommandBundle',
        summary: 'Applied workbook change set at revision r12',
        sheetId: 4,
        sheetName: 'Sheet1',
        anchorAddress: 'B2',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C4',
        },
        rangeInvalid: false,
        undoBundle: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAtUnixMs: 1_234,
      },
      {
        revision: 11,
        actorUserId: 'casey@example.com',
        clientMutationId: 'mutation-11',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 4,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        rangeInvalid: false,
        undoBundle: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAtUnixMs: 1_111,
      },
    ])
    expect(latestQuery(queryable).values).toEqual(['doc-1', 2])
  })

  it('lists recent workbook changes through the shared Zero query model', async () => {
    const queryable = new FakeWorkbookChangeStoreConnection([
      {
        workbookId: 'doc-1',
        revision: 12,
        actorUserId: 'alex@example.com',
        clientMutationId: null,
        eventKind: 'applyAgentCommandBundle',
        summary: 'Applied workbook change set at revision r12',
        sheetId: 4,
        sheetName: 'Sheet1',
        anchorAddress: 'B2',
        rangeJson: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C4',
        },
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: 1_234,
      },
      {
        workbookId: 'doc-1',
        revision: 11,
        actorUserId: 'casey@example.com',
        clientMutationId: 'mutation-11',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 4,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: 1_111,
      },
    ])

    const changes = await listWorkbookChanges(queryable, { documentId: 'doc-1', limit: 2 })

    expect(queryable.zeroChangeInputs).toEqual([{ kind: 'recent', documentId: 'doc-1', limit: 2 }])
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_change'))).toBe(false)
    expect(changes.map((change) => [change.revision, change.createdAtUnixMs])).toEqual([
      [12, 1_234],
      [11, 1_111],
    ])
  })

  it('loads targeted and after-revision workbook history through the shared Zero query model', async () => {
    const queryable = new FakeWorkbookChangeStoreConnection([
      {
        workbookId: 'doc-1',
        revision: 7,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-7',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 4,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: 700,
      },
      {
        workbookId: 'doc-1',
        revision: 8,
        actorUserId: 'casey@example.com',
        clientMutationId: 'mutation-8',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!B1',
        sheetId: 4,
        sheetName: 'Sheet1',
        anchorAddress: 'B1',
        rangeJson: {
          sheetName: 'Sheet1',
          startAddress: 'B1',
          endAddress: 'B1',
        },
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: 800,
      },
    ])

    await expect(loadWorkbookChange(queryable, 'doc-1', 7)).resolves.toMatchObject({
      revision: 7,
      undoBundle: {
        kind: 'engineOps',
      },
    })
    await expect(
      listWorkbookChangesAfterRevision(queryable, {
        documentId: 'doc-1',
        revision: 7,
      }),
    ).resolves.toMatchObject([
      {
        revision: 8,
      },
    ])

    expect(queryable.zeroChangeInputs).toEqual([
      { kind: 'one', documentId: 'doc-1', revision: 7 },
      { kind: 'afterRevision', documentId: 'doc-1', revision: 7 },
    ])
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_change'))).toBe(false)
  })

  it('resolves undo and redo history from Zero-ordered workbook change rows', async () => {
    const queryable = new FakeWorkbookChangeStoreConnection([
      {
        workbookId: 'doc-1',
        revision: 20,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-20',
        eventKind: 'setCellValue',
        summary: 'Updated Sheet1!A1',
        sheetId: 4,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
        },
        revertedByRevision: 21,
        revertsRevision: null,
        createdAt: 2_000,
      },
      {
        workbookId: 'doc-1',
        revision: 21,
        actorUserId: 'alex@example.com',
        clientMutationId: 'mutation-21',
        eventKind: 'revertChange',
        summary: 'Reverted r20',
        sheetId: 4,
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        rangeJson: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        undoBundleJson: {
          kind: 'engineOps',
          ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 5 }],
        },
        revertedByRevision: null,
        revertsRevision: 20,
        createdAt: 2_100,
      },
    ])

    await expect(
      loadLatestUndoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()
    await expect(
      loadLatestRedoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toMatchObject({
      revision: 21,
    })

    expect(queryable.zeroChangeInputs).toEqual([
      { kind: 'history', documentId: 'doc-1' },
      { kind: 'history', documentId: 'doc-1' },
    ])
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_change'))).toBe(false)
  })

  it('backfills missing workbook_change rows from authoritative workbook events', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_event AS event')
          ? [
              {
                workbookId: 'doc-1',
                revision: 9,
                actorUserId: 'sam@example.com',
                clientMutationId: 'mutation-9',
                payload: {
                  kind: 'setCellFormula',
                  sheetName: 'Sheet1',
                  address: 'D5',
                  formula: '=SUM(A1:A4)',
                },
                createdAtUnixMs: 987_654,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) => (text.includes('FROM sheets') ? [{ sheetId: 11, sheetName: 'Sheet1' } satisfies QueryResultRow] : null),
    ])

    await backfillWorkbookChanges(queryable)

    const insertQuery = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_change'))
    expect(insertQuery?.values).toEqual([
      'doc-1',
      9,
      'sam@example.com',
      'mutation-9',
      'setCellFormula',
      'Set formula in Sheet1!D5',
      11,
      'Sheet1',
      'D5',
      JSON.stringify({
        sheetName: 'Sheet1',
        startAddress: 'D5',
        endAddress: 'D5',
      }),
      null,
      null,
      null,
      987_654,
    ])
  })

  it('backfills workbook_change rows sequentially so revert markers land after target rows', async () => {
    const queryable = new ConcurrentHistoryQueryable([
      (text) =>
        text.includes('FROM workbook_event AS event')
          ? [
              {
                workbookId: 'doc-1',
                revision: 7,
                actorUserId: 'sam@example.com',
                clientMutationId: 'mutation-7',
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'B1',
                  value: 1,
                },
                createdAtUnixMs: 987_000,
              } satisfies QueryResultRow,
              {
                workbookId: 'doc-1',
                revision: 8,
                actorUserId: 'sam@example.com',
                clientMutationId: 'mutation-8',
                payload: {
                  kind: 'revertChange',
                  targetRevision: 7,
                  targetSummary: 'Updated Sheet1!B1',
                  sheetName: 'Sheet1',
                  address: 'B1',
                  range: {
                    sheetName: 'Sheet1',
                    startAddress: 'B1',
                    endAddress: 'B1',
                  },
                  appliedBundle: {
                    kind: 'engineOps',
                    ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
                  },
                },
                createdAtUnixMs: 987_100,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) => (text.includes('FROM sheets') ? [{ sheetId: 11, sheetName: 'Sheet1' } satisfies QueryResultRow] : null),
    ])

    await backfillWorkbookChanges(queryable)

    const insertIndexes = queryable.calls
      .map((call, index) => ({ call, index }))
      .filter(({ call }) => call.text.includes('INSERT INTO workbook_change'))
    const markerIndex = queryable.calls.findIndex((call) => call.text.includes('UPDATE workbook_change') && call.values?.[1] === 7)
    expect(queryable.maxActiveInserts).toBe(1)
    expect(insertIndexes.map(({ call }) => call.values?.[1])).toEqual([7, 8])
    expect(markerIndex).toBeGreaterThan(insertIndexes[1]?.index ?? -1)
    expect(queryable.calls[markerIndex]?.values).toEqual(['doc-1', 7, 8])
  })

  it('skips invalid authoritative event revisions during workbook_change backfill', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_event AS event')
          ? [
              {
                workbookId: 'doc-1',
                revision: -1,
                actorUserId: 'sam@example.com',
                clientMutationId: 'mutation-negative',
                payload: {
                  kind: 'setCellFormula',
                  sheetName: 'Sheet1',
                  address: 'D5',
                  formula: '=SUM(A1:A4)',
                },
                createdAtUnixMs: 987_654,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await backfillWorkbookChanges(queryable)

    expect(queryable.calls.some((call) => call.text.includes('INSERT INTO workbook_change'))).toBe(false)
  })

  it('summarizes restoreVersion events as named workbook restores', () => {
    expect(
      buildWorkbookChangeDescriptor({
        kind: 'restoreVersion',
        versionId: 'version-1',
        versionName: 'Month close',
        sheetName: 'Sheet1',
        address: 'D5',
        snapshot: {
          version: 1,
          workbook: {
            name: 'doc-1',
          },
          sheets: [],
        },
      }),
    ).toEqual({
      eventKind: 'restoreVersion',
      summary: 'Restored version Month close',
      sheetName: 'Sheet1',
      anchorAddress: 'D5',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'D5',
        endAddress: 'D5',
      },
    })
  })

  it('summarizes structural metadata changes with layout labels', () => {
    expect(
      buildWorkbookChangeDescriptor({
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        startRow: 2,
        count: 2,
        height: 48,
        hidden: false,
      }),
    ).toEqual({
      eventKind: 'updateRowMetadata',
      summary: 'Updated rows 3:4 on Sheet1',
      sheetName: 'Sheet1',
      anchorAddress: 'A3',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A3',
        endAddress: 'A4',
        scope: 'rows',
      },
    })

    expect(
      buildWorkbookChangeDescriptor({
        kind: 'setFreezePane',
        sheetName: 'Sheet1',
        rows: 1,
        cols: 2,
      }),
    ).toEqual({
      eventKind: 'setFreezePane',
      summary: 'Set freeze panes on Sheet1',
      sheetName: 'Sheet1',
      anchorAddress: 'A1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
        scope: 'sheet',
      },
    })

    expect(
      buildWorkbookChangeDescriptor({
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 2,
        count: 2,
      }),
    ).toEqual({
      eventKind: 'insertRows',
      summary: 'Inserted rows 3:4 on Sheet1',
      sheetName: 'Sheet1',
      anchorAddress: 'A3',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A3',
        endAddress: 'A4',
        scope: 'rows',
      },
    })

    expect(
      buildWorkbookChangeDescriptor({
        kind: 'deleteColumns',
        sheetName: 'Sheet1',
        start: 1,
        count: 2,
      }),
    ).toEqual({
      eventKind: 'deleteColumns',
      summary: 'Deleted columns B:C on Sheet1',
      sheetName: 'Sheet1',
      anchorAddress: 'B1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B1',
        endAddress: 'C1',
        scope: 'columns',
      },
    })

    expect(
      buildWorkbookChangeDescriptor({
        kind: 'redoChange',
        targetRevision: 41,
        targetSummary: 'Updated Sheet1!A1',
        sheetName: 'Sheet1',
        address: 'A1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        appliedBundle: {
          kind: 'engineOps',
          ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }],
        },
      }),
    ).toEqual({
      eventKind: 'redoChange',
      summary: 'Redid r41: Updated Sheet1!A1',
      sheetName: 'Sheet1',
      anchorAddress: 'A1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    })
  })

  it('records revert changes with target metadata and marks the original revision as reverted', async () => {
    const queryable = new FakeQueryable([
      (text) => (text.includes('FROM sheets') ? [{ sheetId: 4, sheetName: 'Sheet1' } satisfies QueryResultRow] : null),
    ])

    await appendWorkbookChange(queryable, {
      documentId: 'doc-1',
      revision: 8,
      actorUserId: 'amy@example.com',
      clientMutationId: 'mutation-8',
      payload: {
        kind: 'revertChange',
        targetRevision: 7,
        targetSummary: 'Filled Sheet1!B1:B2',
        sheetName: 'Sheet1',
        address: 'B1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B1',
          endAddress: 'B2',
        },
        appliedBundle: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
        },
      },
      undoBundle: {
        kind: 'engineOps',
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 1 }],
      },
      createdAtUnixMs: 456_789,
    })

    const insertQuery = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_change'))
    expect(insertQuery?.values).toEqual([
      'doc-1',
      8,
      'amy@example.com',
      'mutation-8',
      'revertChange',
      'Reverted r7: Filled Sheet1!B1:B2',
      4,
      'Sheet1',
      'B1',
      JSON.stringify({
        sheetName: 'Sheet1',
        startAddress: 'B1',
        endAddress: 'B2',
      }),
      JSON.stringify({
        kind: 'engineOps',
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 1 }],
      }),
      null,
      7,
      456_789,
    ])

    const updateQuery = latestQuery(queryable)
    expect(updateQuery.text).toContain('UPDATE workbook_change')
    expect(updateQuery.values).toEqual(['doc-1', 7, 8])
  })

  it('persists revert changes atomically when the queryable supports transactions', async () => {
    const queryable = new FakeTransactionalQueryable([
      (text) => (text.includes('FROM sheets') ? [{ sheetId: 4, sheetName: 'Sheet1' } satisfies QueryResultRow] : null),
    ])

    await appendWorkbookChange(queryable, {
      documentId: 'doc-1',
      revision: 8,
      actorUserId: 'amy@example.com',
      clientMutationId: 'mutation-8',
      payload: {
        kind: 'revertChange',
        targetRevision: 7,
        targetSummary: 'Filled Sheet1!B1:B2',
        sheetName: 'Sheet1',
        address: 'B1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B1',
          endAddress: 'B2',
        },
        appliedBundle: {
          kind: 'engineOps',
          ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
        },
      },
      undoBundle: null,
      createdAtUnixMs: 456_789,
    })

    expect(queryable.connectCount).toBe(1)
    expect(queryable.calls).toEqual([])
    expect(queryable.client.releaseCount).toBe(1)
    expect(queryable.client.calls[0]?.text).toBe('BEGIN')
    expect(queryable.client.calls.at(-1)?.text).toBe('COMMIT')
    expect(queryable.client.calls.some((call) => call.text.includes('INSERT INTO workbook_change'))).toBe(true)
    expect(queryable.client.calls.some((call) => call.text.includes('UPDATE workbook_change') && call.values?.[2] === 8)).toBe(true)
  })

  it('does not wipe an existing revert marker during idempotent change upserts', async () => {
    const queryable = new FakeQueryable([
      (text) => (text.includes('FROM sheets') ? [{ sheetId: 4, sheetName: 'Sheet1' } satisfies QueryResultRow] : null),
    ])

    await appendWorkbookChange(queryable, {
      documentId: 'doc-1',
      revision: 7,
      actorUserId: 'amy@example.com',
      clientMutationId: 'mutation-7',
      payload: {
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'B1',
        value: 1,
      },
      undoBundle: null,
      createdAtUnixMs: 456_000,
    })

    const insertQuery = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_change'))
    expect(insertQuery?.text).toContain(
      'reverted_by_revision = COALESCE(EXCLUDED.reverted_by_revision, workbook_change.reverted_by_revision)',
    )
  })

  it('loads persisted undo metadata and revert markers for a workbook change row', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 12,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-12',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
                },
                revertedByRevision: 13,
                revertsRevision: null,
                createdAtUnixMs: 123_000,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(loadWorkbookChange(queryable, 'doc-1', 12)).resolves.toEqual({
      revision: 12,
      actorUserId: 'alex@example.com',
      clientMutationId: 'mutation-12',
      eventKind: 'setCellValue',
      summary: 'Updated Sheet1!A1',
      sheetId: 1,
      sheetName: 'Sheet1',
      anchorAddress: 'A1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      rangeInvalid: false,
      undoBundle: {
        kind: 'engineOps',
        ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
      },
      revertedByRevision: 13,
      revertsRevision: null,
      createdAtUnixMs: 123_000,
    })
  })

  it('loads persisted structural row and column history events', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 15,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-15',
                eventKind: 'insertRows',
                summary: 'Inserted rows 3:4 on Sheet1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A3',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4' },
                undoBundleJson: null,
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 124_000,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(loadWorkbookChange(queryable, 'doc-1', 15)).resolves.toMatchObject({
      revision: 15,
      eventKind: 'insertRows',
      summary: 'Inserted rows 3:4 on Sheet1',
      sheetName: 'Sheet1',
      anchorAddress: 'A3',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A3',
        endAddress: 'A4',
      },
    })
  })

  it('lists workbook changes after a target revision for stale revert checks', async () => {
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes('revision > $2') && values?.[0] === 'doc-1' && values?.[1] === 7
          ? [
              {
                revision: 8,
                actorUserId: 'morgan@example.com',
                clientMutationId: 'mutation-8',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 124_000,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(
      listWorkbookChangesAfterRevision(queryable, {
        documentId: 'doc-1',
        revision: 7,
      }),
    ).resolves.toMatchObject([
      {
        revision: 8,
        actorUserId: 'morgan@example.com',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
      },
    ])
  })

  it('loads latest undoable and redoable changes for an actor', async () => {
    const queryable = new FakeQueryable([
      (text) => {
        if (!text.includes('FROM workbook_change')) {
          return null
        }
        return [
          {
            revision: 14,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-14',
            eventKind: 'revertChange',
            summary: 'Reverted r13: Updated Sheet1!A1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'A1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 5 }],
            },
            revertedByRevision: null,
            revertsRevision: 13,
            createdAtUnixMs: 123_460,
          } satisfies QueryResultRow,
          {
            revision: 13,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-13',
            eventKind: 'setCellValue',
            summary: 'Updated Sheet1!A1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'A1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
            },
            revertedByRevision: null,
            revertsRevision: null,
            createdAtUnixMs: 123_456,
          } satisfies QueryResultRow,
        ]
      },
    ])

    await expect(
      loadLatestUndoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()

    await expect(
      loadLatestRedoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toMatchObject({
      revision: 14,
      eventKind: 'revertChange',
    })
    expect(latestQuery(queryable).values).toEqual(['doc-1'])
  })

  it('does not expose redo after a fresh authored change branches history after an undo', async () => {
    const queryable = new FakeQueryable([
      (text) => {
        if (!text.includes('FROM workbook_change')) {
          return null
        }
        if (text.includes("event_kind = 'revertChange'")) {
          return [
            {
              revision: 22,
              actorUserId: 'alex@example.com',
              clientMutationId: 'mutation-22',
              eventKind: 'revertChange',
              summary: 'Reverted r21: Updated Sheet1!A1',
              sheetId: 1,
              sheetName: 'Sheet1',
              anchorAddress: 'A1',
              rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
              undoBundleJson: {
                kind: 'engineOps',
                ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'seed' }],
              },
              revertedByRevision: null,
              revertsRevision: 21,
              createdAtUnixMs: 123_460,
            } satisfies QueryResultRow,
          ]
        }
        return [
          {
            revision: 23,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-23',
            eventKind: 'setCellValue',
            summary: 'Updated Sheet1!C1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'C1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'C1' }],
            },
            revertedByRevision: null,
            revertsRevision: null,
            createdAtUnixMs: 123_470,
          } satisfies QueryResultRow,
          {
            revision: 22,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-22',
            eventKind: 'revertChange',
            summary: 'Reverted r21: Updated Sheet1!A1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'A1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'seed' }],
            },
            revertedByRevision: null,
            revertsRevision: 21,
            createdAtUnixMs: 123_460,
          } satisfies QueryResultRow,
          {
            revision: 21,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-21',
            eventKind: 'setCellValue',
            summary: 'Updated Sheet1!A1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'A1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
            },
            revertedByRevision: 22,
            revertsRevision: null,
            createdAtUnixMs: 123_450,
          } satisfies QueryResultRow,
        ]
      },
    ])

    await expect(
      loadLatestRedoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()
  })

  it('does not expose undo after another actor changes an overlapping range', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 32,
                actorUserId: 'morgan@example.com',
                clientMutationId: 'mutation-32',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_480,
              } satisfies QueryResultRow,
              {
                revision: 31,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-31',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_470,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(
      loadLatestUndoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()
  })

  it('does not expose undo when a later malformed range would otherwise fall back to a disjoint anchor', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 52,
                actorUserId: 'morgan@example.com',
                clientMutationId: 'mutation-52',
                eventKind: 'insertRows',
                summary: 'Inserted rows 3:4 on Sheet1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A3',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'row-band' },
                undoBundleJson: null,
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_490,
              } satisfies QueryResultRow,
              {
                revision: 51,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-51',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!Z99',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'Z99',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'Z99', endAddress: 'Z99' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'Z99' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_480,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(
      loadLatestUndoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()
  })

  it('keeps redo after another actor changes a disjoint range', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 42,
                actorUserId: 'morgan@example.com',
                clientMutationId: 'mutation-42',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!C1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'C1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'C1' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_480,
              } satisfies QueryResultRow,
              {
                revision: 41,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-41',
                eventKind: 'revertChange',
                summary: 'Reverted r40: Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'seed' }],
                },
                revertedByRevision: null,
                revertsRevision: 40,
                createdAtUnixMs: 123_470,
              } satisfies QueryResultRow,
              {
                revision: 40,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-40',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
                },
                revertedByRevision: 41,
                revertsRevision: null,
                createdAtUnixMs: 123_460,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(
      loadLatestRedoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toMatchObject({
      revision: 41,
    })
  })
})
