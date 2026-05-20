import { describe, expect, it } from 'vitest'
import { executeZeroQueryTransform, zeroQueryTransformNames } from '../query-transforms.js'
import { queries, workbookColumnTileArgsSchema, workbookRowTileArgsSchema, workbookTileArgsSchema } from '../queries.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

type ExecutableZeroQuery = ((...args: readonly never[]) => unknown) & {
  readonly fn: (input: { readonly args: unknown; readonly ctx?: { readonly userID: string } }) => unknown
}

function isExecutableZeroQuery(value: unknown): value is ExecutableZeroQuery {
  return typeof value === 'function' && 'fn' in value && typeof value.fn === 'function'
}

function executeQueryRequestWithoutContext(request: unknown): unknown {
  if (!isRecord(request) || !isExecutableZeroQuery(request['query'])) {
    throw new Error('Expected a Zero query request')
  }
  return request['query'].fn({ args: request['args'] })
}

function executeQueryRequestWithUser(request: unknown, userID: string): unknown {
  if (!isRecord(request) || !isExecutableZeroQuery(request['query'])) {
    throw new Error('Expected a Zero query request')
  }
  return request['query'].fn({ args: request['args'], ctx: { userID } })
}

function queryRequestArgs(request: unknown): unknown {
  if (!isRecord(request)) {
    throw new Error('Expected a Zero query request')
  }
  return request['args']
}

function flattenQueryAliases(value: unknown, prefix = ''): string[] {
  if (isExecutableZeroQuery(value)) {
    return [prefix]
  }
  if (!isRecord(value)) {
    return []
  }
  return Object.entries(value).flatMap(([key, child]) => flattenQueryAliases(child, prefix ? `${prefix}.${key}` : key))
}

function queryArgsForAlias(alias: string): unknown {
  const workbookArgs = { documentId: 'doc-1' }
  if (alias === 'workbookChange.one' || alias === 'workbookChanges.one') {
    return { ...workbookArgs, revision: 1 }
  }
  if (alias === 'workbookChange.afterRevision' || alias === 'workbookChanges.afterRevision') {
    return { ...workbookArgs, revision: 0 }
  }
  if (alias === 'workbookChange.byWorkbook' || alias === 'workbookChanges.byWorkbook') {
    return { ...workbookArgs, limit: 10 }
  }
  if (alias.endsWith('.one')) {
    return { ...workbookArgs, sheetName: 'Sheet1', address: 'A1' }
  }
  if (alias.endsWith('.tile')) {
    if (alias.includes('Row') || alias.includes('rowMetadata')) {
      return { ...workbookArgs, sheetName: 'Sheet1', rowStart: 0, rowEnd: 2 }
    }
    if (alias.includes('Col') || alias.includes('columnMetadata')) {
      return { ...workbookArgs, sheetName: 'Sheet1', colStart: 0, colEnd: 2 }
    }
    return { ...workbookArgs, sheetName: 'Sheet1', rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 2 }
  }
  if (alias.endsWith('.byThread') || alias.endsWith('.visibleByThread')) {
    return { ...workbookArgs, threadId: 'thr-1' }
  }
  return workbookArgs
}

describe('zero sync query schemas', () => {
  it('accepts valid workbook tile ranges', () => {
    expect(
      workbookTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      }).success,
    ).toBe(true)
    expect(
      workbookTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetId: 7,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      }).success,
    ).toBe(true)
  })

  it('rejects ambiguous or fake sheet id routing', () => {
    expect(
      workbookTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetId: '7',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      }).success,
    ).toBe(false)
    expect(
      workbookTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      }).success,
    ).toBe(false)
  })

  it('rejects unsafe workbook tile range integers', () => {
    const unsafe = Number.MAX_SAFE_INTEGER + 1

    expect(
      workbookTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        rowStart: unsafe,
        rowEnd: unsafe,
        colStart: 0,
        colEnd: 0,
      }).success,
    ).toBe(false)
    expect(
      workbookRowTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: unsafe,
      }).success,
    ).toBe(false)
    expect(
      workbookColumnTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        colStart: 0,
        colEnd: unsafe,
      }).success,
    ).toBe(false)
  })

  it('rejects inverted workbook tile ranges', () => {
    expect(
      workbookTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        rowStart: 2,
        rowEnd: 1,
        colStart: 0,
        colEnd: 0,
      }).success,
    ).toBe(false)
    expect(
      workbookRowTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        rowStart: 2,
        rowEnd: 1,
      }).success,
    ).toBe(false)
    expect(
      workbookColumnTileArgsSchema.safeParse({
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        colStart: 2,
        colEnd: 1,
      }).success,
    ).toBe(false)
  })

  it('requires authenticated user context for workbook agent visible queries', () => {
    expect(() => executeQueryRequestWithoutContext(queries.workbookChatThread.visibleByWorkbook({ documentId: 'doc-1' }))).toThrow()
    expect(() =>
      executeQueryRequestWithoutContext(
        queries.workbookWorkflowRun.visibleByThread({
          documentId: 'doc-1',
          threadId: 'thr-1',
        }),
      ),
    ).toThrow()
    expect(() =>
      executeQueryRequestWithoutContext(
        queries.workbookAgentRun.visibleByThread({
          documentId: 'doc-1',
          threadId: 'thr-1',
        }),
      ),
    ).toThrow()
  })

  it('builds workbook agent visible queries from authenticated user context', () => {
    expect(queryRequestArgs(queries.workbookChatThread.visibleByWorkbook({ documentId: 'doc-1' }))).toEqual({
      documentId: 'doc-1',
    })
    expect(
      queryRequestArgs(
        queries.workbookWorkflowRun.visibleByThread({
          documentId: 'doc-1',
          threadId: 'thr-1',
        }),
      ),
    ).toEqual({
      documentId: 'doc-1',
      threadId: 'thr-1',
    })
    expect(() =>
      executeQueryRequestWithUser(queries.workbookChatThread.visibleByWorkbook({ documentId: 'doc-1' }), 'alex@example.com'),
    ).not.toThrow()
    expect(() =>
      executeQueryRequestWithUser(
        queries.workbookWorkflowRun.visibleByThread({
          documentId: 'doc-1',
          threadId: 'thr-1',
        }),
        'alex@example.com',
      ),
    ).not.toThrow()
    expect(() =>
      executeQueryRequestWithUser(
        queries.workbookAgentRun.visibleByThread({
          documentId: 'doc-1',
          threadId: 'thr-1',
        }),
        'alex@example.com',
      ),
    ).not.toThrow()
  })

  it('exports one shared transform for every query alias', () => {
    const queryAliases = flattenQueryAliases(queries).toSorted()

    expect(zeroQueryTransformNames).toEqual(queryAliases)
    for (const alias of queryAliases) {
      expect(() => executeZeroQueryTransform(alias, queryArgsForAlias(alias), 'alex@example.com'), alias).not.toThrow()
    }
  })

  it('runs query transforms through Zero query validators instead of a duplicate schema registry', () => {
    expect(() =>
      executeZeroQueryTransform(
        'cellInput.tile',
        {
          documentId: 'doc-1',
          sheetName: 'Sheet1',
          rowStart: Number.MAX_SAFE_INTEGER + 1,
          rowEnd: Number.MAX_SAFE_INTEGER + 1,
          colStart: 0,
          colEnd: 0,
        },
        'alex@example.com',
      ),
    ).toThrow()
    expect(() => executeZeroQueryTransform('workbookChatThread', { documentId: 'doc-1' }, 'alex@example.com')).toThrow(
      'Unknown Zero query: workbookChatThread',
    )
    expect(() => executeZeroQueryTransform('missing.query', { documentId: 'doc-1' }, 'alex@example.com')).toThrow(
      'Unknown Zero query: missing.query',
    )
  })
})
