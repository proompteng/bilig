import { describe, expect, it } from 'vitest'
import { queries, workbookColumnTileArgsSchema, workbookRowTileArgsSchema, workbookTileArgsSchema } from '../queries.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

type ExecutableZeroQuery = ((...args: readonly never[]) => unknown) & {
  readonly fn: (input: { readonly args: unknown }) => unknown
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

  it('exposes browser-safe workbook agent live queries without server user context', () => {
    expect(() =>
      executeQueryRequestWithoutContext(
        queries.workbookChatThread.visibleByWorkbook({ documentId: 'doc-1', currentUserId: 'alex@example.com' }),
      ),
    ).not.toThrow()
    expect(() =>
      executeQueryRequestWithoutContext(
        queries.workbookWorkflowRun.visibleByThread({
          documentId: 'doc-1',
          threadId: 'thr-1',
          currentUserId: 'alex@example.com',
        }),
      ),
    ).not.toThrow()
  })
})
