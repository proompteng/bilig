import { describe, expect, it } from 'vitest'
import { workbookColumnTileArgsSchema, workbookRowTileArgsSchema, workbookTileArgsSchema } from '../queries.js'

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
})
