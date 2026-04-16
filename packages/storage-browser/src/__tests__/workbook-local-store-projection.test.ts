import sqlite3InitModule from '@sqlite.org/sqlite-wasm'
import { describe, expect, it } from 'vitest'

import {
  readWorkbookViewportProjection,
  writeWorkbookAuthoritativeBase,
  writeWorkbookAuthoritativeDelta,
  writeWorkbookProjectionOverlay,
} from '../workbook-local-store-projection.js'
import { initializeWorkbookLocalStoreSchema } from '../workbook-local-store-schema.js'
import type {
  WorkbookLocalAuthoritativeBase,
  WorkbookLocalAuthoritativeDelta,
  WorkbookLocalProjectionOverlay,
} from '../workbook-local-base.js'

function createBase(input: { sheetId: number; sheetName: string; value: number }): WorkbookLocalAuthoritativeBase {
  const { sheetId, sheetName, value } = input
  return {
    sheets: [
      {
        sheetId,
        name: sheetName,
        sortOrder: 0,
        freezeRows: 0,
        freezeCols: 0,
      },
    ],
    cellInputs: [
      {
        sheetId,
        sheetName,
        address: 'A1',
        rowNum: 0,
        colNum: 0,
        input: value,
        formula: undefined,
        format: undefined,
      },
    ],
    cellRenders: [
      {
        sheetId,
        sheetName,
        address: 'A1',
        rowNum: 0,
        colNum: 0,
        value: { tag: 1, value },
        flags: 0,
        version: 1,
        styleId: undefined,
        numberFormatId: undefined,
      },
    ],
    rowAxisEntries: [],
    columnAxisEntries: [],
    styles: [],
  }
}

function createEmptyOverlay(): WorkbookLocalProjectionOverlay {
  return {
    cells: [],
    rowAxisEntries: [],
    columnAxisEntries: [],
    styles: [],
  }
}

describe('workbook-local-store projection', () => {
  it('keeps renamed sheets addressable through the same sheet id', async () => {
    const sqlite3 = await sqlite3InitModule()
    const db = new sqlite3.oo1.DB(':memory:', 'c')
    try {
      initializeWorkbookLocalStoreSchema(db)
      writeWorkbookAuthoritativeBase(db, createBase({ sheetId: 7, sheetName: 'Sheet1', value: 11 }))
      writeWorkbookProjectionOverlay(db, createEmptyOverlay())

      const delta: WorkbookLocalAuthoritativeDelta = {
        replaceAll: false,
        replacedSheetIds: [7],
        base: createBase({ sheetId: 7, sheetName: 'Revenue', value: 22 }),
      }
      writeWorkbookAuthoritativeDelta(db, delta)
      writeWorkbookProjectionOverlay(db, createEmptyOverlay())

      expect(
        readWorkbookViewportProjection(db, 'Sheet1', {
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        }),
      ).toBeNull()

      expect(
        readWorkbookViewportProjection(db, 'Revenue', {
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        }),
      ).toMatchObject({
        sheetName: 'Revenue',
        cells: [
          {
            row: 0,
            col: 0,
            snapshot: {
              sheetName: 'Revenue',
              address: 'A1',
              version: 1,
            },
          },
        ],
      })
      expect(
        readWorkbookViewportProjection(db, 'Revenue', {
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        })?.cells[0]?.snapshot.value,
      ).toEqual({ tag: 1, value: 22 })
    } finally {
      db.close()
    }
  })
})
