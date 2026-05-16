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

  it('upserts referenced parent sheets during delta writes even when base.sheets omits them', async () => {
    const sqlite3 = await sqlite3InitModule()
    const db = new sqlite3.oo1.DB(':memory:', 'c')
    try {
      initializeWorkbookLocalStoreSchema(db)
      writeWorkbookAuthoritativeBase(db, createBase({ sheetId: 7, sheetName: 'Revenue', value: 11 }))
      writeWorkbookProjectionOverlay(db, createEmptyOverlay())

      const delta: WorkbookLocalAuthoritativeDelta = {
        replaceAll: false,
        replacedSheetIds: [7],
        base: {
          sheets: [],
          cellInputs: [
            {
              sheetId: 7,
              sheetName: 'Revenue',
              address: 'A1',
              rowNum: 0,
              colNum: 0,
              input: 22,
              formula: undefined,
              format: undefined,
            },
          ],
          cellRenders: [
            {
              sheetId: 7,
              sheetName: 'Revenue',
              address: 'A1',
              rowNum: 0,
              colNum: 0,
              value: { tag: 1, value: 22 },
              flags: 0,
              version: 2,
              styleId: undefined,
              numberFormatId: undefined,
            },
          ],
          rowAxisEntries: [],
          columnAxisEntries: [],
          styles: [],
        },
      }

      expect(() => writeWorkbookAuthoritativeDelta(db, delta)).not.toThrow()
      writeWorkbookProjectionOverlay(db, createEmptyOverlay())

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

  it('sanitizes persisted style JSON when reading viewport projections', async () => {
    const sqlite3 = await sqlite3InitModule()
    const db = new sqlite3.oo1.DB(':memory:', 'c')
    try {
      initializeWorkbookLocalStoreSchema(db)
      const base = createBase({ sheetId: 7, sheetName: 'Sheet1', value: 11 })
      base.cellRenders = [
        {
          ...base.cellRenders[0],
          styleId: 'style-1',
        },
      ]
      writeWorkbookAuthoritativeBase(db, base)

      const insertStyle = db.prepare('INSERT INTO authoritative_style (style_id, record_json) VALUES (?, ?)')
      try {
        insertStyle.bind([
          'style-1',
          JSON.stringify({
            fill: { backgroundColor: '#ffee00' },
            font: { family: 'Inter', size: Number.NaN, bold: 'yes', italic: true },
            alignment: { horizontal: 'diagonal', vertical: 'bottom', readingOrder: Number.POSITIVE_INFINITY },
            borders: {
              top: { style: 'solid', weight: 'thin', color: '#333333' },
              bottom: { style: 'wave', weight: 'thin', color: '#333333' },
            },
            protection: { locked: true, hidden: 0 },
            arbitrary: { trusted: false },
          }),
        ])
        insertStyle.step()
      } finally {
        insertStyle.finalize()
      }

      const projected = readWorkbookViewportProjection(db, 'Sheet1', {
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      })

      expect(projected?.cells[0]?.snapshot.styleId).toBe('style-1')
      expect(projected?.styles.find((style) => style.id === 'style-1')).toEqual({
        id: 'style-1',
        fill: { backgroundColor: '#ffee00' },
        font: { family: 'Inter', italic: true },
        alignment: { vertical: 'bottom' },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#333333' },
        },
        protection: { locked: true },
      })
    } finally {
      db.close()
    }
  })

  it('drops malformed persisted viewport cells and sanitizes parsed input values', async () => {
    const sqlite3 = await sqlite3InitModule()
    const db = new sqlite3.oo1.DB(':memory:', 'c')
    try {
      initializeWorkbookLocalStoreSchema(db)
      writeWorkbookAuthoritativeBase(db, createBase({ sheetId: 7, sheetName: 'Sheet1', value: 11 }))

      const insertRender = db.prepare(
        `
          INSERT INTO authoritative_cell_render (
            sheet_id,
            sheet_name,
            address,
            row_num,
            col_num,
            value_json,
            flags,
            version
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        `,
      )
      const insertInput = db.prepare(
        `
          INSERT INTO authoritative_cell_input (
            sheet_id,
            sheet_name,
            address,
            row_num,
            col_num,
            input_json
          )
          VALUES (?, ?, ?, ?, ?, ?)
        `,
      )
      try {
        insertRender.bind([7, 'Sheet1', 'B1', 0, 1, '{"tag":1,"value":1e999}', 0, 1])
        insertRender.step()
        insertRender.reset()
        insertRender.bind([7, 'Sheet1', 'C1', 0, 2, '{"tag":3,"value":"bad","stringId":1.5}', 0, 1])
        insertRender.step()
        insertRender.reset()
        insertRender.bind([7, 'Sheet1', 'D1', 0.5, 3, '{"tag":1,"value":44}', 0, 1])
        insertRender.step()
        insertRender.reset()
        insertRender.bind([7, 'Sheet1', 'E1', 0, 4, '{"tag":1,"value":55}', 0, 1])
        insertRender.step()

        insertInput.bind([7, 'Sheet1', 'E1', 0, 4, '1e999'])
        insertInput.step()
      } finally {
        insertRender.finalize()
        insertInput.finalize()
      }

      const projected = readWorkbookViewportProjection(db, 'Sheet1', {
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 4,
      })

      expect(projected?.cells.map((cell) => cell.snapshot.address)).toEqual(['A1', 'E1'])
      expect(projected?.cells.find((cell) => cell.snapshot.address === 'E1')?.snapshot).toMatchObject({
        address: 'E1',
        value: { tag: 1, value: 55 },
      })
      expect(projected?.cells.find((cell) => cell.snapshot.address === 'E1')?.snapshot.input).toBeUndefined()
    } finally {
      db.close()
    }
  })

  it('drops malformed persisted axis rows and omits invalid axis fields', async () => {
    const sqlite3 = await sqlite3InitModule()
    const db = new sqlite3.oo1.DB(':memory:', 'c')
    try {
      initializeWorkbookLocalStoreSchema(db)
      writeWorkbookAuthoritativeBase(db, createBase({ sheetId: 7, sheetName: 'Sheet1', value: 11 }))

      const insertRowAxis = db.prepare(
        `
          INSERT INTO authoritative_row_axis (
            sheet_id,
            sheet_name,
            axis_index,
            axis_id,
            size,
            hidden
          )
          VALUES (?, ?, ?, ?, ?, ?)
        `,
      )
      const insertColumnAxis = db.prepare(
        `
          INSERT INTO authoritative_column_axis (
            sheet_id,
            sheet_name,
            axis_index,
            axis_id,
            size,
            hidden
          )
          VALUES (?, ?, ?, ?, ?, ?)
        `,
      )
      try {
        insertRowAxis.bind([7, 'Sheet1', 0, 'row-0', 24, 1])
        insertRowAxis.step()
        insertRowAxis.reset()
        insertRowAxis.bind([7, 'Sheet1', 1.5, 'row-bad-index', 30, 1])
        insertRowAxis.step()
        insertRowAxis.reset()
        insertRowAxis.bind([7, 'Sheet1', 2, 'row-invalid-fields', -10, 2])
        insertRowAxis.step()

        insertColumnAxis.bind([7, 'Sheet1', 0, 'col-0', 96, 0])
        insertColumnAxis.step()
        insertColumnAxis.reset()
        insertColumnAxis.bind([7, 'Sheet1', 0.5, 'col-bad-index', 120, 1])
        insertColumnAxis.step()
        insertColumnAxis.reset()
        insertColumnAxis.bind([7, 'Sheet1', 2, 'col-invalid-fields', -1, 2])
        insertColumnAxis.step()
      } finally {
        insertRowAxis.finalize()
        insertColumnAxis.finalize()
      }

      const projected = readWorkbookViewportProjection(db, 'Sheet1', {
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      })

      expect(projected?.rowAxisEntries).toEqual([
        { id: 'row-0', index: 0, size: 24, hidden: true },
        { id: 'row-invalid-fields', index: 2 },
      ])
      expect(projected?.columnAxisEntries).toEqual([
        { id: 'col-0', index: 0, size: 96, hidden: false },
        { id: 'col-invalid-fields', index: 2 },
      ])
    } finally {
      db.close()
    }
  })

  it('sanitizes persisted sheet metadata when reading viewport projections', async () => {
    const sqlite3 = await sqlite3InitModule()
    const db = new sqlite3.oo1.DB(':memory:', 'c')
    try {
      initializeWorkbookLocalStoreSchema(db)
      writeWorkbookAuthoritativeBase(db, createBase({ sheetId: 7, sheetName: 'Sheet1', value: 11 }))
      db.exec("UPDATE authoritative_sheet SET freeze_rows = -1, freeze_cols = 1.5 WHERE name = 'Sheet1'")

      expect(
        readWorkbookViewportProjection(db, 'Sheet1', {
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        }),
      ).toMatchObject({
        sheetId: 7,
        sheetName: 'Sheet1',
        freezeRows: 0,
        freezeCols: 0,
      })

      db.exec("UPDATE authoritative_sheet SET sheet_id = 7.5 WHERE name = 'Sheet1'")
      expect(
        readWorkbookViewportProjection(db, 'Sheet1', {
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        }),
      ).toBeNull()
    } finally {
      db.close()
    }
  })

  it('rejects malformed viewport bounds before querying local projections', async () => {
    const sqlite3 = await sqlite3InitModule()
    const db = new sqlite3.oo1.DB(':memory:', 'c')
    try {
      initializeWorkbookLocalStoreSchema(db)
      writeWorkbookAuthoritativeBase(db, createBase({ sheetId: 7, sheetName: 'Sheet1', value: 11 }))

      expect(readWorkbookViewportProjection(db, 'Sheet1', { rowStart: -1, rowEnd: 0, colStart: 0, colEnd: 0 })).toBeNull()
      expect(readWorkbookViewportProjection(db, 'Sheet1', { rowStart: 0.5, rowEnd: 1, colStart: 0, colEnd: 0 })).toBeNull()
      expect(readWorkbookViewportProjection(db, 'Sheet1', { rowStart: 2, rowEnd: 1, colStart: 0, colEnd: 0 })).toBeNull()
      expect(
        readWorkbookViewportProjection(db, 'Sheet1', { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: Number.MAX_SAFE_INTEGER + 1 }),
      ).toBeNull()
    } finally {
      db.close()
    }
  })

  it('normalizes delta child rows to the canonical parent sheet name for each sheet id', async () => {
    const sqlite3 = await sqlite3InitModule()
    const db = new sqlite3.oo1.DB(':memory:', 'c')
    try {
      initializeWorkbookLocalStoreSchema(db)
      writeWorkbookAuthoritativeBase(db, createBase({ sheetId: 7, sheetName: 'Revenue', value: 11 }))
      writeWorkbookProjectionOverlay(db, createEmptyOverlay())

      const delta: WorkbookLocalAuthoritativeDelta = {
        replaceAll: false,
        replacedSheetIds: [7],
        base: {
          sheets: [
            {
              sheetId: 7,
              name: 'Revenue 2026',
              sortOrder: 0,
              freezeRows: 0,
              freezeCols: 0,
            },
          ],
          cellInputs: [
            {
              sheetId: 7,
              sheetName: 'Revenue',
              address: 'A1',
              rowNum: 0,
              colNum: 0,
              input: 22,
              formula: undefined,
              format: undefined,
            },
          ],
          cellRenders: [
            {
              sheetId: 7,
              sheetName: 'Revenue',
              address: 'A1',
              rowNum: 0,
              colNum: 0,
              value: { tag: 1, value: 22 },
              flags: 0,
              version: 2,
              styleId: undefined,
              numberFormatId: undefined,
            },
          ],
          rowAxisEntries: [],
          columnAxisEntries: [],
          styles: [],
        },
      }

      expect(() => writeWorkbookAuthoritativeDelta(db, delta)).not.toThrow()
      writeWorkbookProjectionOverlay(db, createEmptyOverlay())

      expect(
        readWorkbookViewportProjection(db, 'Revenue 2026', {
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        })?.cells[0]?.snapshot.value,
      ).toEqual({ tag: 1, value: 22 })
      expect(
        readWorkbookViewportProjection(db, 'Revenue', {
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        }),
      ).toBeNull()
    } finally {
      db.close()
    }
  })
})
