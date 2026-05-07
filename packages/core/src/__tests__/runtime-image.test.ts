import { compileFormulaAst, parseFormula } from '@bilig/formula'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import type { EngineFormulaSourceRefs } from '../cell-mutations-at.js'
import { restoreWorkbookFromRuntimeImage, restoreWorkbookFromSnapshot } from '../snapshot/runtime-image.js'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'

function failOnCellWiseSnapshotRestoreNotification(): void {
  throw new Error('fresh snapshot restore should batch literal column notifications')
}

function collectFormulaSourceRefs(refs: EngineFormulaSourceRefs): Array<{
  cellIndex?: number
  col: number
  row: number
  sheetId: number
  source: string
}> {
  const collected: Array<{ cellIndex?: number; col: number; row: number; sheetId: number; source: string }> = []
  for (let index = 0; index < refs.length; index += 1) {
    const ref = Array.isArray(refs) ? refs[index]! : refs.at(index)
    collected.push({
      sheetId: ref.sheetId,
      cellIndex: ref.cellIndex,
      row: ref.row,
      col: ref.col,
      source: ref.source,
    })
  }
  return collected
}

describe('restoreWorkbookFromRuntimeImage', () => {
  it('prefers prepared template initialization when runtime image instances carry template ids', () => {
    const workbook = new WorkbookStore('runtime-image-test')
    const preparedCalls: Array<{ col: number; row: number; sheetId: number; source: string; templateId?: number }> = []
    const plainCalls: Array<{ mutation: { col: number; formula: string; kind: string; row: number }; sheetId: number }> = []
    const compiled = compileFormulaAst('A1+1', parseFormula('A1+1'))
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-test' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'B1', formula: 'A1+1' },
          ],
        },
      ],
    }

    restoreWorkbookFromRuntimeImage({
      snapshot,
      runtimeImage: {
        version: 1,
        templateBank: [
          {
            id: 7,
            templateKey: 'template:A1+1',
            baseSource: 'A1+1',
            baseRow: 0,
            baseCol: 1,
            compiled,
          },
        ],
        formulaInstances: [
          {
            cellIndex: 2,
            sheetName: 'Sheet1',
            row: 0,
            col: 1,
            source: 'A1+1',
            templateId: 7,
          },
        ],
        formulaValues: [
          {
            sheetName: 'Sheet1',
            row: 0,
            col: 1,
            value: { tag: ValueTag.Number, value: 2 },
          },
        ],
      },
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      resolveTemplateById: (templateId, source, row, col) =>
        templateId === 7
          ? {
              templateId,
              templateKey: 'template:A1+1',
              baseSource: source,
              compiled,
              translated: false,
              rowDelta: row,
              colDelta: col - 1,
            }
          : undefined,
      initializePreparedCellFormulasAt: (refs) => {
        preparedCalls.push(...refs)
      },
      initializeCellFormulasAt: (refs) => {
        plainCalls.push(...refs)
      },
    })

    const sheet = workbook.getSheet('Sheet1')
    expect(sheet).toBeDefined()
    const literalCellIndex = sheet!.grid.get(0, 0)
    expect(literalCellIndex).toBeGreaterThanOrEqual(0)
    expect(workbook.cellStore.tags[literalCellIndex]).toBe(ValueTag.Number)
    expect(workbook.cellStore.numbers[literalCellIndex]).toBe(1)
    expect(preparedCalls).toHaveLength(1)
    expect(preparedCalls[0]).toMatchObject({ sheetId: 1, row: 0, col: 1, source: 'A1+1', templateId: 7 })
    expect(plainCalls).toHaveLength(0)
  })

  it('prefers runtime-image coordinates over reparsing snapshot addresses during restore', () => {
    const workbook = new WorkbookStore('runtime-image-coords')
    const plainCalls: Array<{ mutation: { col: number; formula: string; kind: string; row: number }; sheetId: number }> = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-coords' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: '__ignored_literal__', value: 1 },
            { address: '__ignored_formula__', formula: 'A1+1' },
          ],
        },
      ],
    }

    restoreWorkbookFromRuntimeImage({
      snapshot,
      runtimeImage: {
        version: 1,
        templateBank: [],
        formulaInstances: [],
        formulaValues: [],
        sheetCells: [
          {
            sheetName: 'Sheet1',
            coords: [
              { row: 0, col: 0 },
              { row: 0, col: 1 },
            ],
          },
        ],
      },
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      initializeCellFormulasAt: (refs) => {
        plainCalls.push(...refs)
      },
    })

    const sheet = workbook.getSheet('Sheet1')
    expect(sheet).toBeDefined()
    expect(sheet!.grid.get(0, 0)).toBeGreaterThanOrEqual(0)
    expect(sheet!.grid.get(0, 1)).toBeGreaterThanOrEqual(0)
    const formulaCellIndex = sheet!.grid.get(0, 1)
    expect(plainCalls).toEqual([
      {
        sheetId: 1,
        cellIndex: formulaCellIndex,
        mutation: {
          kind: 'setCellFormula',
          row: 0,
          col: 1,
          formula: 'A1+1',
        },
      },
    ])
  })

  it('falls back to address-matched runtime formula values when value order is not aligned', () => {
    const workbook = new WorkbookStore('runtime-image-misaligned-values')
    const hydratedCalls: Array<{ col: number; row: number; sheetId: number; source: string; value: unknown }> = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-misaligned-values' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'B1', formula: 'A1+1' },
            { address: 'A2', value: 2 },
            { address: 'B2', formula: 'A2+1' },
          ],
        },
      ],
    }

    restoreWorkbookFromRuntimeImage({
      snapshot,
      runtimeImage: {
        version: 1,
        templateBank: [],
        formulaInstances: [
          {
            cellIndex: 2,
            sheetName: 'Sheet1',
            row: 0,
            col: 1,
            source: 'A1+1',
            templateId: 7,
          },
          {
            cellIndex: 4,
            sheetName: 'Sheet1',
            row: 1,
            col: 1,
            source: 'A2+1',
            templateId: 7,
          },
        ],
        formulaValues: [
          {
            sheetName: 'Sheet1',
            row: 1,
            col: 1,
            value: { tag: ValueTag.Number, value: 3 },
          },
          {
            sheetName: 'Sheet1',
            row: 0,
            col: 1,
            value: { tag: ValueTag.Number, value: 2 },
          },
        ],
      },
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      resolveTemplateById: (templateId, source, row, col) => ({
        templateId,
        templateKey: 'template:A1+1',
        baseSource: source,
        compiled: compileFormulaAst(source, parseFormula(source)),
        translated: false,
        rowDelta: row,
        colDelta: col,
      }),
      initializeHydratedPreparedCellFormulasAt: (refs) => {
        hydratedCalls.push(...refs)
      },
      initializeCellFormulasAt: () => {},
    })

    expect(hydratedCalls.map((call) => [call.row, call.col, call.value])).toEqual([
      [0, 1, { tag: ValueTag.Number, value: 2 }],
      [1, 1, { tag: ValueTag.Number, value: 3 }],
    ])
  })
})

describe('restoreWorkbookFromSnapshot', () => {
  it('preserves imported freeze pane scroll metadata during snapshot restore', () => {
    const workbook = new WorkbookStore('snapshot-freeze-pane-metadata')
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'snapshot-freeze-pane-metadata' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          metadata: {
            freezePane: { rows: 3, cols: 2, topLeftCell: 'I32', activePane: 'bottomRight' },
          },
          cells: [],
        },
      ],
    }

    restoreWorkbookFromSnapshot({
      snapshot,
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      initializeCellFormulasAt: () => {},
    })

    expect(workbook.getFreezePane('Sheet1')).toEqual({
      sheetName: 'Sheet1',
      rows: 3,
      cols: 2,
      topLeftCell: 'I32',
      activePane: 'bottomRight',
    })
  })

  it('uses fresh coordinate restore without reparsing address strings or cell-wise value notifications', () => {
    const workbook = new WorkbookStore('snapshot-coordinate-restore')
    const plainCalls: Array<{
      cellIndex?: number
      mutation: { col: number; formula: string; kind: string; row: number }
      sheetId: number
    }> = []
    workbook.cellStore.onSetValue = failOnCellWiseSnapshotRestoreNotification
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'snapshot-coordinate-restore' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: '__ignored_literal__', row: 9, col: 2, value: 42 },
            { address: '__ignored_formula__', row: 9, col: 3, formula: 'C10+1' },
          ],
        },
      ],
    }

    restoreWorkbookFromSnapshot({
      snapshot,
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      initializeCellFormulasAt: (refs) => {
        plainCalls.push(...refs)
      },
    })

    const sheet = workbook.getSheet('Sheet1')
    expect(sheet).toBeDefined()
    const literalCellIndex = sheet!.grid.get(9, 2)
    const formulaCellIndex = sheet!.grid.get(9, 3)
    expect(literalCellIndex).toBeGreaterThanOrEqual(0)
    expect(formulaCellIndex).toBeGreaterThanOrEqual(0)
    expect(workbook.cellStore.tags[literalCellIndex]).toBe(ValueTag.Number)
    expect(workbook.cellStore.numbers[literalCellIndex]).toBe(42)
    expect(plainCalls).toEqual([
      {
        sheetId: 1,
        cellIndex: formulaCellIndex,
        mutation: {
          kind: 'setCellFormula',
          row: 9,
          col: 3,
          formula: 'C10+1',
        },
      },
    ])
    expect(workbook.cellStore.onSetValue).toBe(failOnCellWiseSnapshotRestoreNotification)
  })

  it('reuses restored string ids for repeated snapshot literals', () => {
    const workbook = new WorkbookStore('snapshot-repeated-strings')
    const strings = new StringPool()
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'snapshot-repeated-strings' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'segment-1' },
            { address: 'A2', row: 1, col: 0, value: 'segment-1' },
            { address: 'A3', row: 2, col: 0, value: 'segment-2' },
          ],
        },
      ],
    }

    restoreWorkbookFromSnapshot({
      snapshot,
      workbook,
      strings,
      resetWorkbook: () => {},
      initializeCellFormulasAt: () => {},
    })

    const a1 = workbook.getCellIndex('Sheet1', 'A1')
    const a2 = workbook.getCellIndex('Sheet1', 'A2')
    const a3 = workbook.getCellIndex('Sheet1', 'A3')
    expect(a1).toBeDefined()
    expect(a2).toBeDefined()
    expect(a3).toBeDefined()
    expect(workbook.cellStore.stringIds[a1!]).toBe(workbook.cellStore.stringIds[a2!])
    expect(workbook.cellStore.stringIds[a3!]).not.toBe(workbook.cellStore.stringIds[a1!])
    expect(strings.size).toBe(3)
  })

  it('uses flat formula source refs during fresh snapshot restore when available', () => {
    const workbook = new WorkbookStore('snapshot-flat-formula-sources')
    const sourceCalls: Array<{ cellIndex?: number; col: number; row: number; sheetId: number; source: string }> = []
    const mutationCalls: unknown[] = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'snapshot-flat-formula-sources' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', row: 0, col: 0, value: 1 },
            { address: 'B1', row: 0, col: 1, formula: 'A1+1' },
          ],
        },
      ],
    }

    restoreWorkbookFromSnapshot({
      snapshot,
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      initializeCellFormulasAt: (refs) => {
        mutationCalls.push(...refs)
      },
      initializeFormulaSourcesAt: (refs) => {
        sourceCalls.push(...collectFormulaSourceRefs(refs))
      },
    })

    const formulaCellIndex = workbook.getCellIndex('Sheet1', 'B1')
    expect(formulaCellIndex).toBeDefined()
    expect(mutationCalls).toEqual([])
    expect(sourceCalls).toEqual([
      {
        sheetId: 1,
        cellIndex: formulaCellIndex,
        row: 0,
        col: 1,
        source: 'A1+1',
      },
    ])
  })
})
