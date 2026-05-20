import { compileFormulaAst, parseFormula } from '@bilig/formula'
import { ErrorCode, ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'
import type { EngineFormulaSourceRefs } from '../cell-mutations-at.js'
import { CellFlags } from '../cell-store.js'
import {
  restoreWorkbookFromRuntimeImage,
  restoreWorkbookFromSnapshot,
  type CachedRuntimeFormulaRef,
  type HydratedPreparedRuntimeFormulaRef,
} from '../snapshot/runtime-image.js'
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

function collectHydratedPreparedRefs(
  refs:
    | readonly HydratedPreparedRuntimeFormulaRef[]
    | { readonly length: number; readonly at: (index: number) => HydratedPreparedRuntimeFormulaRef },
): HydratedPreparedRuntimeFormulaRef[] {
  const collected: HydratedPreparedRuntimeFormulaRef[] = []
  for (let index = 0; index < refs.length; index += 1) {
    const ref = Array.isArray(refs) ? refs[index]! : refs.at(index)
    collected.push({ ...ref })
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

  it('uses flat formula source refs for runtime-image snapshot fallback formulas when available', () => {
    const workbook = new WorkbookStore('runtime-image-snapshot-formula-sources')
    const sourceCalls: Array<{ cellIndex?: number; col: number; row: number; sheetId: number; source: string }> = []
    const mutationCalls: unknown[] = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-snapshot-formula-sources' },
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

  it('hydrates cached iterative formula values for runtime-image snapshot fallback formulas', () => {
    const workbook = new WorkbookStore('runtime-image-snapshot-cached-formulas')
    const hydratedCalls: HydratedPreparedRuntimeFormulaRef[] = []
    const mutationCalls: unknown[] = []
    const sourceCalls: unknown[] = []
    const compiled = compileFormulaAst('A1+1', parseFormula('A1+1'))
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'runtime-image-snapshot-cached-formulas',
        metadata: {
          calculationSettings: {
            mode: 'automatic',
            compatibilityMode: 'excel-modern',
            iterate: true,
          },
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', row: 0, col: 0, value: 1 },
            { address: 'B1', row: 0, col: 1, formula: 'A1+1', value: 2 },
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
      resolveTemplateForCell: (source, row, col) => ({
        templateId: 9,
        templateKey: 'template:A1+1',
        baseSource: source,
        compiled,
        translated: false,
        rowDelta: row,
        colDelta: col,
      }),
      initializeHydratedPreparedCellFormulasAt: (refs) => {
        hydratedCalls.push(...collectHydratedPreparedRefs(refs))
      },
      initializeFormulaSourcesAt: (refs) => {
        sourceCalls.push(...collectFormulaSourceRefs(refs))
      },
      initializeCellFormulasAt: (refs) => {
        mutationCalls.push(...refs)
      },
    })

    const formulaCellIndex = workbook.getCellIndex('Sheet1', 'B1')
    expect(formulaCellIndex).toBeDefined()
    expect(mutationCalls).toEqual([])
    expect(sourceCalls).toEqual([])
    expect(hydratedCalls).toEqual([
      expect.objectContaining({
        sheetId: 1,
        cellIndex: formulaCellIndex,
        row: 0,
        col: 1,
        source: 'A1+1',
        templateId: 9,
        value: { tag: ValueTag.Number, value: 2 },
      }),
    ])
  })

  it('preserves cached unsupported runtime-image fallback formulas through full recalculation', () => {
    const workbook = new WorkbookStore('runtime-image-snapshot-unsupported-cached-formulas')
    const hydratedCalls: HydratedPreparedRuntimeFormulaRef[] = []
    const mutationCalls: unknown[] = []
    const sourceCalls: unknown[] = []
    const formula = '_FV(A1,"Industry")'
    const compiled = compileFormulaAst(formula, parseFormula(formula))
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-snapshot-unsupported-cached-formulas' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'ADANIPORTS' },
            { address: 'B1', row: 0, col: 1, formula, value: 'Transport Infrastructure' },
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
      resolveTemplateForCell: (source, row, col) => ({
        templateId: 10,
        templateKey: 'template:_FV',
        baseSource: source,
        compiled,
        translated: false,
        rowDelta: row,
        colDelta: col,
      }),
      initializeHydratedPreparedCellFormulasAt: (refs) => {
        hydratedCalls.push(...collectHydratedPreparedRefs(refs))
      },
      initializeFormulaSourcesAt: (refs) => {
        sourceCalls.push(...collectFormulaSourceRefs(refs))
      },
      initializeCellFormulasAt: (refs) => {
        mutationCalls.push(...refs)
      },
    })

    const formulaCellIndex = workbook.getCellIndex('Sheet1', 'B1')
    expect(formulaCellIndex).toBeDefined()
    expect(mutationCalls).toEqual([])
    expect(sourceCalls).toEqual([])
    expect(hydratedCalls).toEqual([
      expect.objectContaining({
        sheetId: 1,
        cellIndex: formulaCellIndex,
        row: 0,
        col: 1,
        source: formula,
        templateId: 10,
        value: { tag: ValueTag.String, value: 'Transport Infrastructure', stringId: expect.any(Number) },
        preserveCachedValueOnFullRecalc: true,
      }),
    ])
  })

  it('shares restored string ids between runtime-image literals and cached formula values', () => {
    const workbook = new WorkbookStore('runtime-image-shared-cached-formula-strings')
    const strings = new StringPool()
    const intern = vi.spyOn(strings, 'intern')
    const hydratedCalls: HydratedPreparedRuntimeFormulaRef[] = []
    const formula = '_FV(A1,"Industry")'
    const repeatedValue = 'Transport Infrastructure'
    const compiled = compileFormulaAst(formula, parseFormula(formula))
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-shared-cached-formula-strings' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', row: 0, col: 0, value: repeatedValue },
            { address: 'B1', row: 0, col: 1, formula, value: repeatedValue },
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
      strings,
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      resolveTemplateForCell: (source, row, col) => ({
        templateId: 11,
        templateKey: 'template:_FV',
        baseSource: source,
        compiled,
        translated: false,
        rowDelta: row,
        colDelta: col,
      }),
      initializeHydratedPreparedCellFormulasAt: (refs) => {
        hydratedCalls.push(...collectHydratedPreparedRefs(refs))
      },
      initializeCellFormulasAt: () => {},
    })

    const literalCellIndex = workbook.getCellIndex('Sheet1', 'A1')!
    expect(intern).toHaveBeenCalledOnce()
    expect(intern).toHaveBeenCalledWith(repeatedValue)
    expect(hydratedCalls).toHaveLength(1)
    expect(hydratedCalls[0].value).toEqual({
      tag: ValueTag.String,
      value: repeatedValue,
      stringId: workbook.cellStore.stringIds[literalCellIndex],
    })
  })

  it('bulk-allocates dense row-major runtime-image sheets during restore', () => {
    const workbook = new WorkbookStore('runtime-image-dense-restore')
    const allocateDenseRowMajorAtReserved = vi.spyOn(workbook.cellStore, 'allocateDenseRowMajorAtReserved')
    const allocateReserved = vi.spyOn(workbook.cellStore, 'allocateReserved')
    const plainCalls: Array<{ mutation: { col: number; formula: string; kind: string; row: number }; sheetId: number }> = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-dense-restore' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'B1', value: 2 },
            { address: 'C1', formula: 'A1+B1' },
            { address: 'A2', value: 3 },
            { address: 'B2', value: null },
            { address: 'C2', formula: 'A2+1' },
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
          { cellIndex: 3, sheetName: 'Sheet1', row: 0, col: 2, source: 'A1+B1' },
          { cellIndex: 6, sheetName: 'Sheet1', row: 1, col: 2, source: 'A2+1' },
        ],
        formulaValues: [],
        sheetCells: [
          {
            sheetName: 'Sheet1',
            cellCount: 6,
            dimensions: { width: 3, height: 2 },
            coords: [
              { row: 0, col: 0 },
              { row: 0, col: 1 },
              { row: 0, col: 2 },
              { row: 1, col: 0 },
              { row: 1, col: 1 },
              { row: 1, col: 2 },
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

    expect(allocateDenseRowMajorAtReserved).toHaveBeenCalledWith(1, 0, 2, 0, 3)
    expect(allocateReserved).not.toHaveBeenCalled()
    const sheet = workbook.getSheet('Sheet1')
    expect(sheet).toBeDefined()
    const a1 = sheet!.grid.get(0, 0)
    const b2 = sheet!.grid.get(1, 1)
    const c1 = sheet!.grid.get(0, 2)
    const c2 = sheet!.grid.get(1, 2)
    const row2Id = sheet!.logicalAxisMap.getId('row', 1)!
    const colCId = sheet!.logicalAxisMap.getId('column', 2)!
    expect(workbook.cellStore.tags[a1]).toBe(ValueTag.Number)
    expect(workbook.cellStore.numbers[a1]).toBe(1)
    expect(workbook.cellStore.tags[b2]).toBe(ValueTag.Empty)
    expect(workbook.cellStore.flags[b2] & CellFlags.AuthoredBlank).not.toBe(0)
    expect(sheet!.grid.getPhysical(1, 2)).toBe(c2)
    expect(sheet!.logical.getCellVisiblePosition(c2)).toEqual({ row: 1, col: 2 })
    expect(sheet!.logical.listResidentCellIndices('row', [row2Id])).toEqual([3, 4, 5])
    expect(sheet!.logical.listResidentCellIndices('column', [colCId])).toEqual([2, 5])
    expect(plainCalls).toEqual([
      {
        sheetId: 1,
        cellIndex: c1,
        mutation: {
          kind: 'setCellFormula',
          row: 0,
          col: 2,
          formula: 'A1+B1',
        },
      },
      {
        sheetId: 1,
        cellIndex: c2,
        mutation: {
          kind: 'setCellFormula',
          row: 1,
          col: 2,
          formula: 'A2+1',
        },
      },
    ])

    allocateDenseRowMajorAtReserved.mockRestore()
    allocateReserved.mockRestore()
  })

  it('trusts compact dense row-major runtime-image coordinates', () => {
    const workbook = new WorkbookStore('runtime-image-dense-order')
    const allocateDenseRowMajorAtReserved = vi.spyOn(workbook.cellStore, 'allocateDenseRowMajorAtReserved')
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-dense-order' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: '__ignored_a__', value: 1 },
            { address: '__ignored_b__', value: 2 },
            { address: '__ignored_c__', value: 3 },
            { address: '__ignored_d__', value: 4 },
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
            cellCount: 4,
            coordinateOrder: 'dense-row-major',
            dimensions: { width: 2, height: 2 },
            coords: [],
          },
        ],
      },
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      initializeCellFormulasAt: () => {},
    })

    const sheet = workbook.getSheet('Sheet1')
    expect(sheet).toBeDefined()
    expect(allocateDenseRowMajorAtReserved).toHaveBeenCalledWith(1, 0, 2, 0, 2)
    expect(workbook.cellStore.numbers[sheet!.grid.get(0, 0)]).toBe(1)
    expect(workbook.cellStore.numbers[sheet!.grid.get(0, 1)]).toBe(2)
    expect(workbook.cellStore.numbers[sheet!.grid.get(1, 0)]).toBe(3)
    expect(workbook.cellStore.numbers[sheet!.grid.get(1, 1)]).toBe(4)

    allocateDenseRowMajorAtReserved.mockRestore()
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
        expect(Array.isArray(refs)).toBe(false)
        hydratedCalls.push(...collectHydratedPreparedRefs(refs))
      },
      initializeCellFormulasAt: () => {},
    })

    expect(hydratedCalls.map((call) => [call.row, call.col, call.value])).toEqual([
      [0, 1, { tag: ValueTag.Number, value: 2 }],
      [1, 1, { tag: ValueTag.Number, value: 3 }],
    ])
  })

  it('reuses restored string ids for repeated runtime-image literal strings', () => {
    const workbook = new WorkbookStore('runtime-image-repeated-strings')
    const strings = new StringPool()
    const intern = vi.spyOn(strings, 'intern')
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'runtime-image-repeated-strings' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 'North' },
            { address: 'B1', value: 'North' },
            { address: 'C1', value: 'South' },
            { address: 'A2', value: 'North' },
            { address: 'B2', value: 'South' },
            { address: 'C2' },
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
            cellCount: 6,
            dimensions: { width: 3, height: 2 },
            coords: [
              { row: 0, col: 0 },
              { row: 0, col: 1 },
              { row: 0, col: 2 },
              { row: 1, col: 0 },
              { row: 1, col: 1 },
              { row: 1, col: 2 },
            ],
          },
        ],
      },
      workbook,
      strings,
      resetWorkbook: () => {},
      hydrateTemplateBank: () => {},
      initializeCellFormulasAt: () => {},
    })

    const sheet = workbook.getSheet('Sheet1')!
    const northA1 = workbook.cellStore.stringIds[sheet.grid.get(0, 0)]
    const northB1 = workbook.cellStore.stringIds[sheet.grid.get(0, 1)]
    const southC1 = workbook.cellStore.stringIds[sheet.grid.get(0, 2)]
    const northA2 = workbook.cellStore.stringIds[sheet.grid.get(1, 0)]
    const southB2 = workbook.cellStore.stringIds[sheet.grid.get(1, 1)]
    const blankC2 = sheet.grid.get(1, 2)

    expect(intern).toHaveBeenCalledTimes(2)
    expect(intern.mock.calls.map(([value]) => value)).toEqual(['North', 'South'])
    expect(northA1).toBe(northB1)
    expect(northA1).toBe(northA2)
    expect(southC1).toBe(southB2)
    expect(workbook.cellStore.tags[blankC2]).toBe(ValueTag.Empty)
    expect(workbook.cellStore.flags[blankC2] & CellFlags.AuthoredBlank).toBe(0)
  })
})

describe('restoreWorkbookFromSnapshot', () => {
  it('restores cached formula error literals as error values for no-recalc imports', () => {
    const workbook = new WorkbookStore('snapshot-cached-formula-error')
    const cachedCalls: CachedRuntimeFormulaRef[] = []
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'snapshot-cached-formula-error',
        metadata: {
          calculationSettings: {
            mode: 'automatic',
            compatibilityMode: 'excel-modern',
            fullCalcOnLoad: false,
          },
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', row: 0, col: 0, value: '#DIV/0!' },
            { address: 'B1', row: 0, col: 1, formula: '1/0', value: '#DIV/0!' },
          ],
        },
      ],
    }

    restoreWorkbookFromSnapshot({
      snapshot,
      workbook,
      strings: new StringPool(),
      resetWorkbook: () => {},
      initializeCellFormulasAt: () => {},
      initializeCachedFormulaSourcesAt: (refs) => {
        cachedCalls.push(...refs)
      },
    })

    expect(cachedCalls).toHaveLength(1)
    expect(cachedCalls[0]?.value).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    const literalCellIndex = workbook.getCellIndex('Sheet1', 'A1')
    expect(literalCellIndex).toBeDefined()
    expect(workbook.cellStore.tags[literalCellIndex!]).toBe(ValueTag.String)
  })

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
            tabColor: { rgb: 'FF0070C0' },
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
    expect(workbook.getSheetTabColor('Sheet1')).toEqual({
      sheetName: 'Sheet1',
      rgb: 'FF0070C0',
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
