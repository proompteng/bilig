import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { createEngineChangeSetEmitterService } from '../engine/services/change-set-emitter-service.js'
import { createEngineCounters } from '../perf/engine-counters.js'

describe('EngineChangeSetEmitterService', () => {
  it('returns empty results for empty and unresolvable change sets', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-empty' })
    engine.createSheet('Sheet1')

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    })

    expect(emitter.captureChangedCells([])).toEqual([])
    expect(emitter.captureChangedCells([999_999])).toEqual([])
  })

  it('emits empty sheet names for stale tiny-path cell indices after sheet deletion', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-stale-tiny' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(a1).toBeDefined()
    engine.workbook.deleteSheet('Sheet1')

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    })

    expect(emitter.captureChangedCells([a1!])).toEqual([
      {
        kind: 'cell',
        cellIndex: a1,
        address: { sheet: 1, row: 0, col: 0 },
        sheetName: '',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 1 },
      },
    ])
  })

  it('captures tiny same-sheet change sets without requiring per-cell sheet-name lookups', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter' })
    const counters = createEngineCounters()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
        counters,
      },
    })

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()

    const changes = emitter.captureChangedCells([a1!, b1!])

    expect(changes).toHaveLength(2)
    expect(changes[0]?.sheetName).toBe('Sheet1')
    expect(changes[1]?.sheetName).toBe('Sheet1')
    expect(changes[0]?.newValue).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(changes[1]?.newValue).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(counters.changedCellPayloadsBuilt).toBe(2)
  })

  it('captures typed patches for tracked runtime consumers', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-patches' })
    const counters = createEngineCounters()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
        counters,
      },
    })

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()

    const patches = emitter.captureChangedPatches([a1!, b1!])

    expect(patches).toEqual([
      {
        kind: 'cell',
        cellIndex: a1,
        address: { sheet: 1, row: 0, col: 0 },
        sheetName: 'Sheet1',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 1 },
      },
      {
        kind: 'cell',
        cellIndex: b1,
        address: { sheet: 1, row: 0, col: 1 },
        sheetName: 'Sheet1',
        a1: 'B1',
        newValue: { tag: ValueTag.Number, value: 2 },
      },
    ])
    expect(counters.changedCellPayloadsBuilt).toBe(2)
  })

  it('captures large tracked patch sets while skipping unresolved entries', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-patches-large' })
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet2', 'B1', 'pear')

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    })

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const a2 = engine.workbook.getCellIndex('Sheet1', 'A2')
    const b1 = engine.workbook.getCellIndex('Sheet2', 'B1')
    expect(a1).toBeDefined()
    expect(a2).toBeDefined()
    expect(b1).toBeDefined()

    const patches = emitter.captureChangedPatches([a1!, 999_999, a2!, b1!])

    expect(patches.map((patch) => `${patch.sheetName}!${patch.a1}`)).toEqual(['Sheet1!A1', 'Sheet1!A2', 'Sheet2!B1'])
  })

  it('preserves stale tiny-path tracked patches with empty sheet names after deletion', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-patches-stale' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(a1).toBeDefined()
    engine.workbook.deleteSheet('Sheet1')

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    })

    expect(emitter.captureChangedPatches([a1!])).toEqual([
      {
        kind: 'cell',
        cellIndex: a1,
        address: { sheet: 1, row: 0, col: 0 },
        sheetName: '',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 1 },
      },
    ])
  })

  it('keeps the tiny tracked-patch fast path when the second cell is unresolved', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-patches-tiny-mixed' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(a1).toBeDefined()

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    })

    expect(emitter.captureChangedPatches([a1!, 999_999])).toEqual([
      {
        kind: 'cell',
        cellIndex: a1,
        address: { sheet: 1, row: 0, col: 0 },
        sheetName: 'Sheet1',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 1 },
      },
    ])
  })

  it('captures larger cross-sheet change sets while skipping unresolved entries', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-large' })
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet2', 'B1', 'pear')

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    })

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const a2 = engine.workbook.getCellIndex('Sheet1', 'A2')
    const b1 = engine.workbook.getCellIndex('Sheet2', 'B1')
    expect(a1).toBeDefined()
    expect(a2).toBeDefined()
    expect(b1).toBeDefined()

    const changes = emitter.captureChangedCells([a1!, 999_999, a2!, b1!])

    expect(changes.map((change) => `${change.sheetName}!${change.a1}`)).toEqual(['Sheet1!A1', 'Sheet1!A2', 'Sheet2!B1'])
    expect(changes[2]?.newValue).toEqual({
      tag: ValueTag.String,
      value: 'pear',
      stringId: expect.any(Number),
    })
  })

  it('preserves stale large-path cell indices with empty sheet names after sheet deletion', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-stale-large' })
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet2', 'B1', 2)

    const stale = engine.workbook.getCellIndex('Sheet2', 'B1')
    const live = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(stale).toBeDefined()
    expect(live).toBeDefined()
    engine.workbook.deleteSheet('Sheet2')

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    })

    const changes = emitter.captureChangedCells([live!, stale!, live!])
    expect(changes.map((change) => `${change.sheetName}!${change.a1}`)).toEqual(['Sheet1!A1', '!B1', 'Sheet1!A1'])
  })

  it('captures typed invalidation patches alongside changed cell patches', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'change-set-emitter-invalidations' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    })

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(a1).toBeDefined()

    const patches = emitter.captureChangedPatches([a1!], {
      invalidation: 'cells',
      invalidatedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }],
      invalidatedRows: [{ sheetName: 'Sheet1', startIndex: 4, endIndex: 6 }],
      invalidatedColumns: [{ sheetName: 'Sheet1', startIndex: 2, endIndex: 3 }],
    })

    expect(patches).toEqual([
      {
        kind: 'cell',
        cellIndex: a1,
        address: { sheet: 1, row: 0, col: 0 },
        sheetName: 'Sheet1',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 1 },
      },
      {
        kind: 'range-invalidation',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
      },
      {
        kind: 'row-invalidation',
        sheetName: 'Sheet1',
        startIndex: 4,
        endIndex: 6,
      },
      {
        kind: 'column-invalidation',
        sheetName: 'Sheet1',
        startIndex: 2,
        endIndex: 3,
      },
    ])
  })
})
