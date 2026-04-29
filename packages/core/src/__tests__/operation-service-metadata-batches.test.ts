import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import { cellMutationRefToEngineOp, type EngineCellMutationRef } from '../cell-mutations-at.js'
import { createBatch } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import type { EngineOperationService } from '../engine/services/operation-service.js'

function isEngineOperationService(value: unknown): value is EngineOperationService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'applyBatch') === 'function'
}

function getOperationService(engine: SpreadsheetEngine): EngineOperationService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (!isEngineOperationService(operations)) {
    throw new TypeError('Expected engine operation service')
  }
  return operations
}

function getReplicaState(engine: SpreadsheetEngine) {
  const replicaState = Reflect.get(engine, 'replicaState')
  if (typeof replicaState !== 'object' || replicaState === null) {
    throw new TypeError('Expected engine replica state')
  }
  return replicaState
}

function applyRemoteOps(engine: SpreadsheetEngine, ops: readonly EngineOp[]): void {
  Effect.runSync(getOperationService(engine).applyBatch(createBatch(getReplicaState(engine), ops), 'remote'))
}

function expectProtectedRemoteOp(engine: SpreadsheetEngine, op: EngineOp): void {
  expect(() => applyRemoteOps(engine, [op])).toThrow(/Workbook protection blocks this change/)
}

const range = {
  sheetName: 'Sheet1',
  startAddress: 'A1',
  endAddress: 'B4',
} as const

describe('EngineOperationService metadata and object batches', () => {
  it('applies workbook metadata, annotation, protection, validation, and object ops through remote batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-service-metadata-batches', replicaId: 'a' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)

    const upsertOps: EngineOp[] = [
      { kind: 'upsertWorkbook', name: 'Metadata Batch Workbook' },
      { kind: 'setWorkbookMetadata', key: 'locale', value: 'en-US' },
      { kind: 'setCalculationSettings', settings: { mode: 'automatic', iterativeCalculation: false } },
      { kind: 'setVolatileContext', context: { now: '2026-04-29T00:00:00.000Z', randomSeed: 7 } },
      { kind: 'upsertCellStyle', style: { id: 'accent', fill: { backgroundColor: '#dbeafe' }, font: { bold: true } } },
      { kind: 'upsertCellNumberFormat', format: { id: 'money', code: '$#,##0.00' } },
      { kind: 'setStyleRange', range, styleId: 'accent' },
      { kind: 'setFormatRange', range, formatId: 'money' },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'C1', format: '0.00%' },
      { kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 1 },
      { kind: 'setFilter', sheetName: 'Sheet1', range },
      { kind: 'setSort', sheetName: 'Sheet1', range, keys: [{ keyAddress: 'A1', direction: 'asc' }] },
      {
        kind: 'setDataValidation',
        validation: {
          range,
          rule: { kind: 'list', values: ['Open', 'Closed'] },
          allowBlank: false,
        },
      },
      {
        kind: 'upsertConditionalFormat',
        format: {
          id: 'cf-open',
          range,
          rule: { kind: 'cellIs', operator: 'greaterThan', values: [5] },
          style: { fill: { backgroundColor: '#bbf7d0' } },
        },
      },
      {
        kind: 'upsertCommentThread',
        thread: {
          threadId: 'thread-1',
          sheetName: 'Sheet1',
          address: 'E2',
          comments: [{ id: 'comment-1', body: 'Review this input.' }],
        },
      },
      {
        kind: 'upsertNote',
        note: {
          sheetName: 'Sheet1',
          address: 'F3',
          text: 'Manual review',
        },
      },
      {
        kind: 'upsertChart',
        chart: {
          id: 'chart-1',
          sheetName: 'Sheet1',
          address: 'H2',
          source: range,
          chartType: 'line',
          rows: 8,
          cols: 5,
          title: 'Trend',
        },
      },
      {
        kind: 'upsertImage',
        image: {
          id: 'image-1',
          sheetName: 'Sheet1',
          address: 'J2',
          sourceUrl: 'https://example.com/image.png',
          rows: 4,
          cols: 4,
          altText: 'Example image',
        },
      },
      {
        kind: 'upsertShape',
        shape: {
          id: 'shape-1',
          sheetName: 'Sheet1',
          address: 'K4',
          shapeType: 'textBox',
          rows: 3,
          cols: 5,
          text: 'Callout',
          fillColor: '#fef3c7',
        },
      },
      {
        kind: 'upsertRangeProtection',
        protection: {
          id: 'protect-range',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'D1',
            endAddress: 'D2',
          },
          hideFormulas: true,
        },
      },
    ]

    Effect.runSync(getOperationService(engine).applyBatch(createBatch(getReplicaState(engine), upsertOps), 'remote'))

    expect(engine.exportSnapshot().workbook.name).toBe('Metadata Batch Workbook')
    expect(engine.workbook.getCellStyle('accent')).toMatchObject({ id: 'accent' })
    expect(engine.workbook.getCellFormat(engine.workbook.getCellIndex('Sheet1', 'C1')!)).toBe('0.00%')
    expect(engine.workbook.getFreezePane('Sheet1')).toMatchObject({ rows: 1, cols: 1 })
    expect(engine.workbook.getFilter('Sheet1', range)).toBeDefined()
    expect(engine.workbook.getSort('Sheet1', range)).toMatchObject({ keys: [{ keyAddress: 'A1', direction: 'asc' }] })
    expect(engine.workbook.getDataValidation('Sheet1', range)).toMatchObject({ allowBlank: false })
    expect(engine.workbook.getConditionalFormat('cf-open')).toBeDefined()
    expect(engine.workbook.getRangeProtection('protect-range')).toBeDefined()
    expect(engine.workbook.getCommentThread('Sheet1', 'E2')).toMatchObject({ threadId: 'thread-1' })
    expect(engine.workbook.getNote('Sheet1', 'F3')).toMatchObject({ text: 'Manual review' })
    expect(engine.workbook.getChart('chart-1')).toMatchObject({ title: 'Trend' })
    expect(engine.workbook.getImage('image-1')).toMatchObject({ altText: 'Example image' })
    expect(engine.workbook.getShape('shape-1')).toMatchObject({ text: 'Callout' })

    Effect.runSync(
      getOperationService(engine).applyBatch(
        createBatch(getReplicaState(engine), [{ kind: 'deleteRangeProtection', id: 'protect-range', sheetName: 'Sheet1' }]),
        'remote',
      ),
    )

    const deleteOps: EngineOp[] = [
      { kind: 'clearFreezePane', sheetName: 'Sheet1' },
      { kind: 'clearFilter', sheetName: 'Sheet1', range },
      { kind: 'clearSort', sheetName: 'Sheet1', range },
      { kind: 'clearDataValidation', sheetName: 'Sheet1', range },
      { kind: 'deleteConditionalFormat', id: 'cf-open', sheetName: 'Sheet1' },
      { kind: 'deleteCommentThread', sheetName: 'Sheet1', address: 'E2' },
      { kind: 'deleteNote', sheetName: 'Sheet1', address: 'F3' },
      { kind: 'deleteChart', id: 'chart-1' },
      { kind: 'deleteImage', id: 'image-1' },
      { kind: 'deleteShape', id: 'shape-1' },
    ]

    Effect.runSync(getOperationService(engine).applyBatch(createBatch(getReplicaState(engine), deleteOps), 'remote'))

    expect(engine.workbook.getFreezePane('Sheet1')).toBeUndefined()
    expect(engine.workbook.getFilter('Sheet1', range)).toBeUndefined()
    expect(engine.workbook.getSort('Sheet1', range)).toBeUndefined()
    expect(engine.workbook.getDataValidation('Sheet1', range)).toBeUndefined()
    expect(engine.workbook.getConditionalFormat('cf-open')).toBeUndefined()
    expect(engine.workbook.getCommentThread('Sheet1', 'E2')).toBeUndefined()
    expect(engine.workbook.getNote('Sheet1', 'F3')).toBeUndefined()
    expect(engine.workbook.getChart('chart-1')).toBeUndefined()
    expect(engine.workbook.getImage('image-1')).toBeUndefined()
    expect(engine.workbook.getShape('shape-1')).toBeUndefined()
  })

  it('blocks protected range object, cell, and derived mutations before applying remote batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-service-protected-range-batches', replicaId: 'a' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B2', 2)

    applyRemoteOps(engine, [
      {
        kind: 'setFilter',
        sheetName: 'Sheet1',
        range,
      },
      {
        kind: 'setSort',
        sheetName: 'Sheet1',
        range,
        keys: [{ keyAddress: 'A1', direction: 'asc' }],
      },
      {
        kind: 'setDataValidation',
        validation: {
          range,
          rule: { kind: 'list', values: ['Open'] },
        },
      },
      {
        kind: 'upsertConditionalFormat',
        format: {
          id: 'cf-protected',
          range,
          rule: { kind: 'blanks' },
          style: {},
        },
      },
      {
        kind: 'upsertCommentThread',
        thread: {
          threadId: 'thread-protected',
          sheetName: 'Sheet1',
          address: 'A1',
          comments: [{ id: 'comment-protected', body: 'review' }],
        },
      },
      { kind: 'upsertNote', note: { sheetName: 'Sheet1', address: 'A2', text: 'note' } },
      {
        kind: 'upsertTable',
        table: {
          name: 'ProtectedTable',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B4',
          columnNames: ['Item', 'Amount'],
          headerRow: true,
          totalsRow: false,
        },
      },
      {
        kind: 'upsertChart',
        chart: {
          id: 'chart-protected',
          sheetName: 'Sheet1',
          address: 'C2',
          source: range,
          chartType: 'bar',
          rows: 4,
          cols: 5,
        },
      },
      {
        kind: 'upsertImage',
        image: {
          id: 'image-protected',
          sheetName: 'Sheet1',
          address: 'C3',
          sourceUrl: 'https://example.com/protected.png',
          rows: 2,
          cols: 2,
        },
      },
      {
        kind: 'upsertShape',
        shape: {
          id: 'shape-protected',
          sheetName: 'Sheet1',
          address: 'C4',
          shapeType: 'textBox',
          rows: 2,
          cols: 3,
        },
      },
      { kind: 'upsertSpillRange', sheetName: 'Sheet1', address: 'B2', rows: 2, cols: 2 },
      {
        kind: 'upsertPivotTable',
        name: 'ProtectedPivot',
        sheetName: 'Sheet1',
        address: 'D2',
        source: range,
        groupBy: ['Item'],
        values: [{ sourceColumn: 'Amount', summarizeBy: 'sum' }],
        rows: 3,
        cols: 2,
      },
    ])
    applyRemoteOps(engine, [
      {
        kind: 'upsertRangeProtection',
        protection: {
          id: 'range-protect',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'D6',
          },
        },
      },
    ])

    const protectedOps: EngineOp[] = [
      { kind: 'setFilter', sheetName: 'Sheet1', range },
      { kind: 'clearFilter', sheetName: 'Sheet1', range },
      { kind: 'setSort', sheetName: 'Sheet1', range, keys: [{ keyAddress: 'A1', direction: 'desc' }] },
      { kind: 'clearSort', sheetName: 'Sheet1', range },
      { kind: 'setStyleRange', range, styleId: 'accent' },
      { kind: 'setFormatRange', range, formatId: 'money' },
      { kind: 'setDataValidation', validation: { range, rule: { kind: 'list', values: ['Closed'] } } },
      { kind: 'clearDataValidation', sheetName: 'Sheet1', range },
      {
        kind: 'upsertConditionalFormat',
        format: {
          id: 'cf-other',
          range,
          rule: { kind: 'blanks' },
          style: {},
        },
      },
      { kind: 'deleteConditionalFormat', id: 'cf-protected', sheetName: 'Sheet1' },
      {
        kind: 'upsertCommentThread',
        thread: {
          threadId: 'thread-other',
          sheetName: 'Sheet1',
          address: 'A1',
          comments: [{ id: 'comment-other', body: 'blocked' }],
        },
      },
      { kind: 'deleteCommentThread', sheetName: 'Sheet1', address: 'A1' },
      { kind: 'upsertNote', note: { sheetName: 'Sheet1', address: 'A2', text: 'blocked' } },
      { kind: 'deleteNote', sheetName: 'Sheet1', address: 'A2' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 99 },
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'A1', formula: 'B2+1' },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'A1', format: '0.00' },
      { kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' },
      {
        kind: 'upsertTable',
        table: {
          name: 'BlockedTable',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B4',
          columnNames: ['Item', 'Amount'],
          headerRow: true,
          totalsRow: false,
        },
      },
      { kind: 'deleteTable', name: 'ProtectedTable' },
      { kind: 'upsertSpillRange', sheetName: 'Sheet1', address: 'B2', rows: 3, cols: 3 },
      { kind: 'deleteSpillRange', sheetName: 'Sheet1', address: 'B2' },
      {
        kind: 'upsertPivotTable',
        name: 'BlockedPivot',
        sheetName: 'Sheet1',
        address: 'D2',
        source: range,
        groupBy: ['Item'],
        values: [{ sourceColumn: 'Amount', summarizeBy: 'sum' }],
        rows: 3,
        cols: 2,
      },
      { kind: 'deletePivotTable', sheetName: 'Sheet1', address: 'D2' },
      {
        kind: 'upsertChart',
        chart: {
          id: 'chart-other',
          sheetName: 'Sheet1',
          address: 'C2',
          source: range,
          chartType: 'line',
          rows: 4,
          cols: 5,
        },
      },
      { kind: 'deleteChart', id: 'chart-protected' },
      {
        kind: 'upsertImage',
        image: {
          id: 'image-other',
          sheetName: 'Sheet1',
          address: 'C3',
          sourceUrl: 'https://example.com/blocked.png',
          rows: 2,
          cols: 2,
        },
      },
      { kind: 'deleteImage', id: 'image-protected' },
      {
        kind: 'upsertShape',
        shape: {
          id: 'shape-other',
          sheetName: 'Sheet1',
          address: 'C4',
          shapeType: 'textBox',
          rows: 2,
          cols: 3,
        },
      },
      { kind: 'deleteShape', id: 'shape-protected' },
    ]

    protectedOps.forEach((op) => {
      expectProtectedRemoteOp(engine, op)
    })
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.workbook.getTable('ProtectedTable')).toBeDefined()
    expect(engine.workbook.getChart('chart-protected')).toBeDefined()
    expect(engine.workbook.getImage('image-protected')).toBeDefined()
    expect(engine.workbook.getShape('shape-protected')).toBeDefined()
  })

  it('blocks protected sheet structural mutations before applying remote batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-service-protected-sheet-batches', replicaId: 'a' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    applyRemoteOps(engine, [{ kind: 'setSheetProtection', protection: { sheetName: 'Sheet1', hideFormulas: true } }])

    const protectedOps: EngineOp[] = [
      { kind: 'renameSheet', oldName: 'Sheet1', newName: 'Renamed' },
      { kind: 'deleteSheet', name: 'Sheet1' },
      { kind: 'insertRows', sheetName: 'Sheet1', start: 1, count: 2 },
      { kind: 'deleteRows', sheetName: 'Sheet1', start: 1, count: 1 },
      { kind: 'moveRows', sheetName: 'Sheet1', start: 1, count: 1, target: 4 },
      { kind: 'insertColumns', sheetName: 'Sheet1', start: 1, count: 2 },
      { kind: 'deleteColumns', sheetName: 'Sheet1', start: 1, count: 1 },
      { kind: 'moveColumns', sheetName: 'Sheet1', start: 1, count: 1, target: 4 },
      { kind: 'updateRowMetadata', sheetName: 'Sheet1', start: 1, count: 1, size: 32, hidden: false },
      { kind: 'updateColumnMetadata', sheetName: 'Sheet1', start: 1, count: 1, size: 120, hidden: false },
      { kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 1 },
      { kind: 'clearFreezePane', sheetName: 'Sheet1' },
    ]

    protectedOps.forEach((op) => {
      expectProtectedRemoteOp(engine, op)
    })
    expect(engine.workbook.getSheet('Sheet1')).toBeDefined()
    expect(engine.workbook.getSheet('Renamed')).toBeUndefined()
  })

  it('skips stale remote ops behind sheet tombstones and entity versions', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-service-stale-remote-batches', replicaId: 'a' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const replicaState = getReplicaState(engine)
    const staleBatch = createBatch(replicaState, [
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'renameSheet', oldName: 'Sheet1', newName: 'OldName' },
      { kind: 'updateRowMetadata', sheetName: 'Sheet1', start: 2, count: 1, size: 44, hidden: false },
      { kind: 'updateColumnMetadata', sheetName: 'Sheet1', start: 2, count: 1, size: 144, hidden: false },
      { kind: 'insertRows', sheetName: 'Sheet1', start: 1, count: 1 },
      { kind: 'deleteRows', sheetName: 'Sheet1', start: 1, count: 1 },
      { kind: 'moveRows', sheetName: 'Sheet1', start: 1, count: 1, target: 3 },
      { kind: 'insertColumns', sheetName: 'Sheet1', start: 1, count: 1 },
      { kind: 'deleteColumns', sheetName: 'Sheet1', start: 1, count: 1 },
      { kind: 'moveColumns', sheetName: 'Sheet1', start: 1, count: 1, target: 3 },
      { kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 1 },
      { kind: 'clearFreezePane', sheetName: 'Sheet1' },
      { kind: 'setFilter', sheetName: 'Sheet1', range },
      { kind: 'clearFilter', sheetName: 'Sheet1', range },
      { kind: 'setSort', sheetName: 'Sheet1', range, keys: [{ keyAddress: 'A1', direction: 'asc' }] },
      { kind: 'clearSort', sheetName: 'Sheet1', range },
      { kind: 'setDataValidation', validation: { range, rule: { kind: 'list', values: ['Old'] } } },
      { kind: 'clearDataValidation', sheetName: 'Sheet1', range },
      { kind: 'deleteConditionalFormat', id: 'missing-cf', sheetName: 'Sheet1' },
      { kind: 'deleteRangeProtection', id: 'missing-protection', sheetName: 'Sheet1' },
      { kind: 'deleteCommentThread', sheetName: 'Sheet1', address: 'A1' },
      { kind: 'deleteNote', sheetName: 'Sheet1', address: 'A1' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 },
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'A2', formula: 'A1+1' },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'A1', format: '0.00' },
      { kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' },
      { kind: 'upsertSpillRange', sheetName: 'Sheet1', address: 'C1', rows: 2, cols: 2 },
      { kind: 'deleteSpillRange', sheetName: 'Sheet1', address: 'C1' },
      { kind: 'deletePivotTable', sheetName: 'Sheet1', address: 'E1' },
    ])
    const staleMetadataBatch = createBatch(replicaState, [{ kind: 'setWorkbookMetadata', key: 'locale', value: 'stale' }])
    const freshMetadataBatch = createBatch(replicaState, [{ kind: 'setWorkbookMetadata', key: 'locale', value: 'fresh' }])
    const deleteSheetBatch = createBatch(replicaState, [{ kind: 'deleteSheet', name: 'Sheet1' }])

    Effect.runSync(getOperationService(engine).applyBatch(freshMetadataBatch, 'remote'))
    Effect.runSync(getOperationService(engine).applyBatch(staleMetadataBatch, 'remote'))
    expect(engine.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'fresh' }])

    Effect.runSync(getOperationService(engine).applyBatch(deleteSheetBatch, 'remote'))
    Effect.runSync(getOperationService(engine).applyBatch(staleBatch, 'remote'))
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()
    expect(engine.workbook.getSheet('OldName')).toBeUndefined()
  })

  it('applies dense row-pair scalar batches with multi-formula rows without dirty traversal', async () => {
    const rowCount = 48
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-direct-scalar-row-pair-multi-formula-batch',
      replicaId: 'a',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row + 1)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `A${row}*B${row}`)
      engine.setCellFormula('Sheet1', `E${row}`, `A${row}-B${row}`)
    }

    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 0, value: row * 7 },
      })
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `B${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 1, value: row * 11 },
      })
    }
    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 0))

    const lastRow = rowCount - 1
    expect(engine.getCellValue('Sheet1', `C${rowCount}`)).toEqual({ tag: ValueTag.Number, value: lastRow * 18 })
    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({ tag: ValueTag.Number, value: lastRow * 7 * (lastRow * 11) })
    expect(engine.getCellValue('Sheet1', `E${rowCount}`)).toEqual({ tag: ValueTag.Number, value: lastRow * -4 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount * 3)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('coalesces non-contiguous numeric scalar batches and emits authoritative changed slices', async () => {
    const rowCount = 40
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-direct-scalar-coalesced-non-contiguous-batch',
      replicaId: 'a',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `C${row}`, row * 2)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}+10`)
      engine.setCellFormula('Sheet1', `D${row}`, `C${row}*3`)
    }

    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 0, value: 100 + row },
      })
    }
    for (let row = 0; row < rowCount; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `C${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 2, value: 200 + row },
      })
    }

    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )
    const general = vi.fn()
    const tracked = vi.fn()
    const cellListener = vi.fn()
    const unsubscribeGeneral = engine.subscribe(general)
    const unsubscribeTracked = engine.events.subscribeTracked(tracked)
    const unsubscribeCell = engine.subscribeCell('Sheet1', `D${rowCount}`, cellListener)

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 0))

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: 100 + rowCount - 1 + 10 })
    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({ tag: ValueTag.Number, value: (200 + rowCount - 1) * 3 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount * 2)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    expect(general).toHaveBeenCalledTimes(1)
    expect(tracked).toHaveBeenCalledTimes(1)
    expect(cellListener).toHaveBeenCalledTimes(1)
    const event = tracked.mock.calls[0]?.[0]
    expect(event?.explicitChangedCount).toBe(rowCount * 2)
    expect(event?.changedCellIndices).toHaveLength(rowCount * 4)

    unsubscribeCell()
    unsubscribeTracked()
    unsubscribeGeneral()
  })

  it('creates missing numeric cells through the single-literal kernel-sync fast path', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-single-literal-missing-numeric-fast-path',
      replicaId: 'a',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellValue', row: 9, col: 4, value: 42 },
      },
    ]

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, null, 'local', 0))

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'E10')
    expect(cellIndex).toBeDefined()
    expect(engine.getCellValue('Sheet1', 'E10')).toEqual({ tag: ValueTag.Number, value: 42 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBe(1)
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([cellIndex!]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('coalesces division-by-zero scalar batches into rendered formula errors without dirty traversal', async () => {
    const rowCount = 40
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-direct-scalar-coalesced-current-error-batch',
      replicaId: 'a',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row * 10)
      engine.setCellValue('Sheet1', `C${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}/C${row}`)
    }

    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `C${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 2, value: 0 },
      })
    }
    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 0))

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })
})
