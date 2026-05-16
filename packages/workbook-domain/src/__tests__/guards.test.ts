import { describe, expect, it } from 'vitest'
import { isEngineOp, isEngineOpBatch } from '../index.js'

describe('workbook domain guards', () => {
  it('accepts engine op batches with valid workbook ops', () => {
    expect(
      isEngineOpBatch({
        id: 'batch-1',
        replicaId: 'replica-1',
        clock: { counter: 4 },
        ops: [
          { kind: 'upsertWorkbook', name: 'Book' },
          {
            kind: 'setDataValidation',
            validation: {
              range: {
                sheetName: 'Sheet1',
                startAddress: 'D2',
                endAddress: 'D10',
              },
              rule: {
                kind: 'list',
                values: ['Draft', 'Final'],
              },
              allowBlank: false,
            },
          },
          {
            kind: 'upsertCommentThread',
            thread: {
              threadId: 'thread-1',
              sheetName: 'Sheet1',
              address: 'E2',
              comments: [{ id: 'comment-1', body: 'Check this total.' }],
            },
          },
          {
            kind: 'upsertNote',
            note: {
              sheetName: 'Sheet1',
              address: 'F3',
              text: 'Manual override',
            },
          },
          {
            kind: 'upsertConditionalFormat',
            format: {
              id: 'cf-1',
              range: {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'A5',
              },
              rule: {
                kind: 'cellIs',
                operator: 'greaterThan',
                values: [10],
              },
              style: {
                fill: { backgroundColor: '#ff0000' },
              },
            },
          },
          {
            kind: 'setSheetProtection',
            protection: {
              sheetName: 'Sheet1',
              hideFormulas: true,
            },
          },
          {
            kind: 'upsertRangeProtection',
            protection: {
              id: 'protect-a1',
              range: {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'B2',
              },
              hideFormulas: true,
            },
          },
          {
            kind: 'upsertPivotTable',
            name: 'Pivot1',
            sheetName: 'Sheet1',
            address: 'F1',
            source: {
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'C10',
            },
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Total Sales' }],
            rows: 10,
            cols: 3,
          },
          {
            kind: 'upsertChart',
            chart: {
              id: 'chart-1',
              sheetName: 'Sheet1',
              address: 'J2',
              source: {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'C10',
              },
              chartType: 'line',
              rows: 12,
              cols: 8,
              title: 'Sales trend',
            },
          },
          {
            kind: 'upsertImage',
            image: {
              id: 'image-1',
              sheetName: 'Sheet1',
              address: 'L2',
              sourceUrl: 'https://example.com/chart.png',
              rows: 8,
              cols: 5,
              altText: 'Revenue image',
            },
          },
          {
            kind: 'upsertShape',
            shape: {
              id: 'shape-1',
              sheetName: 'Sheet1',
              address: 'M3',
              shapeType: 'textBox',
              rows: 4,
              cols: 6,
              text: 'Review',
              fillColor: '#ffeeaa',
            },
          },
        ],
      }),
    ).toBe(true)
  })

  it('rejects engine ops with malformed nested payloads', () => {
    expect(
      isEngineOp({
        kind: 'upsertCellStyle',
        style: {
          id: 'style-1',
          fill: {
            backgroundColor: 42,
          },
        },
      }),
    ).toBe(false)

    expect(
      isEngineOpBatch({
        id: 'batch-1',
        replicaId: 'replica-1',
        clock: { counter: 4 },
        ops: [
          { kind: 'setSort', sheetName: 'Sheet1', range: { sheetName: 'Sheet1' }, keys: [] },
          {
            kind: 'setDataValidation',
            validation: {
              range: {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'A5',
              },
              rule: {
                kind: 'list',
                values: [undefined],
              },
            },
          },
          {
            kind: 'upsertCommentThread',
            thread: {
              threadId: 'thread-1',
              sheetName: 'Sheet1',
              address: 'E2',
              comments: [{ id: 'comment-1' }],
            },
          },
          {
            kind: 'upsertConditionalFormat',
            format: {
              id: 'cf-1',
              range: {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'A5',
              },
              rule: {
                kind: 'cellIs',
                operator: 'greaterThan',
                values: [],
              },
              style: 'bad',
            },
          },
          {
            kind: 'upsertRangeProtection',
            protection: {
              id: 'protect-a1',
              range: {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'B2',
              },
              hideFormulas: 'yes',
            },
          },
          {
            kind: 'upsertChart',
            chart: {
              id: 'chart-1',
              sheetName: 'Sheet1',
              address: 'J2',
              source: {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'C10',
              },
              chartType: 'donut',
              rows: 12,
              cols: 8,
            },
          },
          {
            kind: 'upsertImage',
            image: {
              id: 'image-1',
              sheetName: 'Sheet1',
              address: 'L2',
              sourceUrl: 7,
              rows: 8,
              cols: 5,
            },
          },
        ],
      }),
    ).toBe(false)
  })

  it('rejects unsafe engine batch clock counters', () => {
    const validBatch = {
      id: 'batch-1',
      replicaId: 'replica-1',
      clock: { counter: 4 },
      ops: [{ kind: 'upsertWorkbook', name: 'Book' }],
    }

    expect(isEngineOpBatch({ ...validBatch, clock: { counter: 1.5 } })).toBe(false)
    expect(isEngineOpBatch({ ...validBatch, clock: { counter: -1 } })).toBe(false)
    expect(isEngineOpBatch({ ...validBatch, clock: { counter: Number.MAX_SAFE_INTEGER + 1 } })).toBe(false)
  })

  it('rejects unsafe sheet identity metadata', () => {
    const unsafe = Number.MAX_SAFE_INTEGER + 1

    expect(isEngineOp({ kind: 'upsertSheet', name: 'Sheet2', order: -1 })).toBe(false)
    expect(isEngineOp({ kind: 'upsertSheet', name: 'Sheet2', order: 1.5 })).toBe(false)
    expect(isEngineOp({ kind: 'upsertSheet', name: 'Sheet2', order: unsafe })).toBe(false)
    expect(isEngineOp({ kind: 'upsertSheet', name: 'Sheet2', order: 1, id: 0 })).toBe(false)
    expect(isEngineOp({ kind: 'upsertSheet', name: 'Sheet2', order: 1, id: unsafe })).toBe(false)
  })

  it('rejects unsafe persisted metadata sequence fields', () => {
    const unsafe = Number.MAX_SAFE_INTEGER + 1
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }

    expect(
      isEngineOp({
        kind: 'upsertConditionalFormat',
        format: {
          id: 'cf-1',
          range,
          rule: { kind: 'textContains', text: 'late' },
          style: {},
          priority: unsafe,
        },
      }),
    ).toBe(false)
    expect(
      isEngineOp({
        kind: 'upsertCommentThread',
        thread: {
          threadId: 'thread-1',
          sheetName: 'Sheet1',
          address: 'A1',
          comments: [{ id: 'comment-1', body: 'Check', createdAtUnixMs: 1.5 }],
        },
      }),
    ).toBe(false)
    expect(
      isEngineOp({
        kind: 'upsertCommentThread',
        thread: {
          threadId: 'thread-1',
          sheetName: 'Sheet1',
          address: 'A1',
          comments: [{ id: 'comment-1', body: 'Check', createdAtUnixMs: 1 }],
          resolvedAtUnixMs: -1,
        },
      }),
    ).toBe(false)
  })

  it('rejects unsafe structural workbook op coordinates', () => {
    const unsafe = Number.MAX_SAFE_INTEGER + 1

    expect(isEngineOp({ kind: 'insertRows', sheetName: 'Sheet1', start: 1.5, count: 1 })).toBe(false)
    expect(isEngineOp({ kind: 'insertColumns', sheetName: 'Sheet1', start: 1, count: 0 })).toBe(false)
    expect(isEngineOp({ kind: 'deleteColumns', sheetName: 'Sheet1', start: 0, count: unsafe })).toBe(false)
    expect(isEngineOp({ kind: 'moveRows', sheetName: 'Sheet1', start: 0, count: 1, target: unsafe })).toBe(false)
    expect(isEngineOp({ kind: 'updateRowMetadata', sheetName: 'Sheet1', start: 0, count: 1, size: 0, hidden: null })).toBe(false)
    expect(isEngineOp({ kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1.5, cols: 0 })).toBe(false)
    expect(
      isEngineOp({
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 0,
        count: 1,
        entries: [{ id: 'row-1', index: unsafe, size: 24, hidden: false }],
      }),
    ).toBe(false)
  })

  it('rejects unsafe workbook object footprint dimensions', () => {
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }
    const unsafe = Number.MAX_SAFE_INTEGER + 1

    expect(isEngineOp({ kind: 'upsertSpillRange', sheetName: 'Sheet1', address: 'C1', rows: 0, cols: 1 })).toBe(false)
    expect(
      isEngineOp({
        kind: 'upsertPivotTable',
        name: 'Pivot1',
        sheetName: 'Sheet1',
        address: 'F1',
        source: range,
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
        rows: unsafe,
        cols: 2,
      }),
    ).toBe(false)
    expect(
      isEngineOp({
        kind: 'upsertChart',
        chart: { id: 'chart-1', sheetName: 'Sheet1', address: 'J2', source: range, chartType: 'line', rows: 1.5, cols: 8 },
      }),
    ).toBe(false)
    expect(
      isEngineOp({
        kind: 'upsertImage',
        image: { id: 'image-1', sheetName: 'Sheet1', address: 'L2', sourceUrl: 'https://example.com/i.png', rows: 8, cols: unsafe },
      }),
    ).toBe(false)
    expect(
      isEngineOp({
        kind: 'upsertShape',
        shape: { id: 'shape-1', sheetName: 'Sheet1', address: 'M3', shapeType: 'textBox', rows: -1, cols: 6 },
      }),
    ).toBe(false)
  })
})
