import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it } from 'vitest'
import { defineModel, planWorkbookAction, toPlanData } from '@bilig/workbook'
import {
  applyWorkbookEvent,
  isAuthoritativeWorkbookEventBatch,
  isAuthoritativeWorkbookEventBatchAfterRevision,
  isAuthoritativeWorkbookEventRecord,
  isWorkbookChangeUndoBundle,
  isWorkbookEventPayload,
} from '../workbook-events.js'

function buildPlanData() {
  const model = defineModel({
    name: 'event-plan',
    find(workbook) {
      return { output: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }) }
    },
    actions: {
      write({ refs, workbook }) {
        workbook.writeValue(refs.output, 1)
      },
    },
  })
  const planned = planWorkbookAction(model, 'write')
  if (planned.status !== 'planned') {
    throw new Error(planned.errors.map((error) => error.message).join('\n'))
  }
  return toPlanData(planned.plan)
}

function workbookPlanDataAppliedOp() {
  return { kind: 'setCellValue' as const, sheetName: 'Sheet1', address: 'A1', value: 1 }
}

function buildAuthoritativeCellEvent(revision: number) {
  return {
    revision,
    clientMutationId: null,
    payload: {
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'A1',
      value: revision,
    },
  }
}

describe('workbook event guards', () => {
  it('accepts applyBatch payloads with a valid engine op batch', () => {
    expect(
      isWorkbookEventPayload({
        kind: 'applyBatch',
        batch: {
          id: 'batch-1',
          replicaId: 'replica-1',
          clock: { counter: 2 },
          ops: [{ kind: 'upsertWorkbook', name: 'Book' }],
        },
      }),
    ).toBe(true)
  })

  it('accepts transported workbook plan data payloads with concrete applied ops', () => {
    const plan = buildPlanData()
    expect(
      isWorkbookEventPayload({
        kind: 'applyWorkbookPlanData',
        plan,
        appliedOps: [workbookPlanDataAppliedOp()],
        result: {
          status: 'done',
          apply: {
            matched: true,
            baseRevision: 0,
            revision: 1,
            previewOps: [workbookPlanDataAppliedOp()],
            appliedOps: [workbookPlanDataAppliedOp()],
            commandReceipts: [
              {
                commandIndex: 0,
                commandKind: 'writeValue',
                commandDigest: 'digest',
                previewOps: [workbookPlanDataAppliedOp()],
                appliedOps: [workbookPlanDataAppliedOp()],
              },
            ],
          },
          changed: [],
          checks: [],
        },
      }),
    ).toBe(true)
    expect(
      isWorkbookEventPayload({
        kind: 'applyWorkbookPlanData',
        plan: buildPlanData(),
        appliedOps: [{ kind: 'not-an-op' }],
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'applyWorkbookPlanData',
        plan,
        appliedOps: [workbookPlanDataAppliedOp()],
        result: {
          status: 'done',
          changed: [],
          checks: [],
        },
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'applyWorkbookPlanData',
        plan,
        appliedOps: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }],
        result: {
          status: 'failed',
          errors: [() => undefined],
          changed: [],
          checks: [],
        },
      }),
    ).toBe(false)
  })

  it('rejects renderCommit payloads with malformed commit ops', () => {
    expect(
      isWorkbookEventPayload({
        kind: 'renderCommit',
        ops: [{ kind: 'renameSheet', oldName: 'Sheet1' }],
      }),
    ).toBe(false)
  })

  it('rejects setCellValue payloads with non-literal values', () => {
    expect(
      isWorkbookEventPayload({
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'A1',
        value: 'ready',
      }),
    ).toBe(true)
    expect(
      isWorkbookEventPayload({
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'A1',
        value: Number.NaN,
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: 'not-literal' },
      }),
    ).toBe(false)
  })

  it('accepts structural metadata payloads', () => {
    expect(
      isWorkbookEventPayload({
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        startRow: 1,
        count: 2,
        height: 32,
        hidden: false,
      }),
    ).toBe(true)

    expect(
      isWorkbookEventPayload({
        kind: 'setFreezePane',
        sheetName: 'Sheet1',
        rows: 1,
        cols: 2,
      }),
    ).toBe(true)

    expect(
      isWorkbookEventPayload({
        kind: 'mergeCells',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      }),
    ).toBe(true)

    expect(
      isWorkbookEventPayload({
        kind: 'unmergeCells',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      }),
    ).toBe(true)

    expect(
      isWorkbookEventPayload({
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 1,
        count: 2,
      }),
    ).toBe(true)

    expect(
      isWorkbookEventPayload({
        kind: 'deleteColumns',
        sheetName: 'Sheet1',
        start: 3,
        count: 1,
      }),
    ).toBe(true)

    expect(
      isWorkbookEventPayload({
        kind: 'redoChange',
        targetRevision: 12,
        targetSummary: 'Updated Sheet1!A1',
        sheetName: 'Sheet1',
        address: 'A1',
        appliedBundle: {
          kind: 'engineOps',
          ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }],
        },
      }),
    ).toBe(true)

    expect(
      isWorkbookEventPayload({
        kind: 'revertChange',
        targetRevision: 13,
        targetSummary: 'Inserted rows 3:4 on Sheet1',
        sheetName: 'Sheet1',
        address: 'A3',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A3',
          endAddress: 'A4',
          scope: 'rows',
        },
        appliedBundle: {
          kind: 'engineOps',
          ops: [{ kind: 'deleteRows', sheetName: 'Sheet1', start: 2, count: 2 }],
        },
      }),
    ).toBe(true)
  })

  it('uses mutator schemas when validating formatting event payloads for replay', () => {
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }

    expect(
      isWorkbookEventPayload({
        kind: 'setRangeStyle',
        range,
        patch: { font: { bold: true }, alignment: { horizontal: 'center' } },
      }),
    ).toBe(true)
    expect(
      isWorkbookEventPayload({
        kind: 'clearRangeStyle',
        range,
        fields: ['fontBold', 'alignmentHorizontal'],
      }),
    ).toBe(true)
    expect(
      isWorkbookEventPayload({
        kind: 'setRangeNumberFormat',
        range,
        format: { kind: 'currency', currency: 'USD', decimals: 2 },
      }),
    ).toBe(true)

    expect(
      isWorkbookEventPayload({
        kind: 'setRangeStyle',
        range,
        patch: null,
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'setRangeStyle',
        range,
        patch: { alignment: { horizontal: 'middle' } },
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'setRangeStyle',
        range,
        patch: { textColor: '#111111' },
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'clearRangeStyle',
        range,
        fields: ['fontBold', 'bogusField'],
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'setRangeNumberFormat',
        range,
        format: { kind: 'currency', negativeStyle: 'red' },
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'setRangeNumberFormat',
        range,
        format: { kind: 'currency', locale: 'en-US' },
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'setRangeNumberFormat',
        range: { sheetName: '', startAddress: 'A1', endAddress: 'B2' },
        format: '$#,##0.00',
      }),
    ).toBe(false)
  })

  it('rejects unsafe structural metadata payload numbers', () => {
    const unsafe = Number.MAX_SAFE_INTEGER + 1

    expect(
      isWorkbookEventPayload({
        kind: 'setFreezePane',
        sheetName: 'Sheet1',
        rows: unsafe,
        cols: 0,
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 0,
        count: unsafe,
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        startRow: 0,
        count: 1,
        height: unsafe,
        hidden: null,
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'updateColumnWidth',
        sheetName: 'Sheet1',
        columnIndex: unsafe,
        width: 44,
      }),
    ).toBe(false)
  })

  it('rejects history payloads with malformed structural range scope', () => {
    expect(
      isWorkbookEventPayload({
        kind: 'revertChange',
        targetRevision: 13,
        targetSummary: 'Inserted rows 3:4 on Sheet1',
        sheetName: 'Sheet1',
        address: 'A3',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A3',
          endAddress: 'A4',
          scope: 'row-band',
        },
        appliedBundle: {
          kind: 'engineOps',
          ops: [{ kind: 'deleteRows', sheetName: 'Sheet1', start: 2, count: 2 }],
        },
      }),
    ).toBe(false)
  })

  it('rejects unsafe event sequence numbers', () => {
    const unsafe = Number.MAX_SAFE_INTEGER + 1

    expect(
      isWorkbookEventPayload({
        kind: 'redoChange',
        targetRevision: unsafe,
        targetSummary: 'Updated Sheet1!A1',
        appliedBundle: {
          kind: 'engineOps',
          ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }],
        },
      }),
    ).toBe(false)
    expect(
      isWorkbookEventPayload({
        kind: 'applyBatch',
        batch: {
          id: 'batch-1',
          replicaId: 'replica-1',
          clock: { counter: unsafe },
          ops: [{ kind: 'upsertWorkbook', name: 'Book' }],
        },
      }),
    ).toBe(false)
    expect(
      isAuthoritativeWorkbookEventRecord({
        revision: unsafe,
        clientMutationId: null,
        payload: {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 1,
        },
      }),
    ).toBe(false)
    expect(
      isAuthoritativeWorkbookEventBatch({
        afterRevision: 0,
        headRevision: unsafe,
        calculatedRevision: unsafe,
        events: [],
      }),
    ).toBe(false)
  })

  it('rejects authoritative event batches with revision gaps or impossible cursors', () => {
    expect(
      isAuthoritativeWorkbookEventBatch({
        afterRevision: 2,
        headRevision: 4,
        calculatedRevision: 4,
        events: [buildAuthoritativeCellEvent(4)],
      }),
    ).toBe(false)
    expect(
      isAuthoritativeWorkbookEventBatch({
        afterRevision: 4,
        headRevision: 3,
        calculatedRevision: 3,
        events: [],
      }),
    ).toBe(false)
    expect(
      isAuthoritativeWorkbookEventBatch({
        afterRevision: 2,
        headRevision: 3,
        calculatedRevision: 4,
        events: [buildAuthoritativeCellEvent(3)],
      }),
    ).toBe(false)
    expect(
      isAuthoritativeWorkbookEventBatch({
        afterRevision: 2,
        headRevision: 4,
        calculatedRevision: 4,
        events: [buildAuthoritativeCellEvent(3), buildAuthoritativeCellEvent(4)],
      }),
    ).toBe(true)
  })

  it('rejects authoritative event batches that do not match the requested after revision', () => {
    const batch = {
      afterRevision: 2,
      headRevision: 4,
      calculatedRevision: 4,
      events: [buildAuthoritativeCellEvent(3), buildAuthoritativeCellEvent(4)],
    }

    expect(isAuthoritativeWorkbookEventBatchAfterRevision(batch, 2)).toBe(true)
    expect(isAuthoritativeWorkbookEventBatchAfterRevision(batch, 1)).toBe(false)
    expect(isAuthoritativeWorkbookEventBatchAfterRevision(batch, Number.NaN)).toBe(false)
    expect(isAuthoritativeWorkbookEventBatchAfterRevision(batch, Number.MAX_SAFE_INTEGER + 1)).toBe(false)
  })

  it('rejects engine undo bundles with malformed engine ops', () => {
    expect(
      isWorkbookChangeUndoBundle({
        kind: 'engineOps',
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1' }],
      }),
    ).toBe(false)
  })
})

it('replays imported snapshots with formulas and names through the authoritative event stream', async () => {
  const snapshot = {
    version: 1 as const,
    workbook: { name: 'Budget' },
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        cells: [
          { address: 'A1', value: 3 },
          { address: 'B1', formula: 'A1*4' },
        ],
      },
    ],
  }
  const payload = { kind: 'importWorkbookSnapshot' as const, snapshot }
  expect(isWorkbookEventPayload(payload)).toBe(true)
  expect(isWorkbookEventPayload({ ...payload, snapshot: { version: 2 } })).toBe(false)
  const engine = new SpreadsheetEngine()
  await engine.ready()
  applyWorkbookEvent(engine, payload)
  expect(engine.exportSnapshot().workbook.name).toBe('Budget')
  expect(engine.exportSnapshot().sheets[0]?.cells.find((cell) => cell.address === 'B1')?.formula).toBe('A1*4')
})
