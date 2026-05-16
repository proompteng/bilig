import { describe, expect, it } from 'vitest'
import {
  buildZeroWorkbookMutation,
  isPendingWorkbookMutation,
  isPendingWorkbookMutationInput,
  isPendingWorkbookMutationList,
  isWorkbookMutationMethod,
  type PendingWorkbookMutation,
} from '../workbook-sync.js'

function createPendingMutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: 'pending-1',
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 17],
    localSeq: 1,
    baseRevision: 0,
    enqueuedAtUnixMs: 100,
    submittedAtUnixMs: null,
    lastAttemptedAtUnixMs: null,
    ackedAtUnixMs: null,
    rebasedAtUnixMs: null,
    failedAtUnixMs: null,
    attemptCount: 0,
    failureMessage: null,
    status: 'local',
    ...overrides,
  }
}

describe('buildZeroWorkbookMutation', () => {
  it('builds structural metadata mutations', () => {
    expect(
      buildZeroWorkbookMutation('doc-1', {
        method: 'updateRowMetadata',
        args: ['Sheet1', 2, 3, 48, true],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 3,
        documentId: 'doc-1',
        height: 48,
        hidden: true,
        sheetName: 'Sheet1',
        startRow: 2,
      },
      '~': 'MutateRequest',
    })

    expect(
      buildZeroWorkbookMutation('doc-1', {
        method: 'updateColumnMetadata',
        args: ['Sheet1', 3, 1, 144, null],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 1,
        documentId: 'doc-1',
        hidden: null,
        sheetName: 'Sheet1',
        startCol: 3,
        width: 144,
      },
      '~': 'MutateRequest',
    })

    expect(
      buildZeroWorkbookMutation('doc-1', {
        method: 'setFreezePane',
        args: ['Sheet1', 1, 2],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        cols: 2,
        documentId: 'doc-1',
        rows: 1,
        sheetName: 'Sheet1',
      },
      '~': 'MutateRequest',
    })

    expect(
      buildZeroWorkbookMutation('doc-1', {
        method: 'insertRows',
        args: ['Sheet1', 4, 2],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 2,
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        start: 4,
      },
      '~': 'MutateRequest',
    })

    expect(
      buildZeroWorkbookMutation('doc-1', {
        method: 'deleteColumns',
        args: ['Sheet1', 1, 3],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 3,
        documentId: 'doc-1',
        sheetName: 'Sheet1',
        start: 1,
      },
      '~': 'MutateRequest',
    })
  })

  it('normalizes legacy updateColumnWidth journal entries onto column metadata mutations', () => {
    const legacyMutation = {
      method: 'updateColumnWidth',
      args: ['Sheet1', 5, 168],
    }

    expect(isPendingWorkbookMutationInput(legacyMutation)).toBe(true)
    if (!isPendingWorkbookMutationInput(legacyMutation)) {
      throw new Error('expected legacy mutation to remain readable')
    }

    expect(buildZeroWorkbookMutation('doc-1', legacyMutation)).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 1,
        documentId: 'doc-1',
        hidden: null,
        sheetName: 'Sheet1',
        startCol: 5,
        width: 168,
      },
      '~': 'MutateRequest',
    })
  })

  it('does not advertise updateColumnWidth as an active workbook mutation method', () => {
    expect(isWorkbookMutationMethod('updateColumnWidth')).toBe(false)
    expect(isWorkbookMutationMethod('updateColumnMetadata')).toBe(true)
    expect(isWorkbookMutationMethod('insertRows')).toBe(true)
    expect(isWorkbookMutationMethod('deleteColumns')).toBe(true)
  })

  it('accepts renderCommit mutations with valid commit ops', () => {
    expect(() =>
      buildZeroWorkbookMutation('doc-1', {
        method: 'renderCommit',
        args: [[{ kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 7 }]],
      }),
    ).not.toThrow()
  })

  it('rejects renderCommit mutations with malformed commit ops', () => {
    expect(() =>
      buildZeroWorkbookMutation('doc-1', {
        method: 'renderCommit',
        args: [[{ kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1' }]],
      }),
    ).toThrow('Invalid renderCommit args')
  })

  it('rejects malformed structural mutation numeric args before building Zero payloads', () => {
    for (const mutation of [
      { method: 'insertRows' as const, args: ['Sheet1', Number.NaN, 1], error: 'Invalid insertRows args' },
      { method: 'deleteColumns' as const, args: ['Sheet1', 1.5, 1], error: 'Invalid deleteColumns args' },
      { method: 'updateRowMetadata' as const, args: ['Sheet1', 1, 1, Number.NaN, null], error: 'Invalid updateRowMetadata args' },
      {
        method: 'updateColumnMetadata' as const,
        args: ['Sheet1', 1, 1, Number.POSITIVE_INFINITY, null],
        error: 'Invalid updateColumnMetadata args',
      },
      { method: 'setFreezePane' as const, args: ['Sheet1', 1.5, 0], error: 'Invalid setFreezePane args' },
      { method: 'updateColumnWidth' as const, args: ['Sheet1', -1, 120], error: 'Invalid updateColumnWidth args' },
    ]) {
      expect(() => buildZeroWorkbookMutation('doc-1', mutation)).toThrow(mutation.error)
    }
  })

  it('rejects malformed persisted pending mutation structural args', () => {
    for (const mutation of [
      createPendingMutation({ method: 'insertRows', args: ['Sheet1', Number.NaN, 1] }),
      createPendingMutation({ method: 'deleteColumns', args: ['Sheet1', 1.5, 1] }),
      createPendingMutation({ method: 'updateRowMetadata', args: ['Sheet1', 1, 1, Number.NaN, null] }),
      createPendingMutation({ method: 'updateColumnMetadata', args: ['Sheet1', 1, 1, Number.POSITIVE_INFINITY, null] }),
      createPendingMutation({ method: 'setFreezePane', args: ['Sheet1', 1.5, 0] }),
      createPendingMutation({ method: 'updateColumnWidth', args: ['Sheet1', -1, 120] }),
    ]) {
      expect(isPendingWorkbookMutationInput(mutation)).toBe(false)
      expect(isPendingWorkbookMutation(mutation)).toBe(false)
      expect(isPendingWorkbookMutationList([createPendingMutation(), mutation])).toBe(false)
    }
  })

  it('rejects persisted pending mutations with extra ignored args', () => {
    for (const mutation of [
      createPendingMutation({ method: 'setCellValue', args: ['Sheet1', 'A1', 17, { ignored: true }] }),
      createPendingMutation({ method: 'clearCell', args: ['Sheet1', 'A1', 'ignored'] }),
      createPendingMutation({ method: 'renderCommit', args: [[{ kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 7 }], null] }),
      createPendingMutation({ method: 'insertRows', args: ['Sheet1', 1, 1, 'ignored'] }),
      createPendingMutation({ method: 'updateRowMetadata', args: ['Sheet1', 1, 1, 24, null, 'ignored'] }),
      createPendingMutation({ method: 'setFreezePane', args: ['Sheet1', 1, 1, 'ignored'] }),
      createPendingMutation({
        method: 'clearRangeNumberFormat',
        args: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, null],
      }),
    ]) {
      expect(isPendingWorkbookMutationInput(mutation)).toBe(false)
      expect(isPendingWorkbookMutation(mutation)).toBe(false)
    }

    expect(
      isPendingWorkbookMutationInput({
        method: 'clearRangeStyle',
        args: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }],
      }),
    ).toBe(true)
    expect(
      isPendingWorkbookMutationInput({
        method: 'clearRangeStyle',
        args: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, ['fontBold']],
      }),
    ).toBe(true)
    expect(
      isPendingWorkbookMutationInput({
        method: 'clearRangeStyle',
        args: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, ['fontBold'], 'ignored'],
      }),
    ).toBe(false)
  })

  it('rejects extra ignored args before building Zero payloads', () => {
    for (const mutation of [
      { method: 'setCellValue' as const, args: ['Sheet1', 'A1', 17, { ignored: true }], error: 'Invalid setCellValue args' },
      { method: 'clearCell' as const, args: ['Sheet1', 'A1', 'ignored'], error: 'Invalid clearCell args' },
      {
        method: 'renderCommit' as const,
        args: [[{ kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 7 }], null],
        error: 'Invalid renderCommit args',
      },
      { method: 'insertRows' as const, args: ['Sheet1', 1, 1, 'ignored'], error: 'Invalid insertRows args' },
      { method: 'updateRowMetadata' as const, args: ['Sheet1', 1, 1, 24, null, 'ignored'], error: 'Invalid updateRowMetadata args' },
      { method: 'setFreezePane' as const, args: ['Sheet1', 1, 1, 'ignored'], error: 'Invalid setFreezePane args' },
      {
        method: 'clearRangeNumberFormat' as const,
        args: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, null],
        error: 'Invalid clearRangeNumberFormat args',
      },
      {
        method: 'clearRangeStyle' as const,
        args: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, ['fontBold'], 'ignored'],
        error: 'Invalid clearRangeStyle args',
      },
    ]) {
      expect(() => buildZeroWorkbookMutation('doc-1', mutation)).toThrow(mutation.error)
    }
  })

  it('rejects malformed style and number-format payloads', () => {
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }

    for (const mutation of [
      createPendingMutation({ method: 'setRangeStyle', args: [range, { font: { size: Number.NaN } }] }),
      createPendingMutation({ method: 'setRangeStyle', args: [range, { alignment: { horizontal: 'sideways' } }] }),
      createPendingMutation({ method: 'setRangeStyle', args: [range, { borders: { top: { style: 'wavy' } } }] }),
      createPendingMutation({ method: 'setRangeNumberFormat', args: [range, { kind: 'bogus' }] }),
      createPendingMutation({ method: 'setRangeNumberFormat', args: [range, { kind: 'number', decimals: 1.5 }] }),
      createPendingMutation({ method: 'setRangeNumberFormat', args: [range, { kind: 'currency', negativeStyle: 'red' }] }),
    ]) {
      expect(isPendingWorkbookMutationInput(mutation)).toBe(false)
      expect(isPendingWorkbookMutation(mutation)).toBe(false)
    }

    expect(isPendingWorkbookMutationInput(createPendingMutation({ method: 'setRangeStyle', args: [range, { font: { size: 12 } }] }))).toBe(
      true,
    )
    expect(
      isPendingWorkbookMutationInput(
        createPendingMutation({ method: 'setRangeNumberFormat', args: [range, { kind: 'number', decimals: 2 }] }),
      ),
    ).toBe(true)
  })

  it('rejects unknown clearRangeStyle fields', () => {
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }
    const mutation = createPendingMutation({
      method: 'clearRangeStyle',
      args: [range, ['fontBold', 'notAStyleField']],
    })

    expect(isPendingWorkbookMutationInput(mutation)).toBe(false)
    expect(isPendingWorkbookMutation(mutation)).toBe(false)
    expect(() =>
      buildZeroWorkbookMutation('doc-1', {
        method: 'clearRangeStyle',
        args: [range, ['fontBold', 'notAStyleField']],
      }),
    ).toThrow('Invalid clearRangeStyle args')
  })

  it('rejects malformed persisted pending mutation counters and timestamps', () => {
    expect(isPendingWorkbookMutation(createPendingMutation())).toBe(true)

    for (const mutation of [
      createPendingMutation({ id: '' }),
      createPendingMutation({ localSeq: 1.5 }),
      createPendingMutation({ baseRevision: -1 }),
      createPendingMutation({ enqueuedAtUnixMs: Number.NaN }),
      createPendingMutation({ submittedAtUnixMs: Number.POSITIVE_INFINITY }),
      createPendingMutation({ ackedAtUnixMs: -1 }),
      createPendingMutation({ attemptCount: 1.5 }),
      createPendingMutation({ attemptCount: -1 }),
    ]) {
      expect(isPendingWorkbookMutation(mutation)).toBe(false)
      expect(isPendingWorkbookMutationList([createPendingMutation(), mutation])).toBe(false)
    }
  })
})
