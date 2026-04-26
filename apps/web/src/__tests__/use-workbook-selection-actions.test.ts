import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import {
  applyOptimisticClearRange,
  applyOptimisticCommitOps,
  applyOptimisticCopyRange,
  buildPasteCommitOps,
  createSheetScopedRangePair,
} from '../use-workbook-selection-actions.js'

function emptyCell(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  }
}

describe('use workbook selection action helpers', () => {
  it('builds paste commit ops for formulas, clears, booleans, numbers, and strings', () => {
    expect(
      buildPasteCommitOps('Sheet1', 'B2', [
        ['=SUM(A1:A2)', ''],
        ['TRUE', '42', 'text'],
      ]),
    ).toEqual([
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B2', formula: 'SUM(A1:A2)' },
      { kind: 'deleteCell', sheetName: 'Sheet1', addr: 'C2' },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B3', value: true },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'C3', value: 42 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'D3', value: 'text' },
    ])
  })

  it('preserves pasted whitespace and multiline literals instead of trimming them', () => {
    expect(
      buildPasteCommitOps('Sheet1', 'A1', [
        ['  SKU  ', ' line 1\nline 2 '],
        [' 42 ', ' TRUE '],
      ]),
    ).toEqual([
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: '  SKU  ' },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B1', value: ' line 1\nline 2 ' },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A2', value: ' 42 ' },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B2', value: ' TRUE ' },
    ])
  })

  it('creates source and target ranges scoped to one sheet', () => {
    expect(createSheetScopedRangePair('Sheet1', 'A1', 'B2', 'C3', 'D4')).toEqual({
      source: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
      target: {
        sheetName: 'Sheet1',
        startAddress: 'C3',
        endAddress: 'D4',
      },
    })
  })

  it('applies pasted commit ops to the projected viewport before worker patches arrive', () => {
    const cells = new Map<string, CellSnapshot>()
    const writes: CellSnapshot[] = []
    const viewportStore = {
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        writes.push(snapshot)
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticCommitOps(viewportStore, [
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'E5', value: 21 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'F5', value: 'ok' },
      { kind: 'deleteCell', sheetName: 'Sheet1', addr: 'E6' },
    ])

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:E5')).toMatchObject({
      value: { tag: ValueTag.Number, value: 21 },
      input: 21,
      version: 1,
    })
    expect(cells.get('Sheet1:F5')).toMatchObject({
      value: { tag: ValueTag.String, value: 'ok' },
      input: 'ok',
      version: 1,
    })
    expect(cells.get('Sheet1:E6')).toMatchObject({
      value: { tag: ValueTag.Empty },
      version: 1,
    })
    expect(writes).toHaveLength(3)
  })

  it('clears the projected selected range before worker patches arrive', () => {
    const cells = new Map<string, CellSnapshot>([
      ['Sheet1:B2', { ...emptyCell('Sheet1', 'B2'), input: 12, value: { tag: ValueTag.Number, value: 12 }, version: 3 }],
      ['Sheet1:C2', { ...emptyCell('Sheet1', 'C2'), input: 13, value: { tag: ValueTag.Number, value: 13 }, version: 4 }],
    ])
    const viewportStore = {
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticClearRange(viewportStore, {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'C2',
    })

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:B2')).toMatchObject({
      value: { tag: ValueTag.Empty },
      version: 4,
    })
    expect(cells.get('Sheet1:C2')).toMatchObject({
      value: { tag: ValueTag.Empty },
      version: 5,
    })
  })

  it('copies projected ranges with translated formulas before worker patches arrive', () => {
    const cells = new Map<string, CellSnapshot>([
      ['Sheet1:B2', { ...emptyCell('Sheet1', 'B2'), input: 3, value: { tag: ValueTag.Number, value: 3 }, version: 1 }],
      [
        'Sheet1:C2',
        {
          ...emptyCell('Sheet1', 'C2'),
          formula: 'B2*2',
          value: { tag: ValueTag.Number, value: 6 },
          version: 2,
        },
      ],
    ])
    const viewportStore = {
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticCopyRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C2' },
      { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'E2' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:D2')).toMatchObject({
      input: 3,
      value: { tag: ValueTag.Number, value: 3 },
      version: 1,
    })
    expect(cells.get('Sheet1:E2')).toMatchObject({
      formula: 'D2*2',
      value: { tag: ValueTag.Number, value: 6 },
      version: 1,
    })
  })
})
