import { describe, expect, it } from 'vitest'
import { ValueTag, type CellRangeRef, type CellSnapshot } from '@bilig/protocol'
import {
  applyOptimisticClearRange,
  applyOptimisticCommitOps,
  applyOptimisticCopyRange,
  applyOptimisticFillRange,
  applyOptimisticMoveRange,
  buildPasteCommitOps,
  createSheetScopedRangePair,
} from '../use-workbook-selection-actions.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'

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
      [
        'Sheet1:B2',
        {
          ...emptyCell('Sheet1', 'B2'),
          input: 12,
          format: '0.00',
          numberFormatId: 'fmt-body',
          styleId: 'style-fill',
          value: { tag: ValueTag.Number, value: 12 },
          version: 3,
        },
      ],
      [
        'Sheet1:C2',
        {
          ...emptyCell('Sheet1', 'C2'),
          input: 13,
          styleId: 'style-fill',
          value: { tag: ValueTag.Number, value: 13 },
          version: 4,
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

    const rollback = applyOptimisticClearRange(viewportStore, {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'C2',
    })

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:B2')).toMatchObject({
      format: '0.00',
      numberFormatId: 'fmt-body',
      styleId: 'style-fill',
      value: { tag: ValueTag.Empty },
      version: 4,
    })
    expect('input' in (cells.get('Sheet1:B2') ?? {})).toBe(false)
    expect(cells.get('Sheet1:C2')).toMatchObject({
      styleId: 'style-fill',
      value: { tag: ValueTag.Empty },
      version: 5,
    })
    expect('input' in (cells.get('Sheet1:C2') ?? {})).toBe(false)
  })

  it('clears cached cells for huge selected ranges without materializing every address', () => {
    const cached = {
      ...emptyCell('Sheet1', 'B1'),
      input: 'clear',
      numberFormatId: 'fmt-cached',
      styleId: 'style-cached-fill',
      value: { tag: ValueTag.String, value: 'clear' } as const,
      version: 3,
    }
    const writes: CellSnapshot[] = []
    const viewportStore = {
      forEachCellSnapshotInRange(
        _range: { sheetName: string; startAddress: string; endAddress: string },
        listener: (snapshot: CellSnapshot) => void,
      ) {
        listener(cached)
      },
      getCell() {
        throw new Error('huge optimistic clears must not materialize missing cells')
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        writes.push(snapshot)
      },
    }

    const rollback = applyOptimisticClearRange(viewportStore, {
      sheetName: 'Sheet1',
      startAddress: 'B1',
      endAddress: 'B1048576',
    })

    expect(rollback).toEqual(expect.any(Function))
    expect(writes).toHaveLength(1)
    expect(writes[0]).toMatchObject({
      address: 'B1',
      sheetName: 'Sheet1',
      numberFormatId: 'fmt-cached',
      styleId: 'style-cached-fill',
      value: { tag: ValueTag.Empty },
      version: 4,
    })
  })

  it('clears active visible cells for huge selected ranges so deleted tile content cannot reappear', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:B2',
        {
          ...emptyCell('Sheet1', 'B2'),
          input: 'cached-visible',
          styleId: 'style-visible-fill',
          value: { tag: ValueTag.String, value: 'cached-visible', stringId: 1 },
          version: 1,
        },
      ],
    ])
    const viewportStore = {
      forEachCellSnapshotInRange() {
        throw new Error('large clears must use cached-or-visible candidates when available')
      },
      forEachCachedOrVisibleCellSnapshotInRange(
        _range: { sheetName: string; startAddress: string; endAddress: string },
        listener: (snapshot: CellSnapshot) => void,
      ) {
        listener(cells.get('Sheet1:B2') ?? emptyCell('Sheet1', 'B2'))
        listener(emptyCell('Sheet1', 'C12'))
      },
      getCell() {
        throw new Error('huge optimistic clears must not materialize missing cells')
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticClearRange(viewportStore, {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'XFD1048576',
    })

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:B2')).toMatchObject({
      styleId: 'style-visible-fill',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 2,
    })
    expect(cells.get('Sheet1:C12')).toMatchObject({
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
  })

  it('marks moved target cells optimistic so stale patches cannot erase them', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:B2',
        { ...emptyCell('Sheet1', 'B2'), input: 'left', value: { tag: ValueTag.String, value: 'left', stringId: 1 }, version: 3 },
      ],
      [
        'Sheet1:C2',
        { ...emptyCell('Sheet1', 'C2'), input: 'right', value: { tag: ValueTag.String, value: 'right', stringId: 2 }, version: 4 },
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

    const rollback = applyOptimisticMoveRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C2' },
      { sheetName: 'Sheet1', startAddress: 'E5', endAddress: 'F5' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:B2')).toMatchObject({
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 4,
    })
    expect(cells.get('Sheet1:C2')).toMatchObject({
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 5,
    })
    expect(cells.get('Sheet1:E5')).toMatchObject({
      input: 'left',
      value: { tag: ValueTag.String, value: 'left', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 4,
    })
    expect(cells.get('Sheet1:F5')).toMatchObject({
      input: 'right',
      value: { tag: ValueTag.String, value: 'right', stringId: 2 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
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

  it('copies projected presentation and clears stale target formatting before worker patches arrive', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:B2',
        {
          ...emptyCell('Sheet1', 'B2'),
          input: 'styled-source',
          format: '0.00',
          numberFormatId: 'fmt-source',
          styleId: 'style-source',
          value: { tag: ValueTag.String, value: 'styled-source', stringId: 1 },
          version: 3,
        },
      ],
      [
        'Sheet1:D2',
        {
          ...emptyCell('Sheet1', 'D2'),
          input: 'styled-target',
          format: '$0',
          numberFormatId: 'fmt-target',
          styleId: 'style-target',
          value: { tag: ValueTag.String, value: 'styled-target', stringId: 2 },
          version: 4,
        },
      ],
      [
        'Sheet1:B3',
        {
          ...emptyCell('Sheet1', 'B3'),
          input: 'plain-source',
          value: { tag: ValueTag.String, value: 'plain-source', stringId: 3 },
          version: 5,
        },
      ],
      [
        'Sheet1:D3',
        {
          ...emptyCell('Sheet1', 'D3'),
          input: 'stale-style-target',
          format: '$0',
          numberFormatId: 'fmt-stale',
          styleId: 'style-stale',
          value: { tag: ValueTag.String, value: 'stale-style-target', stringId: 4 },
          version: 6,
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
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B3' },
      { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'D3' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:D2')).toMatchObject({
      input: 'styled-source',
      format: '0.00',
      numberFormatId: 'fmt-source',
      styleId: 'style-source',
      value: { tag: ValueTag.String, value: 'styled-source' },
    })
    expect(cells.get('Sheet1:D3')).toMatchObject({
      input: 'plain-source',
      value: { tag: ValueTag.String, value: 'plain-source' },
    })
    expect(cells.get('Sheet1:D3')?.format).toBeUndefined()
    expect(cells.get('Sheet1:D3')?.numberFormatId).toBeUndefined()
    expect(cells.get('Sheet1:D3')?.styleId).toBeUndefined()
  })

  it('fills projected ranges before worker patches arrive', () => {
    const cells = new Map<string, CellSnapshot>([
      ['Sheet1:F6', { ...emptyCell('Sheet1', 'F6'), input: 7, value: { tag: ValueTag.Number, value: 7 }, version: 3 }],
    ])
    const viewportStore = {
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticFillRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'F6', endAddress: 'F6' },
      { sheetName: 'Sheet1', startAddress: 'F7', endAddress: 'F8' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:F7')).toMatchObject({
      input: 7,
      value: { tag: ValueTag.Number, value: 7 },
      version: 1,
    })
    expect(cells.get('Sheet1:F8')).toMatchObject({
      input: 7,
      value: { tag: ValueTag.Number, value: 7 },
      version: 1,
    })
  })

  it('fills projected presentation as part of the immediate visual update', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:F6',
        {
          ...emptyCell('Sheet1', 'F6'),
          input: 7,
          format: '0.00',
          numberFormatId: 'fmt-source',
          styleId: 'style-source',
          value: { tag: ValueTag.Number, value: 7 },
          version: 3,
        },
      ],
      [
        'Sheet1:F7',
        {
          ...emptyCell('Sheet1', 'F7'),
          input: 'target',
          format: '$0',
          numberFormatId: 'fmt-target',
          styleId: 'style-target',
          value: { tag: ValueTag.String, value: 'target' },
          version: 4,
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

    const rollback = applyOptimisticFillRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'F6', endAddress: 'F6' },
      { sheetName: 'Sheet1', startAddress: 'F7', endAddress: 'F8' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:F7')).toMatchObject({
      input: 7,
      format: '0.00',
      numberFormatId: 'fmt-source',
      styleId: 'style-source',
      value: { tag: ValueTag.Number, value: 7 },
    })
    expect(cells.get('Sheet1:F8')).toMatchObject({
      input: 7,
      format: '0.00',
      numberFormatId: 'fmt-source',
      styleId: 'style-source',
      value: { tag: ValueTag.Number, value: 7 },
    })
  })

  it('copies visible cells for huge ranges without materializing offscreen targets', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:B1',
        {
          ...emptyCell('Sheet1', 'B1'),
          input: 'styled-source',
          format: '0.00',
          numberFormatId: 'fmt-source',
          styleId: 'style-source',
          value: { tag: ValueTag.String, value: 'styled-source' },
          version: 3,
        },
      ],
      [
        'Sheet1:B2',
        {
          ...emptyCell('Sheet1', 'B2'),
          input: 'plain-source',
          value: { tag: ValueTag.String, value: 'plain-source' },
          version: 4,
        },
      ],
      [
        'Sheet1:D1',
        {
          ...emptyCell('Sheet1', 'D1'),
          input: 'stale-target',
          styleId: 'style-stale',
          value: { tag: ValueTag.String, value: 'stale-target' },
          version: 5,
        },
      ],
      [
        'Sheet1:D2',
        {
          ...emptyCell('Sheet1', 'D2'),
          input: 'stale-target-two',
          styleId: 'style-stale-two',
          value: { tag: ValueTag.String, value: 'stale-target-two' },
          version: 6,
        },
      ],
    ])
    const visitedRanges: string[] = []
    const viewportStore = {
      forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void) {
        visitedRanges.push(`${range.startAddress}:${range.endAddress}`)
        if (range.startAddress !== 'D1') {
          throw new Error('huge copy should only enumerate cached-or-visible target cells')
        }
        listener(cells.get('Sheet1:D1') ?? emptyCell('Sheet1', 'D1'))
        listener(cells.get('Sheet1:D2') ?? emptyCell('Sheet1', 'D2'))
      },
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticCopyRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B20000' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D20000' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(visitedRanges).toEqual(['D1:D20000'])
    expect(cells.get('Sheet1:D1')).toMatchObject({
      input: 'styled-source',
      format: '0.00',
      numberFormatId: 'fmt-source',
      styleId: 'style-source',
      value: { tag: ValueTag.String, value: 'styled-source' },
    })
    expect(cells.get('Sheet1:D2')).toMatchObject({
      input: 'plain-source',
      value: { tag: ValueTag.String, value: 'plain-source' },
    })
    expect(cells.get('Sheet1:D2')?.styleId).toBeUndefined()
  })

  it('does not ghost empty source presentation into huge copy targets when the mapped source is offscreen', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:D15000',
        {
          ...emptyCell('Sheet1', 'D15000'),
          input: 'keep-visible-target',
          styleId: 'style-visible-target',
          value: { tag: ValueTag.String, value: 'keep-visible-target' },
          version: 5,
        },
      ],
    ])
    const writes: CellSnapshot[] = []
    const viewportStore = {
      forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void) {
        expect(range).toMatchObject({ startAddress: 'D1', endAddress: 'D20000' })
        listener(cells.get('Sheet1:D15000') ?? emptyCell('Sheet1', 'D15000'))
      },
      peekCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`)
      },
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        writes.push(snapshot)
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticCopyRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B20000' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D20000' },
    )

    expect(rollback).toBeNull()
    expect(writes).toEqual([])
    expect(cells.get('Sheet1:D15000')).toMatchObject({
      input: 'keep-visible-target',
      styleId: 'style-visible-target',
      value: { tag: ValueTag.String, value: 'keep-visible-target' },
    })
  })

  it('fills visible cells for huge ranges without materializing offscreen targets', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:B1',
        {
          ...emptyCell('Sheet1', 'B1'),
          input: 7,
          format: '0.00',
          numberFormatId: 'fmt-source',
          styleId: 'style-source',
          value: { tag: ValueTag.Number, value: 7 },
          version: 3,
        },
      ],
      [
        'Sheet1:D1',
        { ...emptyCell('Sheet1', 'D1'), input: 'target-one', value: { tag: ValueTag.String, value: 'target-one' }, version: 5 },
      ],
      [
        'Sheet1:D2',
        { ...emptyCell('Sheet1', 'D2'), input: 'target-two', value: { tag: ValueTag.String, value: 'target-two' }, version: 6 },
      ],
    ])
    const viewportStore = {
      forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void) {
        if (range.startAddress !== 'D1') {
          throw new Error('huge fill should only enumerate cached-or-visible target cells')
        }
        listener(cells.get('Sheet1:D1') ?? emptyCell('Sheet1', 'D1'))
        listener(cells.get('Sheet1:D2') ?? emptyCell('Sheet1', 'D2'))
      },
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticFillRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D20000' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(cells.get('Sheet1:D1')).toMatchObject({
      input: 7,
      format: '0.00',
      numberFormatId: 'fmt-source',
      styleId: 'style-source',
      value: { tag: ValueTag.Number, value: 7 },
    })
    expect(cells.get('Sheet1:D2')).toMatchObject({
      input: 7,
      format: '0.00',
      numberFormatId: 'fmt-source',
      styleId: 'style-source',
      value: { tag: ValueTag.Number, value: 7 },
    })
  })

  it('uses a range overlay for huge fill targets so future viewports inherit source presentation', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:B1',
        {
          ...emptyCell('Sheet1', 'B1'),
          input: 7,
          format: '0.00',
          numberFormatId: 'fmt-source',
          styleId: 'style-source',
          value: { tag: ValueTag.Number, value: 7 },
          version: 3,
        },
      ],
    ])
    let overlayApply: ((snapshot: CellSnapshot) => CellSnapshot) | null = null
    let rollbackCalled = false
    const viewportStore = {
      beginOptimisticRangeOverlay(range: CellRangeRef, apply: (snapshot: CellSnapshot) => CellSnapshot) {
        expect(range).toMatchObject({ sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D20000' })
        overlayApply = apply
        return () => {
          rollbackCalled = true
        }
      },
      forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void) {
        expect(range).toMatchObject({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' })
        listener(cells.get('Sheet1:B1') ?? emptyCell('Sheet1', 'B1'))
      },
      peekCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`)
      },
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticFillRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D20000' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(overlayApply).toEqual(expect.any(Function))
    const futureTarget = overlayApply?.(emptyCell('Sheet1', 'D15000'))
    expect(futureTarget).toMatchObject({
      input: 7,
      format: '0.00',
      numberFormatId: 'fmt-source',
      styleId: 'style-source',
      value: { tag: ValueTag.Number, value: 7 },
    })
    rollback?.()
    expect(rollbackCalled).toBe(true)
  })

  it('does not ghost empty source presentation into huge fill targets when the mapped source is offscreen', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:D15000',
        {
          ...emptyCell('Sheet1', 'D15000'),
          input: 'keep-filled-target',
          styleId: 'style-filled-target',
          value: { tag: ValueTag.String, value: 'keep-filled-target' },
          version: 6,
        },
      ],
    ])
    const writes: CellSnapshot[] = []
    const viewportStore = {
      forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void) {
        expect(range).toMatchObject({ startAddress: 'D1', endAddress: 'D20000' })
        listener(cells.get('Sheet1:D15000') ?? emptyCell('Sheet1', 'D15000'))
      },
      peekCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`)
      },
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        writes.push(snapshot)
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticFillRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B20000' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D20000' },
    )

    expect(rollback).toBeNull()
    expect(writes).toEqual([])
    expect(cells.get('Sheet1:D15000')).toMatchObject({
      input: 'keep-filled-target',
      styleId: 'style-filled-target',
      value: { tag: ValueTag.String, value: 'keep-filled-target' },
    })
  })

  it('moves visible cells for huge ranges and clears visible source presentation', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:B1',
        {
          ...emptyCell('Sheet1', 'B1'),
          input: 'move-source',
          format: '0.00',
          numberFormatId: 'fmt-source',
          styleId: 'style-source',
          value: { tag: ValueTag.String, value: 'move-source' },
          version: 3,
        },
      ],
      [
        'Sheet1:D1',
        { ...emptyCell('Sheet1', 'D1'), input: 'target-one', value: { tag: ValueTag.String, value: 'target-one' }, version: 5 },
      ],
    ])
    const visitedRanges: string[] = []
    const viewportStore = {
      forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void) {
        visitedRanges.push(`${range.startAddress}:${range.endAddress}`)
        if (range.startAddress === 'B1') {
          listener(cells.get('Sheet1:B1') ?? emptyCell('Sheet1', 'B1'))
          return
        }
        if (range.startAddress === 'D1') {
          listener(cells.get('Sheet1:D1') ?? emptyCell('Sheet1', 'D1'))
          return
        }
        throw new Error('unexpected huge move range enumeration')
      },
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticMoveRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B20000' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D20000' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(visitedRanges).toEqual(['B1:B20000', 'D1:D20000'])
    expect(cells.get('Sheet1:B1')).toMatchObject({
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 4,
    })
    expect(cells.get('Sheet1:B1')?.styleId).toBeUndefined()
    expect(cells.get('Sheet1:D1')).toMatchObject({
      input: 'move-source',
      format: '0.00',
      numberFormatId: 'fmt-source',
      styleId: 'style-source',
      value: { tag: ValueTag.String, value: 'move-source' },
    })
  })

  it('does not ghost empty source presentation into huge move targets when the mapped source is offscreen', () => {
    const cells = new Map<string, CellSnapshot>([
      [
        'Sheet1:B1',
        {
          ...emptyCell('Sheet1', 'B1'),
          input: 'visible-source-clear',
          styleId: 'style-source-clear',
          value: { tag: ValueTag.String, value: 'visible-source-clear' },
          version: 3,
        },
      ],
      [
        'Sheet1:D15000',
        {
          ...emptyCell('Sheet1', 'D15000'),
          input: 'keep-move-target',
          styleId: 'style-move-target',
          value: { tag: ValueTag.String, value: 'keep-move-target' },
          version: 7,
        },
      ],
    ])
    const writes: CellSnapshot[] = []
    const viewportStore = {
      forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void) {
        if (range.startAddress === 'B1') {
          listener(cells.get('Sheet1:B1') ?? emptyCell('Sheet1', 'B1'))
          return
        }
        expect(range).toMatchObject({ startAddress: 'D1', endAddress: 'D20000' })
        listener(cells.get('Sheet1:D15000') ?? emptyCell('Sheet1', 'D15000'))
      },
      peekCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`)
      },
      getCell(sheetName: string, address: string) {
        return cells.get(`${sheetName}:${address}`) ?? emptyCell(sheetName, address)
      },
      setCellSnapshot(snapshot: CellSnapshot) {
        writes.push(snapshot)
        cells.set(`${snapshot.sheetName}:${snapshot.address}`, snapshot)
      },
    }

    const rollback = applyOptimisticMoveRange(
      viewportStore,
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B20000' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D20000' },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(writes.map((snapshot) => snapshot.address)).toEqual(['B1'])
    expect(cells.get('Sheet1:B1')).toMatchObject({
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
    })
    expect(cells.get('Sheet1:D15000')).toMatchObject({
      input: 'keep-move-target',
      styleId: 'style-move-target',
      value: { tag: ValueTag.String, value: 'keep-move-target' },
    })
  })
})
