import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot, type Viewport } from '@bilig/protocol'
import { selectProjectedViewportKeysToEvict } from '../projected-viewport-cache-pruning.js'

function createSnapshot(address: string, value: number): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    value: { tag: ValueTag.Number, value },
    flags: 0,
    version: 1,
  }
}

describe('projected viewport cache pruning', () => {
  it('evicts the oldest offscreen unpinned cells first', () => {
    const cellSnapshots = new Map<string, CellSnapshot>([
      ['Sheet1!A1', createSnapshot('A1', 1)],
      ['Sheet1!A2', createSnapshot('A2', 2)],
      ['Sheet1!A3', createSnapshot('A3', 3)],
      ['Sheet1!A4', createSnapshot('A4', 4)],
    ])
    const activeViewports: Viewport[] = [
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
    ]

    const keys = selectProjectedViewportKeysToEvict({
      sheetCellKeys: ['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3', 'Sheet1!A4'],
      cellSnapshots,
      cellAccessTicks: new Map([
        ['Sheet1!A1', 9],
        ['Sheet1!A2', 1],
        ['Sheet1!A3', 3],
        ['Sheet1!A4', 7],
      ]),
      pinnedKeys: new Set(['Sheet1!A4']),
      activeViewports,
      maxCachedCellsPerSheet: 2,
    })

    expect(keys).toEqual(['Sheet1!A2', 'Sheet1!A3'])
  })

  it('drops missing snapshot keys before evicting intact cells', () => {
    const keys = selectProjectedViewportKeysToEvict({
      sheetCellKeys: ['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3'],
      cellSnapshots: new Map([
        ['Sheet1!A1', createSnapshot('A1', 1)],
        ['Sheet1!A3', createSnapshot('A3', 3)],
      ]),
      cellAccessTicks: new Map([
        ['Sheet1!A1', 10],
        ['Sheet1!A3', 20],
      ]),
      pinnedKeys: new Set(),
      activeViewports: [],
      maxCachedCellsPerSheet: 2,
    })

    expect(keys).toEqual(['Sheet1!A2'])
  })

  it('keeps pinned and onscreen cells even when the cache is oversized', () => {
    const keys = selectProjectedViewportKeysToEvict({
      sheetCellKeys: ['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3'],
      cellSnapshots: new Map([
        ['Sheet1!A1', createSnapshot('A1', 1)],
        ['Sheet1!A2', createSnapshot('A2', 2)],
        ['Sheet1!A3', createSnapshot('A3', 3)],
      ]),
      cellAccessTicks: new Map([
        ['Sheet1!A1', 1],
        ['Sheet1!A2', 2],
        ['Sheet1!A3', 3],
      ]),
      pinnedKeys: new Set(['Sheet1!A2']),
      activeViewports: [
        {
          sheetName: 'Sheet1',
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        },
      ],
      maxCachedCellsPerSheet: 1,
    })

    expect(keys).toEqual(['Sheet1!A3'])
  })
})
