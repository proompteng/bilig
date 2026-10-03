import { describe, expect, it } from 'vitest'
import { collectWorkPaperRangeDependencies, toWorkPaperDependencyRefs } from '../work-paper-dependency-refs.js'
import type { WorkPaperCellAddress } from '../work-paper-types.js'

const resolver = {
  defaultSheetName: () => 'Sheet1',
  requireSheetId: (name: string) => {
    if (name === 'Sheet1') {
      return 1
    }
    if (name === 'Data') {
      return 2
    }
    throw new Error(`Unknown sheet ${name}`)
  },
}

describe('work paper dependency refs', () => {
  it('converts cell, range, and named dependency strings into public refs', () => {
    expect(toWorkPaperDependencyRefs(['A1', 'Data!B2:C3', 'RevenueRate'], resolver)).toEqual([
      {
        kind: 'cell',
        address: { sheet: 1, row: 0, col: 0 },
      },
      {
        kind: 'range',
        range: {
          start: { sheet: 2, row: 1, col: 1 },
          end: { sheet: 2, row: 2, col: 2 },
        },
      },
      { kind: 'name', name: 'RevenueRate' },
    ])
  })

  it('collects range dependencies in row-major order while deduping repeated refs', () => {
    const reads: WorkPaperCellAddress[] = []
    const dependencies = collectWorkPaperRangeDependencies({
      range: {
        start: { sheet: 1, row: 0, col: 0 },
        end: { sheet: 1, row: 1, col: 1 },
      },
      readDependencies: (address) => {
        reads.push(address)
        return address.row === 0 ? ['A1', 'SharedName'] : ['A1', 'Data!A1:B1']
      },
      resolver,
    })

    expect(reads).toEqual([
      { sheet: 1, row: 0, col: 0 },
      { sheet: 1, row: 0, col: 1 },
      { sheet: 1, row: 1, col: 0 },
      { sheet: 1, row: 1, col: 1 },
    ])
    expect(dependencies).toEqual([
      { kind: 'cell', address: { sheet: 1, row: 0, col: 0 } },
      { kind: 'name', name: 'SharedName' },
      {
        kind: 'range',
        range: {
          start: { sheet: 2, row: 0, col: 0 },
          end: { sheet: 2, row: 0, col: 1 },
        },
      },
    ])
  })
})
