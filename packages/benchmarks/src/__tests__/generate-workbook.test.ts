import { describe, expect, it } from 'vitest'
import { buildRangeAggregateSnapshot, buildTopologyEditSnapshot } from '../generate-workbook.js'

describe('benchmark workbook generators', () => {
  it('builds a range-heavy aggregate fixture with repeated bounded sums', () => {
    const snapshot = buildRangeAggregateSnapshot(4, 3)
    const cells = snapshot.sheets[0]?.cells ?? []

    expect(cells).toContainEqual({ address: 'A1', value: 1 })
    expect(cells).toContainEqual({ address: 'A4', value: 4 })
    expect(cells).toContainEqual({ address: 'B1', formula: 'SUM(A1:A4)+1' })
    expect(cells).toContainEqual({ address: 'B3', formula: 'SUM(A1:A4)+3' })
  })

  it('builds a topology-edit chain fixture whose head formula can be rebound', () => {
    const snapshot = buildTopologyEditSnapshot(4)
    const cells = snapshot.sheets[0]?.cells ?? []

    expect(cells).toEqual([
      { address: 'A1', value: 1 },
      { address: 'A2', value: 2 },
      { address: 'B1', formula: 'A1*2' },
      { address: 'B2', formula: 'B1+1' },
      { address: 'B3', formula: 'B2+1' },
      { address: 'B4', formula: 'B3+1' },
    ])
  })
})
