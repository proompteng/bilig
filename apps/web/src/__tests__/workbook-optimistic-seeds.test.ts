import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { supersedeOptimisticSeedsInRange } from '../workbook-optimistic-seeds.js'

const wholeSheet = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'XFD1048576' }

function guardBlankCellScans(seeds: Map<string, string>) {
  const read = seeds.get.bind(seeds)
  let reads = 0
  return vi.spyOn(seeds, 'get').mockImplementation((key) => {
    if (++reads > 10) throw new Error('Clear scanned blank addresses')
    return read(key)
  })
}

describe('superseding pending edits in a range', () => {
  it('clears an empty whole sheet without reading blank addresses', () => {
    const seeds = new Map<string, string>()
    const reads = guardBlankCellScans(seeds)
    expect(supersedeOptimisticSeedsInRange(wholeSheet, seeds, new Map(), new Map())).toBeNull()
    expect(reads).not.toHaveBeenCalled()
  })

  it('clears sparse pending edits across the whole sheet and keeps another sheet', () => {
    const seeds = new Map([
      ['Sheet1:A1', '1'],
      ['Sheet1:XFD1048576', '2'],
      ['Other:A1', 'keep'],
    ])
    guardBlankCellScans(seeds)
    const rollback = supersedeOptimisticSeedsInRange(wholeSheet, seeds, new Map(), new Map())
    expect([...seeds]).toEqual([['Other:A1', 'keep']])
    rollback?.()
    expect(seeds.get('Sheet1:XFD1048576')).toBe('2')
  })

  it('normalizes reversed bounds and restores pending values and snapshots', () => {
    const snapshot: CellSnapshot = { sheetName: 'Sheet1', address: 'C3', value: { tag: ValueTag.Number, value: 5 }, flags: 0, version: 1 }
    const seeds = new Map([
      ['Sheet1:C3', '=2+3'],
      ['Sheet1:A1', 'keep'],
      ['Sheet1:archive:C3', 'other sheet'],
    ])
    const resolvedValues = new Map([['Sheet1:C3', '5']])
    const snapshots = new Map([['Sheet1:C3', snapshot]])
    const rollback = supersedeOptimisticSeedsInRange(
      { sheetName: 'Sheet1', startAddress: 'D4', endAddress: 'B2' },
      seeds,
      resolvedValues,
      snapshots,
    )
    expect(seeds.has('Sheet1:C3')).toBe(false)
    expect(seeds.size).toBe(2)
    expect(resolvedValues.size).toBe(0)
    expect(snapshots.size).toBe(0)
    rollback?.()
    expect(seeds.get('Sheet1:C3')).toBe('=2+3')
    expect(resolvedValues.get('Sheet1:C3')).toBe('5')
    expect(snapshots.get('Sheet1:C3')).toBe(snapshot)
  })

  it('does not mix rolled back values into a newer edit', () => {
    const seeds = new Map([['Sheet1:A1', 'old']])
    const resolvedValues = new Map([['Sheet1:A1', 'old value']])
    const snapshots = new Map<string, CellSnapshot>()
    const rollback = supersedeOptimisticSeedsInRange({ ...wholeSheet, endAddress: 'A1' }, seeds, resolvedValues, snapshots)
    seeds.set('Sheet1:A1', 'new')
    rollback?.()
    expect(seeds.get('Sheet1:A1')).toBe('new')
    expect(resolvedValues.has('Sheet1:A1')).toBe(false)
  })
})
