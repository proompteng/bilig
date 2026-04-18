import { describe, expect, it } from 'vitest'
import { AxisMap } from '../storage/axis-map.js'

describe('AxisMap', () => {
  it('preserves sparse insert holes and stable ids at visible positions', () => {
    const axisMap = new AxisMap()

    const removed = axisMap.splice(0, 0, 3, [
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])

    expect(removed).toEqual([])
    expect(axisMap.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])
    expect(axisMap.snapshot(0, 3)).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])
    expect(axisMap.getId(1)).toBeUndefined()
    expect(axisMap.indexOf('row-c')).toBe(2)
  })

  it('replaces, snapshots, splices, and moves axis ids by visible index', () => {
    const axisMap = new AxisMap()

    axisMap.replaceRange(0, [
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
      { id: 'row-c', index: 2 },
    ])

    expect(axisMap.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
      { id: 'row-c', index: 2 },
    ])
    expect(axisMap.snapshot(1, 2)).toEqual([
      { id: 'row-b', index: 1 },
      { id: 'row-c', index: 2 },
    ])

    const removed = axisMap.splice(1, 1, 1, [{ id: 'row-x', index: 1 }])
    expect(removed).toEqual([{ id: 'row-b', index: 1 }])
    expect(axisMap.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-x', index: 1 },
      { id: 'row-c', index: 2 },
    ])

    axisMap.move(0, 1, 3)
    expect(axisMap.list()).toEqual([
      { id: 'row-x', index: 0 },
      { id: 'row-c', index: 1 },
      { id: 'row-a', index: 2 },
    ])
  })

  it('ignores holes when snapshotting or listing', () => {
    const axisMap = new AxisMap()

    axisMap.replaceRange(2, [{ id: 'column-c', index: 2 }])

    expect(axisMap.snapshot(0, 3)).toEqual([{ id: 'column-c', index: 2 }])
    expect(axisMap.list()).toEqual([{ id: 'column-c', index: 2 }])
  })

  it('ensures stable ids for visible positions', () => {
    const axisMap = new AxisMap()

    expect(axisMap.ensureId(2, () => 'row-a')).toBe('row-a')
    expect(axisMap.getId(2)).toBe('row-a')
    expect(axisMap.ensureId(2, () => 'row-b')).toBe('row-a')
    expect(axisMap.indexOf('row-a')).toBe(2)
    expect(axisMap.indexOf('row-missing')).toBe(-1)
  })
})
