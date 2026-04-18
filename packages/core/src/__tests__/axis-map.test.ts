import { describe, expect, it } from 'vitest'
import { AxisMap } from '../storage/axis-map.js'

describe('AxisMap', () => {
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

    const removed = axisMap.splice(1, 1, [{ id: 'row-x', index: 1 }])
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
})
