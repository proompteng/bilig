import { describe, expect, it } from 'vitest'
import { SheetAxisMap } from '../storage/sheet-axis-map.js'

describe('SheetAxisMap', () => {
  it('sets ids, finds ids, and supports both splice overloads', () => {
    const axisMap = new SheetAxisMap()

    axisMap.setId('row', 1, 'row-1')
    axisMap.setId('column', 0, 'column-0')

    expect(axisMap.getId('row', 1)).toBe('row-1')
    expect(axisMap.indexOf('column', 'column-0')).toBe(0)

    expect(axisMap.splice('row', 1, 0, [{ id: 'row-2', index: 1 }])).toEqual([])
    expect(axisMap.list('row')).toEqual([
      { id: 'row-2', index: 1 },
      { id: 'row-1', index: 2 },
    ])

    expect(axisMap.splice('column', 0, 1, 1, [{ id: 'column-1', index: 0 }])).toEqual([{ id: 'column-0', index: 0 }])
    expect(axisMap.list('column')).toEqual([{ id: 'column-1', index: 0 }])
  })
})
