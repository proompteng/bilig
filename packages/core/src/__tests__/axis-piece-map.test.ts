import { describe, expect, it } from 'vitest'
import { AxisPieceMap } from '../storage/axis-piece-map.js'

describe('AxisPieceMap', () => {
  it('splices sparse axis ids while preserving stable indexes for listed ids', () => {
    const axis = new AxisPieceMap()

    expect(
      axis.splice(0, 0, 4, [
        { id: 'row-a', index: 0 },
        { id: 'row-c', index: 2 },
      ]),
    ).toEqual([])

    expect(axis.length).toBe(4)
    expect(axis.getId(1)).toBeUndefined()
    expect(axis.idAt(2)).toBe('row-c')
    expect(axis.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])
    expect(axis.snapshot(0, 3)).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])
  })

  it('ensures, replaces, deletes, and moves ids without rewriting ids themselves', () => {
    const axis = new AxisPieceMap()

    expect(axis.ensureId(1, () => 'row-b')).toBe('row-b')
    expect(axis.ensureId(1, () => 'row-x')).toBe('row-b')
    axis.replaceRange(0, [
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])

    expect(axis.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
      { id: 'row-c', index: 2 },
    ])

    expect(axis.move(0, 1, 3)).toEqual([{ id: 'row-a', index: 0 }])
    expect(axis.list()).toEqual([
      { id: 'row-b', index: 0 },
      { id: 'row-c', index: 1 },
      { id: 'row-a', index: 2 },
    ])

    expect(axis.delete(1, 1)).toEqual([{ id: 'row-c', index: 1 }])
    expect(axis.list()).toEqual([
      { id: 'row-b', index: 0 },
      { id: 'row-a', index: 1 },
    ])
  })

  it('rejects duplicate ids outside the deleted span', () => {
    const axis = new AxisPieceMap()
    axis.replace(['column-a', 'column-b'])

    expect(() => axis.splice(1, 0, ['column-a'])).toThrow('Axis id already exists: column-a')
    expect(() => axis.replace(['column-x', 'column-x'])).toThrow('Axis ids must be unique')
  })
})
