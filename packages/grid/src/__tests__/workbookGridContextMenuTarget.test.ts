import { describe, expect, it } from 'vitest'
import { createColumnSliceSelection, createRowSliceSelection } from '../gridSelection.js'
import { resolveKeyboardHeaderContextMenuTarget } from '../workbookGridContextMenuTarget.js'

describe('resolveKeyboardHeaderContextMenuTarget', () => {
  it('anchors row selections to the row header', () => {
    expect(
      resolveKeyboardHeaderContextMenuTarget({
        gridSelection: createRowSliceSelection(4, 7, 7),
        targetCellBounds: { x: 160, y: 240, width: 96, height: 28 },
        hostLeft: 32,
        hostTop: 48,
        rowMarkerWidth: 56,
        headerHeight: 32,
      }),
    ).toEqual({
      target: { kind: 'row', index: 7 },
      x: 60,
      y: 254,
    })
  })

  it('anchors column selections to the column header', () => {
    expect(
      resolveKeyboardHeaderContextMenuTarget({
        gridSelection: createColumnSliceSelection(3, 3, 5),
        targetCellBounds: { x: 212, y: 180, width: 88, height: 28 },
        hostLeft: 32,
        hostTop: 48,
        rowMarkerWidth: 56,
        headerHeight: 32,
      }),
    ).toEqual({
      target: { kind: 'column', index: 3 },
      x: 256,
      y: 64,
    })
  })

  it('returns null when the selection is not row-only or column-only', () => {
    expect(
      resolveKeyboardHeaderContextMenuTarget({
        gridSelection: {
          current: undefined,
          columns: createColumnSliceSelection(0, 1, 0).columns,
          rows: createRowSliceSelection(0, 1, 2).rows,
        },
        targetCellBounds: { x: 0, y: 0, width: 10, height: 10 },
        hostLeft: 0,
        hostTop: 0,
        rowMarkerWidth: 56,
        headerHeight: 32,
      }),
    ).toBeNull()
  })
})
