import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { formatAddress, indexToColumn, parseCellAddress } from '@bilig/formula'
import {
  CompactSelection,
  type GridSelection,
  type GridSelectionKind,
  type GridSelectionSnapshot,
  type Item,
  type Rectangle,
} from './gridTypes.js'

export function createGridSelection(col: number, row: number): GridSelection {
  return {
    current: {
      cell: [col, row],
      range: { x: col, y: row, width: 1, height: 1 },
      rangeStack: [],
    },
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  }
}

export function clampCell([col, row]: Item): Item {
  return [Math.min(MAX_COLS - 1, Math.max(0, col)), Math.min(MAX_ROWS - 1, Math.max(0, row))]
}

export function clampSelectionRange(range: Rectangle): Rectangle {
  const x = Math.min(MAX_COLS - 1, Math.max(0, range.x))
  const y = Math.min(MAX_ROWS - 1, Math.max(0, range.y))
  const maxWidth = MAX_COLS - x
  const maxHeight = MAX_ROWS - y
  return {
    x,
    y,
    width: Math.max(1, Math.min(maxWidth, range.width)),
    height: Math.max(1, Math.min(maxHeight, range.height)),
  }
}

export function rectangleToAddresses(range: Rectangle): {
  startAddress: string
  endAddress: string
} {
  const clamped = clampSelectionRange(range)
  return {
    startAddress: formatAddress(clamped.y, clamped.x),
    endAddress: formatAddress(clamped.y + clamped.height - 1, clamped.x + clamped.width - 1),
  }
}

export function selectionToAddresses(
  selection: GridSelection,
  fallbackAddress: string,
): {
  startAddress: string
  endAddress: string
} {
  const selectedColumnStart = selection.columns.first()
  const selectedColumnEnd = selection.columns.last()
  const selectedRowStart = selection.rows.first()
  const selectedRowEnd = selection.rows.last()

  if (selectedColumnStart === 0 && selectedColumnEnd === MAX_COLS - 1 && selectedRowStart === 0 && selectedRowEnd === MAX_ROWS - 1) {
    return {
      startAddress: 'A1',
      endAddress: formatAddress(MAX_ROWS - 1, MAX_COLS - 1),
    }
  }

  if (selectedColumnStart !== undefined && selectedColumnEnd !== undefined) {
    return {
      startAddress: formatAddress(0, Math.min(selectedColumnStart, selectedColumnEnd)),
      endAddress: formatAddress(MAX_ROWS - 1, Math.max(selectedColumnStart, selectedColumnEnd)),
    }
  }

  if (selectedRowStart !== undefined && selectedRowEnd !== undefined) {
    return {
      startAddress: formatAddress(Math.min(selectedRowStart, selectedRowEnd), 0),
      endAddress: formatAddress(Math.max(selectedRowStart, selectedRowEnd), MAX_COLS - 1),
    }
  }

  const range = selection.current?.range
  if (!range) {
    return {
      startAddress: fallbackAddress,
      endAddress: fallbackAddress,
    }
  }

  return rectangleToAddresses(range)
}

export function selectionToSnapshot(selection: GridSelection, sheetName: string, fallbackAddress: string): GridSelectionSnapshot {
  const range = selectionToAddresses(selection, fallbackAddress)
  const currentCell = selection.current?.cell
  const address = currentCell ? formatAddress(currentCell[1], currentCell[0]) : fallbackAddress
  let kind: GridSelectionKind = 'cell'
  const selectedColumnStart = selection.columns.first()
  const selectedColumnEnd = selection.columns.last()
  const selectedRowStart = selection.rows.first()
  const selectedRowEnd = selection.rows.last()

  if (selectedColumnStart === 0 && selectedColumnEnd === MAX_COLS - 1 && selectedRowStart === 0 && selectedRowEnd === MAX_ROWS - 1) {
    kind = 'sheet'
  } else if (selectedColumnStart !== undefined && selectedColumnEnd !== undefined) {
    kind = 'column'
  } else if (selectedRowStart !== undefined && selectedRowEnd !== undefined) {
    kind = 'row'
  } else if (range.startAddress !== range.endAddress) {
    kind = 'range'
  }

  return {
    sheetName,
    address,
    kind,
    range,
  }
}

export function snapshotToSelection(selection: GridSelectionSnapshot): GridSelection {
  switch (selection.kind) {
    case 'sheet':
      return createSheetSelection()
    case 'column': {
      const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
      const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
      return createColumnSliceSelection(start.col, end.col, start.row)
    }
    case 'row': {
      const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
      const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
      return createRowSliceSelection(start.col, start.row, end.row)
    }
    case 'range': {
      const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
      const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
      return createRectangleSelectionFromRange({
        x: Math.min(start.col, end.col),
        y: Math.min(start.row, end.row),
        width: Math.abs(end.col - start.col) + 1,
        height: Math.abs(end.row - start.row) + 1,
      })
    }
    case 'cell': {
      const parsed = parseCellAddress(selection.address, selection.sheetName)
      return createGridSelection(parsed.col, parsed.row)
    }
  }
}

export function createRangeSelection(base: GridSelection, anchor: Item, target: Item): GridSelection {
  const startCol = Math.min(anchor[0], target[0])
  const endCol = Math.max(anchor[0], target[0])
  const startRow = Math.min(anchor[1], target[1])
  const endRow = Math.max(anchor[1], target[1])

  return {
    ...base,
    current: {
      cell: anchor,
      range: {
        x: startCol,
        y: startRow,
        width: endCol - startCol + 1,
        height: endRow - startRow + 1,
      },
      rangeStack: [],
    },
  }
}

export function createRectangleSelectionFromRange(range: Rectangle): GridSelection {
  return {
    current: {
      cell: [range.x, range.y],
      range: { ...range },
      rangeStack: [],
    },
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  }
}

export function formatSelectionSummary(selection: GridSelection, fallbackAddress: string): string {
  const selectedColumnStart = selection.columns.first()
  const selectedColumnEnd = selection.columns.last()
  const selectedRowStart = selection.rows.first()
  const selectedRowEnd = selection.rows.last()
  if (selectedColumnStart === 0 && selectedColumnEnd === MAX_COLS - 1 && selectedRowStart === 0 && selectedRowEnd === MAX_ROWS - 1) {
    return 'All'
  }

  if (selectedColumnStart !== undefined && selectedColumnEnd !== undefined) {
    const start = indexToColumn(selectedColumnStart)
    const end = indexToColumn(selectedColumnEnd)
    return start === end ? `${start}:${start}` : `${start}:${end}`
  }

  if (selectedRowStart !== undefined && selectedRowEnd !== undefined) {
    const start = String(selectedRowStart + 1)
    const end = String(selectedRowEnd + 1)
    return start === end ? `${start}:${start}` : `${start}:${end}`
  }

  const range = selection.current?.range
  if (!range) {
    return fallbackAddress
  }
  const start = formatAddress(range.y, range.x)
  if (range.width === 1 && range.height === 1) {
    return start
  }
  const end = formatAddress(range.y + range.height - 1, range.x + range.width - 1)
  return `${start}:${end}`
}

export function createColumnSelection(col: number, row: number): GridSelection {
  return {
    current: {
      cell: [col, row],
      range: { x: col, y: row, width: 1, height: 1 },
      rangeStack: [],
    },
    columns: CompactSelection.fromSingleSelection(col),
    rows: CompactSelection.empty(),
  }
}

export function createColumnSliceSelection(startCol: number, endCol: number, row: number): GridSelection {
  const left = Math.min(startCol, endCol)
  const right = Math.max(startCol, endCol)
  return {
    current: {
      cell: [startCol, row],
      range: { x: left, y: row, width: right - left + 1, height: 1 },
      rangeStack: [],
    },
    columns: CompactSelection.fromSingleSelection([left, right + 1]),
    rows: CompactSelection.empty(),
  }
}

export function createRowSelection(col: number, row: number): GridSelection {
  return {
    current: {
      cell: [col, row],
      range: { x: col, y: row, width: 1, height: 1 },
      rangeStack: [],
    },
    columns: CompactSelection.empty(),
    rows: CompactSelection.fromSingleSelection(row),
  }
}

export function createRowSliceSelection(col: number, startRow: number, endRow: number): GridSelection {
  const top = Math.min(startRow, endRow)
  const bottom = Math.max(startRow, endRow)
  return {
    current: {
      cell: [col, startRow],
      range: { x: col, y: top, width: 1, height: bottom - top + 1 },
      rangeStack: [],
    },
    columns: CompactSelection.empty(),
    rows: CompactSelection.fromSingleSelection([top, bottom + 1]),
  }
}

export function createSheetSelection(): GridSelection {
  return {
    current: {
      cell: [0, 0],
      range: { x: 0, y: 0, width: MAX_COLS, height: MAX_ROWS },
      rangeStack: [],
    },
    columns: CompactSelection.fromSingleSelection([0, MAX_COLS]),
    rows: CompactSelection.fromSingleSelection([0, MAX_ROWS]),
  }
}

export function isSheetSelection(selection: GridSelection): boolean {
  return (
    selection.columns.first() === 0 &&
    selection.columns.last() === MAX_COLS - 1 &&
    selection.rows.first() === 0 &&
    selection.rows.last() === MAX_ROWS - 1
  )
}

export function sameItem(left: Item | null, right: Item | null): boolean {
  return left !== null && right !== null && left[0] === right[0] && left[1] === right[1]
}
