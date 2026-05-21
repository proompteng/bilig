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
  const normalizedSelection = normalizeGridSelection(selection)
  const selectedColumnStart = normalizedSelection.columns.first()
  const selectedColumnEnd = normalizedSelection.columns.last()
  const selectedRowStart = normalizedSelection.rows.first()
  const selectedRowEnd = normalizedSelection.rows.last()

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

  const range = normalizedSelection.current?.range
  if (!range) {
    return {
      startAddress: fallbackAddress,
      endAddress: fallbackAddress,
    }
  }

  return rectangleToAddresses(range)
}

export function selectionToSnapshot(selection: GridSelection, sheetName: string, fallbackAddress: string): GridSelectionSnapshot {
  const normalizedSelection = normalizeGridSelection(selection)
  const range = selectionToAddresses(normalizedSelection, fallbackAddress)
  const currentCell = normalizedSelection.current?.cell
  const address = currentCell ? formatAddress(currentCell[1], currentCell[0]) : fallbackAddress
  let kind: GridSelectionKind = 'cell'
  const selectedColumnStart = normalizedSelection.columns.first()
  const selectedColumnEnd = normalizedSelection.columns.last()
  const selectedRowStart = normalizedSelection.rows.first()
  const selectedRowEnd = normalizedSelection.rows.last()

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
    case 'sheet': {
      const active = parseCellAddress(selection.address, selection.sheetName)
      return {
        ...createSheetSelection(),
        current: {
          cell: [active.col, active.row],
          range: { x: 0, y: 0, width: MAX_COLS, height: MAX_ROWS },
          rangeStack: [],
        },
      }
    }
    case 'column': {
      const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
      const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
      const active = parseCellAddress(selection.address, selection.sheetName)
      const left = Math.min(start.col, end.col)
      const right = Math.max(start.col, end.col)
      if (active.col < left || active.col > right) {
        return createColumnSliceSelection(start.col, end.col, start.row)
      }
      return {
        ...createColumnSliceSelection(start.col, end.col, active.row),
        current: {
          cell: [active.col, active.row],
          range: { x: left, y: active.row, width: right - left + 1, height: 1 },
          rangeStack: [],
        },
      }
    }
    case 'row': {
      const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
      const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
      const active = parseCellAddress(selection.address, selection.sheetName)
      const top = Math.min(start.row, end.row)
      const bottom = Math.max(start.row, end.row)
      if (active.row < top || active.row > bottom) {
        return createRowSliceSelection(start.col, start.row, end.row)
      }
      return {
        ...createRowSliceSelection(active.col, start.row, end.row),
        current: {
          cell: [active.col, active.row],
          range: { x: active.col, y: top, width: 1, height: bottom - top + 1 },
          rangeStack: [],
        },
      }
    }
    case 'range': {
      const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
      const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
      const active = parseCellAddress(selection.address, selection.sheetName)
      const range = {
        x: Math.min(start.col, end.col),
        y: Math.min(start.row, end.row),
        width: Math.abs(end.col - start.col) + 1,
        height: Math.abs(end.row - start.row) + 1,
      }
      const activeInsideRange =
        active.col >= range.x && active.col < range.x + range.width && active.row >= range.y && active.row < range.y + range.height
      if (!activeInsideRange) {
        return createRectangleSelectionFromRange(range)
      }
      return {
        ...createRectangleSelectionFromRange(range),
        current: {
          cell: [active.col, active.row],
          range,
          rangeStack: [],
        },
      }
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

export function gridSelectionCurrentCellInRange(selection: GridSelection): boolean {
  const current = selection.current
  if (!current) {
    return true
  }
  const [col, row] = current.cell
  const range = current.range
  return col >= range.x && col < range.x + range.width && row >= range.y && row < range.y + range.height
}

export function gridSelectionCurrentCellInAxisSelections(selection: GridSelection): boolean {
  const current = selection.current
  if (!current) {
    return true
  }
  const [col, row] = current.cell
  if (selection.columns.length > 0 && !selection.columns.hasIndex(col)) {
    return false
  }
  if (selection.rows.length > 0 && !selection.rows.hasIndex(row)) {
    return false
  }
  return true
}

export function isGridSelectionCoherent(selection: GridSelection): boolean {
  return gridSelectionCurrentCellInRange(selection) && gridSelectionCurrentCellInAxisSelections(selection)
}

export function normalizeGridSelection(selection: GridSelection): GridSelection {
  if (isGridSelectionCoherent(selection)) {
    return selection
  }
  const cell = selection.current?.cell
  if (!cell) {
    return selection
  }
  const [col, row] = clampCell(cell)
  return createGridSelection(col, row)
}

export function formatSelectionSummary(selection: GridSelection, fallbackAddress: string): string {
  const normalizedSelection = normalizeGridSelection(selection)
  const selectedColumnStart = normalizedSelection.columns.first()
  const selectedColumnEnd = normalizedSelection.columns.last()
  const selectedRowStart = normalizedSelection.rows.first()
  const selectedRowEnd = normalizedSelection.rows.last()
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

  const range = normalizedSelection.current?.range
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

export function formatSelectionSnapshotSummary(selection: GridSelectionSnapshot): string {
  switch (selection.kind) {
    case 'sheet':
      return 'All'
    case 'column': {
      const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
      const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
      const startLabel = indexToColumn(start.col)
      const endLabel = indexToColumn(end.col)
      return startLabel === endLabel ? `${startLabel}:${startLabel}` : `${startLabel}:${endLabel}`
    }
    case 'row': {
      const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
      const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
      const startLabel = String(start.row + 1)
      const endLabel = String(end.row + 1)
      return startLabel === endLabel ? `${startLabel}:${startLabel}` : `${startLabel}:${endLabel}`
    }
    case 'range':
      return selection.range.startAddress === selection.range.endAddress
        ? selection.range.startAddress
        : `${selection.range.startAddress}:${selection.range.endAddress}`
    case 'cell':
      return selection.address
  }
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
