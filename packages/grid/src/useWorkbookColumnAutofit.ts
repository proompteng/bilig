import { useCallback, useRef } from 'react'
import { formatAddress, indexToColumn } from '@bilig/formula'
import { MAX_ROWS, type CellSnapshot } from '@bilig/protocol'
import { getResolvedCellFontFamily, snapshotToRenderCell } from './gridCells.js'
import { MAX_COLUMN_WIDTH, MIN_COLUMN_WIDTH, type GridMetrics } from './gridMetrics.js'
import type { GridEngineLike } from './grid-engine.js'
import type { VisibleRegionState } from './gridPointer.js'

export interface WorkbookColumnTextMeasurer {
  font: string
  measureText(text: string): Pick<TextMetrics, 'width'>
}

export function measureWorkbookColumnAutofit(input: {
  readonly columnIndex: number
  readonly editorFontSize: string
  readonly engine: GridEngineLike
  readonly freezeRows: number
  readonly getCellEditorSeed?: ((sheetName: string, address: string) => string | undefined) | undefined
  readonly getVisibleRegion: () => VisibleRegionState
  readonly headerFontStyle: string
  readonly measurer: WorkbookColumnTextMeasurer
  readonly selectedCell: {
    readonly col: number
    readonly row: number
  }
  readonly selectedCellSnapshot: CellSnapshot
  readonly sheetName: string
}): number {
  const {
    columnIndex,
    editorFontSize,
    engine,
    freezeRows,
    getCellEditorSeed,
    getVisibleRegion,
    headerFontStyle,
    measurer,
    selectedCell,
    selectedCellSnapshot,
    sheetName,
  } = input
  let measuredWidth = 0

  measurer.font = headerFontStyle
  measuredWidth = Math.max(measuredWidth, measurer.measureText(indexToColumn(columnIndex)).width)

  const sheet = engine.workbook.getSheet(sheetName)
  const measureCell = (row: number, col: number) => {
    const address = formatAddress(row, col)
    const optimisticSeed = getCellEditorSeed?.(sheetName, address)
    if (optimisticSeed !== undefined) {
      measurer.font = `400 ${editorFontSize} ${getResolvedCellFontFamily()}`
      measuredWidth = Math.max(measuredWidth, measurer.measureText(optimisticSeed).width)
      return
    }
    const snapshot =
      col === selectedCell.col &&
      row === selectedCell.row &&
      selectedCellSnapshot.sheetName === sheetName &&
      selectedCellSnapshot.address === address
        ? selectedCellSnapshot
        : engine.getCell(sheetName, address)
    const renderCell = snapshotToRenderCell(snapshot, engine.getCellStyle(snapshot.styleId))
    const displayText = renderCell.displayText || renderCell.copyText
    measurer.font = `400 ${editorFontSize} ${getResolvedCellFontFamily()}`
    measuredWidth = Math.max(measuredWidth, measurer.measureText(displayText).width)
  }

  if (sheet) {
    sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
      if (col !== columnIndex) {
        return
      }
      measureCell(row, col)
    })
  } else {
    const liveVisibleRegion = getVisibleRegion()
    const visibleStartRow = liveVisibleRegion.range.y
    const visibleEndRow = Math.min(MAX_ROWS - 1, liveVisibleRegion.range.y + liveVisibleRegion.range.height - 1)
    for (let row = 0; row < Math.max(0, freezeRows); row += 1) {
      measureCell(row, columnIndex)
    }
    for (let row = visibleStartRow; row <= visibleEndRow; row += 1) {
      if (row < freezeRows) {
        continue
      }
      measureCell(row, columnIndex)
    }
  }

  return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.ceil(measuredWidth + 28)))
}

export function useWorkbookColumnAutofit(input: {
  readonly editorFontSize: string
  readonly engine: GridEngineLike
  readonly freezeRows: number
  readonly getCellEditorSeed?: ((sheetName: string, address: string) => string | undefined) | undefined
  readonly getVisibleRegion: () => VisibleRegionState
  readonly gridMetrics: GridMetrics
  readonly headerFontStyle: string
  readonly selectedCell: {
    readonly col: number
    readonly row: number
  }
  readonly selectedCellSnapshot: CellSnapshot
  readonly sheetName: string
}): (columnIndex: number) => number {
  const {
    editorFontSize,
    engine,
    freezeRows,
    getCellEditorSeed,
    getVisibleRegion,
    gridMetrics,
    headerFontStyle,
    selectedCell,
    selectedCellSnapshot,
    sheetName,
  } = input
  const textMeasureCanvasRef = useRef<HTMLCanvasElement | null>(null)
  const fallbackColumnWidth = gridMetrics.columnWidth

  return useCallback(
    (columnIndex: number): number => {
      const canvas = textMeasureCanvasRef.current ?? document.createElement('canvas')
      textMeasureCanvasRef.current = canvas
      const context = canvas.getContext('2d')
      if (!context) {
        return fallbackColumnWidth
      }

      return measureWorkbookColumnAutofit({
        columnIndex,
        editorFontSize,
        engine,
        freezeRows,
        getCellEditorSeed,
        getVisibleRegion,
        headerFontStyle,
        measurer: context,
        selectedCell,
        selectedCellSnapshot,
        sheetName,
      })
    },
    [
      editorFontSize,
      engine,
      freezeRows,
      getCellEditorSeed,
      getVisibleRegion,
      fallbackColumnWidth,
      headerFontStyle,
      selectedCell,
      selectedCellSnapshot,
      sheetName,
    ],
  )
}
