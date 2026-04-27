import { useCallback, useEffect, useMemo, useRef, useState, type Dispatch, type SetStateAction } from 'react'
import {
  EMPTY_COLUMN_WIDTHS,
  EMPTY_ROW_HEIGHTS,
  MAX_COLUMN_WIDTH,
  MAX_ROW_HEIGHT,
  MIN_COLUMN_WIDTH,
  MIN_ROW_HEIGHT,
} from './gridMetrics.js'
import { applyHiddenAxisSizes } from './gridScrollSurface.js'

type AxisSizeOverrides = Readonly<Record<number, number>>

interface ColumnResizePreview {
  readonly sheetName: string
  readonly columnIndex: number
  readonly width: number
}

interface RowResizePreview {
  readonly sheetName: string
  readonly rowIndex: number
  readonly height: number
}

export interface WorkbookAxisResizeState {
  readonly activeResizeColumn: number | null
  readonly activeResizeRow: number | null
  readonly clearColumnResizePreview: (columnIndex: number) => void
  readonly clearRowResizePreview: (rowIndex: number) => void
  readonly columnWidths: AxisSizeOverrides
  readonly commitColumnWidth: (columnIndex: number, newSize: number) => void
  readonly commitRowHeight: (rowIndex: number, newSize: number) => void
  readonly getPreviewColumnWidth: (columnIndex: number) => number | null
  readonly getPreviewRowHeight: (rowIndex: number) => number | null
  readonly hasColumnResizePreview: boolean
  readonly hasRowResizePreview: boolean
  readonly previewColumnWidth: (columnIndex: number, newSize: number) => number
  readonly previewRowHeight: (rowIndex: number, newSize: number) => number
  readonly rowHeights: AxisSizeOverrides
  readonly setActiveResizeColumn: Dispatch<SetStateAction<number | null>>
  readonly setActiveResizeRow: Dispatch<SetStateAction<number | null>>
}

function clampColumnWidth(size: number): number {
  return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(size)))
}

function clampRowHeight(size: number): number {
  return Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, Math.round(size)))
}

export function useWorkbookAxisResizeState(input: {
  readonly controlledColumnWidths?: AxisSizeOverrides | undefined
  readonly controlledHiddenColumns?: Readonly<Record<number, true>> | undefined
  readonly controlledHiddenRows?: Readonly<Record<number, true>> | undefined
  readonly controlledRowHeights?: AxisSizeOverrides | undefined
  readonly onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined
  readonly onRowHeightChange?: ((rowIndex: number, newSize: number) => void) | undefined
  readonly sheetName: string
}): WorkbookAxisResizeState {
  const {
    controlledColumnWidths,
    controlledHiddenColumns,
    controlledHiddenRows,
    controlledRowHeights,
    onColumnWidthChange,
    onRowHeightChange,
    sheetName,
  } = input
  const [activeResizeColumn, setActiveResizeColumn] = useState<number | null>(null)
  const [activeResizeRow, setActiveResizeRow] = useState<number | null>(null)
  const [columnWidthsBySheet, setColumnWidthsBySheet] = useState<Record<string, Record<number, number>>>({})
  const [rowHeightsBySheet, setRowHeightsBySheet] = useState<Record<string, Record<number, number>>>({})
  const [columnResizePreview, setColumnResizePreview] = useState<ColumnResizePreview | null>(null)
  const [rowResizePreview, setRowResizePreview] = useState<RowResizePreview | null>(null)
  const columnResizePreviewRef = useRef<ColumnResizePreview | null>(null)
  const rowResizePreviewRef = useRef<RowResizePreview | null>(null)

  const baseColumnWidths = controlledColumnWidths ?? columnWidthsBySheet[sheetName] ?? EMPTY_COLUMN_WIDTHS
  const baseRowHeights = controlledRowHeights ?? rowHeightsBySheet[sheetName] ?? EMPTY_ROW_HEIGHTS
  const sizedColumnWidths = useMemo(() => {
    if (!columnResizePreview || columnResizePreview.sheetName !== sheetName) {
      return baseColumnWidths
    }
    if (baseColumnWidths[columnResizePreview.columnIndex] === columnResizePreview.width) {
      return baseColumnWidths
    }
    return {
      ...baseColumnWidths,
      [columnResizePreview.columnIndex]: columnResizePreview.width,
    }
  }, [baseColumnWidths, columnResizePreview, sheetName])
  const sizedRowHeights = useMemo(() => {
    if (!rowResizePreview || rowResizePreview.sheetName !== sheetName) {
      return baseRowHeights
    }
    if (baseRowHeights[rowResizePreview.rowIndex] === rowResizePreview.height) {
      return baseRowHeights
    }
    return {
      ...baseRowHeights,
      [rowResizePreview.rowIndex]: rowResizePreview.height,
    }
  }, [baseRowHeights, rowResizePreview, sheetName])
  const columnWidths = useMemo(
    () => applyHiddenAxisSizes(sizedColumnWidths, controlledHiddenColumns),
    [controlledHiddenColumns, sizedColumnWidths],
  )
  const rowHeights = useMemo(() => applyHiddenAxisSizes(sizedRowHeights, controlledHiddenRows), [controlledHiddenRows, sizedRowHeights])

  useEffect(() => {
    const preview = columnResizePreviewRef.current
    if (!preview || preview.sheetName !== sheetName) {
      return
    }
    if (baseColumnWidths[preview.columnIndex] !== preview.width) {
      return
    }
    columnResizePreviewRef.current = null
    setColumnResizePreview((current) =>
      current?.sheetName === preview.sheetName && current.columnIndex === preview.columnIndex && current.width === preview.width
        ? null
        : current,
    )
  }, [baseColumnWidths, sheetName])

  useEffect(() => {
    const preview = rowResizePreviewRef.current
    if (!preview || preview.sheetName !== sheetName) {
      return
    }
    if (baseRowHeights[preview.rowIndex] !== preview.height) {
      return
    }
    rowResizePreviewRef.current = null
    setRowResizePreview((current) =>
      current?.sheetName === preview.sheetName && current.rowIndex === preview.rowIndex && current.height === preview.height
        ? null
        : current,
    )
  }, [baseRowHeights, sheetName])

  const commitColumnWidth = useCallback(
    (columnIndex: number, newSize: number) => {
      const clampedSize = clampColumnWidth(newSize)
      if (onColumnWidthChange) {
        onColumnWidthChange(columnIndex, clampedSize)
        return
      }
      setColumnWidthsBySheet((current) => {
        const nextSheetWidths = current[sheetName] ?? EMPTY_COLUMN_WIDTHS
        if (nextSheetWidths[columnIndex] === clampedSize) {
          return current
        }
        return {
          ...current,
          [sheetName]: {
            ...nextSheetWidths,
            [columnIndex]: clampedSize,
          },
        }
      })
    },
    [onColumnWidthChange, sheetName],
  )

  const commitRowHeight = useCallback(
    (rowIndex: number, newSize: number) => {
      const clampedSize = clampRowHeight(newSize)
      if (onRowHeightChange) {
        onRowHeightChange(rowIndex, clampedSize)
        return
      }
      setRowHeightsBySheet((current) => {
        const nextSheetHeights = current[sheetName] ?? EMPTY_ROW_HEIGHTS
        if (nextSheetHeights[rowIndex] === clampedSize) {
          return current
        }
        return {
          ...current,
          [sheetName]: {
            ...nextSheetHeights,
            [rowIndex]: clampedSize,
          },
        }
      })
    },
    [onRowHeightChange, sheetName],
  )

  const previewColumnWidth = useCallback(
    (columnIndex: number, newSize: number): number => {
      const clampedSize = clampColumnWidth(newSize)
      const nextPreview = { sheetName, columnIndex, width: clampedSize }
      columnResizePreviewRef.current = nextPreview
      setColumnResizePreview((current) =>
        current?.sheetName === nextPreview.sheetName &&
        current.columnIndex === nextPreview.columnIndex &&
        current.width === nextPreview.width
          ? current
          : nextPreview,
      )
      return clampedSize
    },
    [sheetName],
  )

  const previewRowHeight = useCallback(
    (rowIndex: number, newSize: number): number => {
      const clampedSize = clampRowHeight(newSize)
      const nextPreview = { sheetName, rowIndex, height: clampedSize }
      rowResizePreviewRef.current = nextPreview
      setRowResizePreview((current) =>
        current?.sheetName === nextPreview.sheetName && current.rowIndex === nextPreview.rowIndex && current.height === nextPreview.height
          ? current
          : nextPreview,
      )
      return clampedSize
    },
    [sheetName],
  )

  const getPreviewColumnWidth = useCallback(
    (columnIndex: number): number | null => {
      const preview = columnResizePreviewRef.current
      return preview?.sheetName === sheetName && preview.columnIndex === columnIndex ? preview.width : null
    },
    [sheetName],
  )

  const getPreviewRowHeight = useCallback(
    (rowIndex: number): number | null => {
      const preview = rowResizePreviewRef.current
      return preview?.sheetName === sheetName && preview.rowIndex === rowIndex ? preview.height : null
    },
    [sheetName],
  )

  const clearColumnResizePreview = useCallback(
    (columnIndex: number) => {
      const preview = columnResizePreviewRef.current
      if (preview?.sheetName === sheetName && preview.columnIndex === columnIndex) {
        columnResizePreviewRef.current = null
      }
      setColumnResizePreview((current) => (current?.sheetName === sheetName && current.columnIndex === columnIndex ? null : current))
    },
    [sheetName],
  )

  const clearRowResizePreview = useCallback(
    (rowIndex: number) => {
      const preview = rowResizePreviewRef.current
      if (preview?.sheetName === sheetName && preview.rowIndex === rowIndex) {
        rowResizePreviewRef.current = null
      }
      setRowResizePreview((current) => (current?.sheetName === sheetName && current.rowIndex === rowIndex ? null : current))
    },
    [sheetName],
  )

  return {
    activeResizeColumn,
    activeResizeRow,
    clearColumnResizePreview,
    clearRowResizePreview,
    columnWidths,
    commitColumnWidth,
    commitRowHeight,
    getPreviewColumnWidth,
    getPreviewRowHeight,
    hasColumnResizePreview: columnResizePreview !== null,
    hasRowResizePreview: rowResizePreview !== null,
    previewColumnWidth,
    previewRowHeight,
    rowHeights,
    setActiveResizeColumn,
    setActiveResizeRow,
  }
}
