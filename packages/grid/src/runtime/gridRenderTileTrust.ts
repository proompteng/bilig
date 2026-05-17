import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { Item } from '../gridTypes.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'

export function hasCompleteRenderTileGrid(tile: GridRenderTile): boolean {
  const expectedBorderCount = expectedRenderTileGridBorderCount(tile)
  return expectedBorderCount === 0 || countRenderTileGridBorderRects(tile) >= expectedBorderCount
}

export function tileSelectedTextNeedsLocalRefresh(
  tile: GridRenderTile | null,
  selectedCell: Item | undefined,
  selectedCellSnapshot: CellSnapshot | null | undefined,
): boolean {
  if (!selectedCell) {
    return false
  }
  const selectedRun = findSelectedTextRun(tile, selectedCell)
  const expectedText = selectedSnapshotTextHint(selectedCellSnapshot)
  if (expectedText === undefined) {
    return false
  }
  if (expectedText === null) {
    return selectedRun !== null
  }
  return selectedRun?.text !== expectedText
}

function expectedRenderTileGridBorderCount(tile: GridRenderTile): number {
  const rowCount = tile.bounds.rowEnd - tile.bounds.rowStart + 1
  const colCount = tile.bounds.colEnd - tile.bounds.colStart + 1
  return rowCount > 0 && colCount > 0 ? rowCount + colCount : 0
}

function countRenderTileGridBorderRects(tile: GridRenderTile): number {
  const readableRectCount = Math.min(tile.rectCount, Math.floor(tile.rectInstances.length / GRID_RECT_INSTANCE_FLOAT_COUNT_V3))
  let count = 0
  for (let index = 0; index < readableRectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const width = tile.rectInstances[offset + 2] ?? 0
    const height = tile.rectInstances[offset + 3] ?? 0
    const borderAlpha = tile.rectInstances[offset + 11] ?? 0
    const borderThickness = tile.rectInstances[offset + 13] ?? 0
    if (borderAlpha > 0 && borderThickness > 0 && ((width <= 1.5 && height > 0) || (height <= 1.5 && width > 0))) {
      count += 1
    }
  }
  return count
}

function selectedSnapshotTextHint(snapshot: CellSnapshot | null | undefined): string | null | undefined {
  if (!snapshot) {
    return undefined
  }
  if (snapshot.input !== undefined && snapshot.input !== null && snapshot.input !== '') {
    return String(snapshot.input)
  }
  if (snapshot.formula !== undefined && snapshot.formula.length > 0) {
    return snapshot.formula
  }
  if (snapshot.value.tag === ValueTag.String) {
    return snapshot.value.value
  }
  if (snapshot.value.tag === ValueTag.Number) {
    return String(snapshot.value.value)
  }
  if (isDefaultPlaceholderEmptySnapshot(snapshot)) {
    return undefined
  }
  return null
}

function isDefaultPlaceholderEmptySnapshot(snapshot: CellSnapshot): boolean {
  return (
    snapshot.value.tag === ValueTag.Empty &&
    snapshot.version === 0 &&
    snapshot.flags === 0 &&
    snapshot.formula === undefined &&
    (snapshot.input === undefined || snapshot.input === '') &&
    snapshot.format === undefined &&
    snapshot.styleId === undefined &&
    snapshot.numberFormatId === undefined
  )
}

function findSelectedTextRun(tile: GridRenderTile | null, selectedCell: Item | undefined): { readonly text: string } | null {
  if (!tile || !selectedCell) {
    return null
  }
  return tile.textRuns.find((run) => run.col === selectedCell[0] && run.row === selectedCell[1] && run.text.length > 0) ?? null
}
