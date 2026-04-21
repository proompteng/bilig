import { indexToColumn } from '@bilig/formula'
import type { SpreadsheetEngine } from '@bilig/core'
import type { CellValue } from '@bilig/protocol'
import type { WorkPaperCellChange } from './work-paper-types.js'

export function materializeTrackedIndexChanges(
  engine: SpreadsheetEngine,
  changedCellIndices: readonly number[] | Uint32Array,
): readonly WorkPaperCellChange[] {
  if (changedCellIndices.length === 0) {
    return []
  }
  const workbook = engine.workbook
  const cellStore = workbook.cellStore
  const readValue = (cellIndex: number): CellValue => cellStore.getValue(cellIndex, (id) => engine.strings.get(id))
  const columnLabels: string[] = []
  const formatAddressCached = (row: number, col: number): string => {
    let label = columnLabels[col]
    if (label === undefined) {
      label = indexToColumn(col)
      columnLabels[col] = label
    }
    return `${label}${row + 1}`
  }
  let firstSheetId: number | undefined
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const sheetId = cellStore.sheetIds[changedCellIndices[index]!]
    if (sheetId !== undefined) {
      firstSheetId = sheetId
      break
    }
  }
  if (firstSheetId === undefined) {
    return []
  }
  let isSingleSheetChangeSet = true
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const sheetId = cellStore.sheetIds[changedCellIndices[index]!]
    if (sheetId !== undefined && sheetId !== firstSheetId) {
      isSingleSheetChangeSet = false
      break
    }
  }
  const changes: WorkPaperCellChange[] = []
  if (isSingleSheetChangeSet) {
    const sheet = workbook.getSheetById(firstSheetId)
    const sheetName = sheet?.name ?? workbook.getSheetNameById(firstSheetId)
    const isPhysicalSheet = !sheet || sheet.structureVersion === 1
    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const cellIndex = changedCellIndices[index]!
      if (cellStore.sheetIds[cellIndex] !== firstSheetId) {
        continue
      }
      let row: number
      let col: number
      if (isPhysicalSheet) {
        row = cellStore.rows[cellIndex]!
        col = cellStore.cols[cellIndex]!
      } else {
        const position = workbook.getCellPosition(cellIndex)
        /* v8 ignore next -- defensive guard for stale logical indices without visible positions. */
        if (!position) {
          continue
        }
        row = position.row
        col = position.col
      }
      changes.push({
        kind: 'cell',
        address: { sheet: firstSheetId, row, col },
        sheetName,
        a1: formatAddressCached(row, col),
        newValue: readValue(cellIndex),
      })
    }
    return changes
  }

  const sheetNames = new Map<number, string>()
  const physicalSheetIds = new Set<number>()
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    const sheetId = cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      continue
    }
    let sheetName = sheetNames.get(sheetId)
    if (sheetName === undefined) {
      const sheet = workbook.getSheetById(sheetId)
      sheetName = sheet?.name ?? workbook.getSheetNameById(sheetId)
      sheetNames.set(sheetId, sheetName)
      if (!sheet || sheet.structureVersion === 1) {
        physicalSheetIds.add(sheetId)
      }
    }
    let row: number
    let col: number
    if (physicalSheetIds.has(sheetId)) {
      row = cellStore.rows[cellIndex]!
      col = cellStore.cols[cellIndex]!
    } else {
      const position = workbook.getCellPosition(cellIndex)
      /* v8 ignore next -- defensive guard for stale logical indices without visible positions. */
      if (!position) {
        continue
      }
      row = position.row
      col = position.col
    }
    changes.push({
      kind: 'cell',
      address: { sheet: sheetId, row, col },
      sheetName,
      a1: formatAddressCached(row, col),
      newValue: readValue(cellIndex),
    })
  }
  return changes
}
