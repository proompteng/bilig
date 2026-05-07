import { createHash, type Hash } from 'node:crypto'

import { formatAddress, parseCellAddress } from '../packages/formula/src/index.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import { projectSupportedSnapshotSemantics } from './import-export-fidelity-projection.ts'

export function roundTripSemanticsDigest(snapshot: WorkbookSnapshot): string {
  const hash = createHash('sha256')
  const populatedCellRowsBySheet = new Map<string, Map<number, PopulatedCellPosition[]>>()
  for (const sheet of snapshot.sheets) {
    for (const cell of sheet.cells) {
      if (cell.value !== undefined || cell.formula !== undefined) {
        const parsed = parseCellAddress(cell.address, sheet.name)
        const rows = populatedCellRowsBySheet.get(sheet.name) ?? new Map<number, PopulatedCellPosition[]>()
        const rowCells = rows.get(parsed.row) ?? []
        rowCells.push({ row: parsed.row, col: parsed.col, address: cell.address })
        rows.set(parsed.row, rowCells)
        populatedCellRowsBySheet.set(sheet.name, rows)
      }
    }
  }
  const metadataOnlyProjection = projectSupportedSnapshotSemantics({
    ...snapshot,
    sheets: snapshot.sheets.map((sheet) => ({
      ...sheet,
      cells: [],
    })),
  })
  const { styleRanges, ...metadataProjectionWithoutStyleRanges } = metadataOnlyProjection
  updateDigestJson(hash, {
    ...metadataProjectionWithoutStyleRanges,
    valueFormulaFormatSheets: metadataOnlyProjection.valueFormulaFormatSheets.map((sheet) => ({ ...sheet, cells: [] })),
    populatedCellStyles: populatedStyleCells(styleRanges, populatedCellRowsBySheet),
    dimensionSheets: metadataOnlyProjection.dimensionSheets.map((sheet) => ({
      name: sheet.name,
      columns: sheet.columns.filter((column) => column.size > 0).map((column) => ({ index: column.index, size: 0 })),
      rows: sheet.rows.filter((row) => row.size > 0).map((row) => ({ index: row.index, size: 0 })),
      merges: sheet.merges,
    })),
  })
  for (const sheet of snapshot.sheets.toSorted((left, right) => left.order - right.order)) {
    updateDigestJson(hash, ['sheet-cells', sheet.name, sheet.order])
    for (const cell of sheet.cells
      .filter((entry) => entry.value !== undefined || entry.formula !== undefined)
      .toSorted((left, right) => left.address.localeCompare(right.address))) {
      updateDigestJson(hash, {
        address: cell.address,
        ...(cell.value !== undefined ? { value: cell.value } : {}),
        ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
        ...(cell.format !== undefined ? { format: cell.format } : {}),
      })
    }
  }
  return hash.digest('hex')
}

interface PopulatedCellPosition {
  readonly row: number
  readonly col: number
  readonly address: string
}

type ProjectedStyleRange = ReturnType<typeof projectSupportedSnapshotSemantics>['styleRanges'][number]

function populatedStyleCells(
  styleRanges: readonly ProjectedStyleRange[],
  populatedCellRowsBySheet: ReadonlyMap<string, ReadonlyMap<number, readonly PopulatedCellPosition[]>>,
): { readonly sheetName: string; readonly address: string; readonly style: ProjectedStyleRange['style'] }[] {
  const cellsByKey = new Map<
    string,
    { readonly sheetName: string; readonly address: string; readonly style: ProjectedStyleRange['style'] }
  >()
  for (const styleRange of styleRanges) {
    const sheetRows = populatedCellRowsBySheet.get(styleRange.range.sheetName)
    if (!sheetRows) {
      continue
    }
    const start = parseCellAddress(styleRange.range.startAddress, styleRange.range.sheetName)
    const end = parseCellAddress(styleRange.range.endAddress, styleRange.range.sheetName)
    const startRow = Math.min(start.row, end.row)
    const endRow = Math.max(start.row, end.row)
    const startCol = Math.min(start.col, end.col)
    const endCol = Math.max(start.col, end.col)
    for (const [row, rowCells] of sheetRows) {
      if (row < startRow || row > endRow) {
        continue
      }
      for (const cell of rowCells) {
        if (cell.col < startCol || cell.col > endCol) {
          continue
        }
        const address = formatAddress(cell.row, cell.col)
        cellsByKey.set(`${styleRange.range.sheetName}:${address}`, {
          sheetName: styleRange.range.sheetName,
          address,
          style: styleRange.style,
        })
      }
    }
  }
  return [...cellsByKey.values()].toSorted((left, right) =>
    `${left.sheetName}:${left.address}`.localeCompare(`${right.sheetName}:${right.address}`),
  )
}

function updateDigestJson(hash: Hash, value: unknown): void {
  hash.update(JSON.stringify(value))
  hash.update('\n')
}
