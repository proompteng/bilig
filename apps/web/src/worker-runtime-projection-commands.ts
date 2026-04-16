import { formatAddress, indexToColumn } from '@bilig/formula'
import type { CommitOp } from '@bilig/core'
import type { CellNumberFormatInput, CellRangeRef, CellSnapshot, CellStyleField, CellStylePatch } from '@bilig/protocol'

export function normalizeProjectedRowHeight(height: number | null): number | null {
  return height === null ? null : Math.max(1, Math.round(height))
}

export function normalizeProjectedColumnWidth(width: number | null, minColumnWidth: number, maxColumnWidth: number): number | null {
  return width === null ? null : Math.max(minColumnWidth, Math.min(maxColumnWidth, Math.round(width)))
}

export function autofitProjectedColumnWidth(args: {
  columnIndex: number
  charWidth: number
  padding: number
  sheet:
    | {
        grid: {
          forEachCellEntry(listener: (cellIndex: number, row: number, col: number) => void): void
        }
      }
    | undefined
  getCellDisplayValue: (row: number, col: number) => string
}): number {
  let widest = indexToColumn(args.columnIndex).length * args.charWidth

  args.sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
    if (col !== args.columnIndex) {
      return
    }
    const display = args.getCellDisplayValue(row, col)
    widest = Math.max(widest, display.length * args.charWidth)
  })

  return widest + args.padding
}

interface ProjectionCommandEngine {
  workbook: {
    getSheet(sheetName: string):
      | {
          grid: {
            forEachCellEntry(listener: (cellIndex: number, row: number, col: number) => void): void
          }
        }
      | null
      | undefined
  }
  setCellValue(sheetName: string, address: string, value: unknown): void
  setCellFormula(sheetName: string, address: string, formula: string): void
  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): void
  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void
  setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): void
  clearRangeNumberFormat(range: CellRangeRef): void
  clearRange(range: CellRangeRef): void
  clearCell(sheetName: string, address: string): void
  renderCommit(ops: CommitOp[]): void
  fillRange(source: CellRangeRef, target: CellRangeRef): void
  copyRange(source: CellRangeRef, target: CellRangeRef): void
  moveRange(source: CellRangeRef, target: CellRangeRef): void
  updateRowMetadata(sheetName: string, start: number, count: number, size: number | null, hidden: boolean | null): void
  updateColumnMetadata(sheetName: string, start: number, count: number, size: number | null, hidden: boolean | null): void
  setFreezePane(sheetName: string, rows: number, cols: number): void
  getCell(sheetName: string, address: string): Pick<CellSnapshot, 'value' | 'format'>
}

export class WorkerRuntimeProjectionCommands {
  constructor(
    private readonly options: {
      markProjectionDivergedFromLocalStore: () => void
      getProjectionEngine: () => Promise<ProjectionCommandEngine>
      getCell: (sheetName: string, address: string) => CellSnapshot
      minColumnWidth: number
      maxColumnWidth: number
      autofitCharWidth: number
      autofitPadding: number
      formatCellDisplayValue: (value: CellSnapshot['value'], format: string | undefined) => string
    },
  ) {}

  private async withProjectionMutation<T>(run: (engine: ProjectionCommandEngine) => T): Promise<T> {
    this.options.markProjectionDivergedFromLocalStore()
    return run(await this.options.getProjectionEngine())
  }

  async setCellValue(sheetName: string, address: string, value: CellSnapshot['input']): Promise<CellSnapshot> {
    await this.withProjectionMutation((engine) => {
      engine.setCellValue(sheetName, address, value ?? null)
    })
    return this.options.getCell(sheetName, address)
  }

  async setCellFormula(sheetName: string, address: string, formula: string): Promise<CellSnapshot> {
    await this.withProjectionMutation((engine) => {
      engine.setCellFormula(sheetName, address, formula)
    })
    return this.options.getCell(sheetName, address)
  }

  async setRangeStyle(range: CellRangeRef, patch: CellStylePatch): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.setRangeStyle(range, patch)
    })
  }

  async clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.clearRangeStyle(range, fields)
    })
  }

  async setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.setRangeNumberFormat(range, format)
    })
  }

  async clearRangeNumberFormat(range: CellRangeRef): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.clearRangeNumberFormat(range)
    })
  }

  async clearRange(range: CellRangeRef): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.clearRange(range)
    })
  }

  async clearCell(sheetName: string, address: string): Promise<CellSnapshot> {
    await this.withProjectionMutation((engine) => {
      engine.clearCell(sheetName, address)
    })
    return this.options.getCell(sheetName, address)
  }

  async renderCommit(ops: CommitOp[]): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.renderCommit(ops)
    })
  }

  async fillRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.fillRange(source, target)
    })
  }

  async copyRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.copyRange(source, target)
    })
  }

  async moveRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.moveRange(source, target)
    })
  }

  async updateRowMetadata(
    sheetName: string,
    startRow: number,
    count: number,
    height: number | null,
    hidden: boolean | null,
  ): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.updateRowMetadata(sheetName, startRow, count, normalizeProjectedRowHeight(height), hidden)
    })
  }

  async updateColumnMetadata(
    sheetName: string,
    startCol: number,
    count: number,
    width: number | null,
    hidden: boolean | null,
  ): Promise<number | null> {
    const normalizedWidth = normalizeProjectedColumnWidth(width, this.options.minColumnWidth, this.options.maxColumnWidth)
    await this.withProjectionMutation((engine) => {
      engine.updateColumnMetadata(sheetName, startCol, count, normalizedWidth, hidden)
    })
    return normalizedWidth
  }

  async updateColumnWidth(sheetName: string, columnIndex: number, width: number): Promise<number> {
    const normalizedWidth = await this.updateColumnMetadata(sheetName, columnIndex, 1, width, null)
    return normalizedWidth ?? width
  }

  async setFreezePane(sheetName: string, rows: number, cols: number): Promise<void> {
    await this.withProjectionMutation((engine) => {
      engine.setFreezePane(sheetName, Math.max(0, Math.round(rows)), Math.max(0, Math.round(cols)))
    })
  }

  async autofitColumn(sheetName: string, columnIndex: number): Promise<number> {
    const engine = await this.options.getProjectionEngine()
    const width = autofitProjectedColumnWidth({
      columnIndex,
      charWidth: this.options.autofitCharWidth,
      padding: this.options.autofitPadding,
      sheet: engine.workbook.getSheet(sheetName) ?? undefined,
      getCellDisplayValue: (row, col) => {
        const snapshot = engine.getCell(sheetName, formatAddress(row, col))
        return this.options.formatCellDisplayValue(snapshot.value, snapshot.format)
      },
    })
    return await this.updateColumnWidth(sheetName, columnIndex, width)
  }
}
