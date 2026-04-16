import type { CellSnapshot, CellStyleRecord, WorkbookAxisEntrySnapshot } from '@bilig/protocol'

export interface WorkbookLocalBaseSheetRecord {
  readonly sheetId: number
  readonly name: string
  readonly sortOrder: number
  readonly freezeRows: number
  readonly freezeCols: number
}

export interface WorkbookLocalBaseCellInputRecord {
  readonly sheetId: number
  readonly sheetName: string
  readonly address: string
  readonly rowNum: number
  readonly colNum: number
  readonly input: CellSnapshot['input']
  readonly formula: CellSnapshot['formula']
  readonly format: CellSnapshot['format']
}

export interface WorkbookLocalBaseCellRenderRecord {
  readonly sheetId: number
  readonly sheetName: string
  readonly address: string
  readonly rowNum: number
  readonly colNum: number
  readonly value: CellSnapshot['value']
  readonly flags: number
  readonly version: number
  readonly styleId: CellSnapshot['styleId']
  readonly numberFormatId: CellSnapshot['numberFormatId']
}

export interface WorkbookLocalAuthoritativeBase {
  readonly sheets: readonly WorkbookLocalBaseSheetRecord[]
  readonly cellInputs: readonly WorkbookLocalBaseCellInputRecord[]
  readonly cellRenders: readonly WorkbookLocalBaseCellRenderRecord[]
  readonly rowAxisEntries: readonly {
    sheetId: number
    sheetName: string
    entry: WorkbookAxisEntrySnapshot
  }[]
  readonly columnAxisEntries: readonly {
    sheetId: number
    sheetName: string
    entry: WorkbookAxisEntrySnapshot
  }[]
  readonly styles: readonly CellStyleRecord[]
}

export interface WorkbookLocalProjectionOverlayCellRecord {
  readonly sheetId: number
  readonly sheetName: string
  readonly address: string
  readonly rowNum: number
  readonly colNum: number
  readonly value: CellSnapshot['value']
  readonly flags: number
  readonly version: number
  readonly input: CellSnapshot['input']
  readonly formula: CellSnapshot['formula']
  readonly format: CellSnapshot['format']
  readonly styleId: CellSnapshot['styleId']
  readonly numberFormatId: CellSnapshot['numberFormatId']
}

export interface WorkbookLocalProjectionOverlay {
  readonly cells: readonly WorkbookLocalProjectionOverlayCellRecord[]
  readonly rowAxisEntries: readonly {
    sheetId: number
    sheetName: string
    entry: WorkbookAxisEntrySnapshot
  }[]
  readonly columnAxisEntries: readonly {
    sheetId: number
    sheetName: string
    entry: WorkbookAxisEntrySnapshot
  }[]
  readonly styles: readonly CellStyleRecord[]
}

export interface WorkbookLocalAuthoritativeDelta {
  readonly replaceAll: boolean
  readonly replacedSheetIds: readonly number[]
  readonly base: WorkbookLocalAuthoritativeBase
}

export interface WorkbookLocalViewportCell {
  readonly row: number
  readonly col: number
  readonly snapshot: CellSnapshot
}

export interface WorkbookLocalViewportBase {
  readonly sheetId: number
  readonly sheetName: string
  readonly freezeRows: number
  readonly freezeCols: number
  readonly cells: readonly WorkbookLocalViewportCell[]
  readonly rowAxisEntries: readonly WorkbookAxisEntrySnapshot[]
  readonly columnAxisEntries: readonly WorkbookAxisEntrySnapshot[]
  readonly styles: readonly CellStyleRecord[]
}
