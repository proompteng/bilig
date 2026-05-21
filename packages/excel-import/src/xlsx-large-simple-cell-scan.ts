import type { WorkbookRichTextCellSnapshot } from '@bilig/protocol'
import type { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena } from './xlsx-large-simple-arena.js'

export interface ImportedWorksheetCellScan {
  readonly arena: ImportedWorkbookArena
  readonly sheetIndex: number
  readonly richTextCells: WorkbookRichTextCellSnapshot[]
  readonly styleIndexes: ImportedWorksheetStyleIndexArena
  readonly blankStyleCellCount: number
  readonly cellCount: number
  readonly valueCellCount: number
  readonly formulaCellCount: number
  readonly mergeCount?: number
  readonly conditionalFormatCount?: number
  readonly dataValidationCount?: number
  readonly tableCount?: number
  readonly rowCount: number
  readonly columnCount: number
  readonly usedRange: {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  } | null
}
