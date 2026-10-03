import type { SpreadsheetEngine } from '@bilig/core/headless-runtime'

export function tryRenameSheetMetadataOnlyPrevalidated(
  engine: SpreadsheetEngine,
  sheetId: number,
  oldName: string,
  newName: string,
  hasWorkbookRenameMetadata: boolean,
): boolean {
  return engine.renameSheetMetadataOnlyByIdTrustedPrevalidated(sheetId, oldName, newName, false, !hasWorkbookRenameMetadata)
}
