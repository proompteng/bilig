import type { SheetMetadataSnapshot, WorkbookConditionalFormatSnapshot } from '@bilig/protocol'

export function appendLargeSimpleConditionalFormats<T extends object>(
  input: T,
  conditionalFormats: readonly WorkbookConditionalFormatSnapshot[] | undefined,
): T | (T & { conditionalFormats: WorkbookConditionalFormatSnapshot[] }) {
  if (!conditionalFormats || conditionalFormats.length === 0) {
    return input
  }
  const existingConditionalFormats = (input as { readonly conditionalFormats?: readonly WorkbookConditionalFormatSnapshot[] })
    .conditionalFormats
  return {
    ...input,
    conditionalFormats: [...(existingConditionalFormats ?? []), ...conditionalFormats],
  }
}

export function normalizeLargeSimpleConditionalFormatIds(
  sheetName: string,
  conditionalFormats: readonly WorkbookConditionalFormatSnapshot[] | undefined,
): SheetMetadataSnapshot['conditionalFormats'] | undefined {
  if (!conditionalFormats || conditionalFormats.length === 0) {
    return undefined
  }
  return conditionalFormats.map((format, index) => ({
    ...format,
    id: `xlsx-cf:${sheetName}:${format.range.startAddress}:${format.range.endAddress}:${String(index + 1)}`,
  }))
}

export function readLargeSimpleConditionalFormattingBlockCount(worksheetXml: string): number {
  return [...worksheetXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b/gu)].length
}
