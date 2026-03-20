import { formatAddress } from "@bilig/formula";
import type { Rectangle } from "@glideapps/glide-data-grid";
import { serializeClipboardMatrix, serializeClipboardPlainText } from "./gridClipboard.js";

export interface InternalClipboardRange {
  sourceStartAddress: string;
  sourceEndAddress: string;
  signature: string;
  plainText: string;
  rowCount: number;
  colCount: number;
}

export function buildInternalClipboardRange(
  range: Rectangle,
  values: readonly (readonly string[])[]
): InternalClipboardRange {
  return {
    sourceStartAddress: formatAddress(range.y, range.x),
    sourceEndAddress: formatAddress(range.y + range.height - 1, range.x + range.width - 1),
    signature: serializeClipboardMatrix(values),
    plainText: serializeClipboardPlainText(values),
    rowCount: range.height,
    colCount: range.width
  };
}

export function matchesInternalClipboardPaste(
  internalClipboard: InternalClipboardRange | null,
  values: readonly (readonly string[])[]
): boolean {
  if (!internalClipboard || values.length === 0 || values[0]?.length === 0) {
    return false;
  }
  return (
    internalClipboard.signature === serializeClipboardMatrix(values)
    && internalClipboard.rowCount === values.length
    && internalClipboard.colCount === (values[0]?.length ?? 0)
  );
}
