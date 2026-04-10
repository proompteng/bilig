import {
  buildCellNumberFormatCode,
  getCellNumberFormatKind,
  type CellBorderSideSnapshot,
  type CellHorizontalAlignment,
  type CellNumberFormatRecord,
  type CellStyleRecord,
  type CellVerticalAlignment,
} from "@bilig/protocol";
import type {
  WorkbookCellNumberFormatRecord,
  WorkbookCellStyleRecord,
} from "./workbook-metadata-types.js";

export function deleteRecordsBySheet<T>(
  bucket: Map<string, T>,
  sheetName: string,
  readSheetName: (record: T) => string,
): void {
  for (const [key, record] of bucket.entries()) {
    if (readSheetName(record) === sheetName) {
      bucket.delete(key);
    }
  }
}

export function normalizeCellStyleRecord(style: CellStyleRecord): WorkbookCellStyleRecord {
  const record: WorkbookCellStyleRecord = { id: style.id.trim() };
  if (record.id.length === 0) {
    throw new Error("Cell style id must be non-empty");
  }
  const fillColor = normalizeBackgroundColor(style.fill?.backgroundColor);
  if (fillColor) {
    record.fill = { backgroundColor: fillColor };
  }
  const fontFamily = normalizeFontFamily(style.font?.family);
  const fontColor = normalizeColor(style.font?.color);
  const fontSize =
    typeof style.font?.size === "number" && Number.isFinite(style.font.size)
      ? Math.max(1, Math.min(144, style.font.size))
      : undefined;
  if (
    fontFamily ||
    fontColor ||
    fontSize !== undefined ||
    style.font?.bold === true ||
    style.font?.italic === true ||
    style.font?.underline === true
  ) {
    record.font = {
      ...(fontFamily ? { family: fontFamily } : {}),
      ...(fontColor ? { color: fontColor } : {}),
      ...(fontSize !== undefined ? { size: fontSize } : {}),
      ...(style.font?.bold ? { bold: true } : {}),
      ...(style.font?.italic ? { italic: true } : {}),
      ...(style.font?.underline ? { underline: true } : {}),
    };
  }
  const horizontal = normalizeHorizontalAlignment(style.alignment?.horizontal);
  const vertical = normalizeVerticalAlignment(style.alignment?.vertical);
  const wrap = style.alignment?.wrap === true ? true : undefined;
  const indent =
    typeof style.alignment?.indent === "number" && Number.isFinite(style.alignment.indent)
      ? Math.max(0, Math.min(16, Math.trunc(style.alignment.indent)))
      : undefined;
  if (horizontal || vertical || wrap || indent !== undefined) {
    record.alignment = {
      ...(horizontal ? { horizontal } : {}),
      ...(vertical ? { vertical } : {}),
      ...(wrap ? { wrap: true } : {}),
      ...(indent !== undefined ? { indent } : {}),
    };
  }
  const borders = normalizeBorders(style.borders);
  if (borders) {
    record.borders = borders;
  }
  return record;
}

export function normalizeCellNumberFormatRecord(
  format: CellNumberFormatRecord,
): WorkbookCellNumberFormatRecord {
  const id = format.id.trim();
  if (id.length === 0) {
    throw new Error("Cell number format id must be non-empty");
  }
  const code = buildCellNumberFormatCode(format.code);
  return {
    id,
    code,
    kind: format.kind ?? getCellNumberFormatKind(code),
  };
}

export function cellStyleKey(style: CellStyleRecord): string {
  return JSON.stringify({
    fill: style.fill?.backgroundColor ?? null,
    font: style.font ?? null,
    alignment: style.alignment ?? null,
    borders: style.borders ?? null,
  });
}

export function cellStyleIdForKey(key: string): string {
  let hash = 2166136261;
  for (let index = 0; index < key.length; index += 1) {
    hash ^= key.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return `style-${(hash >>> 0).toString(16)}`;
}

export function cellNumberFormatIdForCode(code: string): string {
  let hash = 2166136261;
  for (let index = 0; index < code.length; index += 1) {
    hash ^= code.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return `format-${(hash >>> 0).toString(16)}`;
}

export function axisMetadataKey(sheetName: string, start: number, count: number): string {
  return `${sheetName}:${start}:${count}`;
}

function normalizeBackgroundColor(color: string | undefined): string | undefined {
  return normalizeColor(color);
}

function normalizeColor(color: string | undefined): string | undefined {
  if (!color) {
    return undefined;
  }
  const trimmed = color.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(trimmed)) {
    return trimmed;
  }
  if (/^#[0-9a-f]{3}$/.test(trimmed)) {
    return `#${trimmed[1]}${trimmed[1]}${trimmed[2]}${trimmed[2]}${trimmed[3]}${trimmed[3]}`;
  }
  throw new Error(`Unsupported background color: ${color}`);
}

function normalizeFontFamily(family: string | undefined): string | undefined {
  if (!family) {
    return undefined;
  }
  const trimmed = family.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}

function normalizeHorizontalAlignment(
  value: CellHorizontalAlignment | undefined,
): CellHorizontalAlignment | undefined {
  switch (value) {
    case undefined:
      return undefined;
    case "general":
    case "left":
    case "center":
    case "right":
      return value;
    default:
      return undefined;
  }
}

function normalizeVerticalAlignment(
  value: CellVerticalAlignment | undefined,
): CellVerticalAlignment | undefined {
  switch (value) {
    case undefined:
      return undefined;
    case "top":
    case "middle":
    case "bottom":
      return value;
    default:
      return undefined;
  }
}

function normalizeBorderStyle(
  value: CellBorderSideSnapshot["style"] | undefined,
): CellBorderSideSnapshot["style"] | undefined {
  switch (value) {
    case undefined:
      return undefined;
    case "solid":
    case "dashed":
    case "dotted":
    case "double":
      return value;
    default:
      return undefined;
  }
}

function normalizeBorderWeight(
  value: CellBorderSideSnapshot["weight"] | undefined,
): CellBorderSideSnapshot["weight"] | undefined {
  switch (value) {
    case undefined:
      return undefined;
    case "thin":
    case "medium":
    case "thick":
      return value;
    default:
      return undefined;
  }
}

function normalizeBorderSide(
  side: CellBorderSideSnapshot | undefined,
): CellBorderSideSnapshot | undefined {
  if (!side) {
    return undefined;
  }
  const color = normalizeColor(side.color);
  const style = normalizeBorderStyle(side.style);
  const weight = normalizeBorderWeight(side.weight);
  if (!color || !style || !weight) {
    return undefined;
  }
  return { color, style, weight };
}

function normalizeBorders(
  style: CellStyleRecord["borders"],
): CellStyleRecord["borders"] | undefined {
  if (!style) {
    return undefined;
  }
  const top = normalizeBorderSide(style.top);
  const right = normalizeBorderSide(style.right);
  const bottom = normalizeBorderSide(style.bottom);
  const left = normalizeBorderSide(style.left);
  if (!top && !right && !bottom && !left) {
    return undefined;
  }
  return {
    ...(top ? { top } : {}),
    ...(right ? { right } : {}),
    ...(bottom ? { bottom } : {}),
    ...(left ? { left } : {}),
  };
}
