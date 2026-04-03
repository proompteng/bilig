import {
  ValueTag,
  formatCellDisplayValue,
  formatErrorCode,
  shouldRightAlignCell,
  type CellSnapshot,
  type CellStyleRecord,
} from "@bilig/protocol";
import { GridCellKind, type GridCell } from "@glideapps/glide-data-grid";
import type { GridEngineLike } from "./grid-engine.js";

const DEFAULT_FONT_FALLBACK = '"JetBrainsMono Nerd Font","JetBrains Mono",monospace';

export function cellToEditorSeed(
  snapshot: Pick<CellSnapshot, "formula" | "input" | "value">,
): string {
  if (snapshot.formula) {
    return `=${snapshot.formula}`;
  }
  if (snapshot.input === null || snapshot.input === undefined) {
    switch (snapshot.value.tag) {
      case ValueTag.Number:
        return String(snapshot.value.value);
      case ValueTag.Boolean:
        return snapshot.value.value ? "TRUE" : "FALSE";
      case ValueTag.Empty:
        return "";
      case ValueTag.String:
        return snapshot.value.value;
      case ValueTag.Error:
        return formatErrorCode(snapshot.value.code);
    }
  }
  if (typeof snapshot.input === "boolean") {
    return snapshot.input ? "TRUE" : "FALSE";
  }
  return String(snapshot.input);
}

export function snapshotToGridCell(
  snapshot: Pick<CellSnapshot, "formula" | "input" | "value" | "format">,
  style?: CellStyleRecord,
): GridCell {
  const rawValue = cellToEditorSeed(snapshot);
  const themeOverride = cellStyleToThemeOverride(style);
  const displayText = formatCellDisplayValue(snapshot.value, snapshot.format);
  const contentAlign = resolveContentAlign(snapshot, style);
  const allowWrapping = style?.alignment?.wrap === true;

  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return {
        kind: GridCellKind.Number,
        allowOverlay: false,
        data: snapshot.value.value,
        displayData: displayText,
        readonly: false,
        copyData: snapshot.formula ? rawValue : displayText,
        contentAlign,
        ...(allowWrapping ? { allowWrapping: true } : {}),
        ...(themeOverride ? { themeOverride } : {}),
      };
    case ValueTag.Boolean:
      return {
        kind: GridCellKind.Boolean,
        allowOverlay: false,
        data: snapshot.value.value,
        readonly: false,
        copyData: snapshot.formula ? rawValue : snapshot.value.value ? "TRUE" : "FALSE",
        ...(themeOverride ? { themeOverride } : {}),
      };
    case ValueTag.Error:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: formatErrorCode(snapshot.value.code),
        displayData: formatErrorCode(snapshot.value.code),
        readonly: false,
        copyData: snapshot.formula ? rawValue : formatErrorCode(snapshot.value.code),
        ...(allowWrapping ? { allowWrapping: true } : {}),
        ...(themeOverride ? { themeOverride } : {}),
      };
    case ValueTag.String:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: snapshot.value.value,
        displayData: displayText,
        readonly: false,
        copyData: snapshot.formula ? rawValue : displayText,
        contentAlign,
        ...(allowWrapping ? { allowWrapping: true } : {}),
        ...(themeOverride ? { themeOverride } : {}),
      };
    case ValueTag.Empty:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: "",
        displayData: displayText,
        readonly: false,
        copyData: snapshot.formula ? rawValue : displayText,
        contentAlign,
        ...(allowWrapping ? { allowWrapping: true } : {}),
        ...(themeOverride ? { themeOverride } : {}),
      };
  }
}

export function cellToGridCell(engine: GridEngineLike, sheetName: string, addr: string): GridCell {
  const snapshot = engine.getCell(sheetName, addr);
  return snapshotToGridCell(snapshot, engine.getCellStyle(snapshot.styleId));
}

export function cellStyleToThemeOverride(style: CellStyleRecord | undefined) {
  if (!style?.font) {
    return undefined;
  }
  const fontStyleParts: string[] = [];
  if (style.font?.italic) {
    fontStyleParts.push("italic");
  }
  fontStyleParts.push(style.font?.bold ? "700" : "400");
  fontStyleParts.push(`${style.font?.size ?? 13}px`);
  return {
    ...(style.font?.color ? { textDark: style.font.color } : {}),
    ...(fontStyleParts.length > 0 ? { baseFontStyle: fontStyleParts.join(" ") } : {}),
    fontFamily: getResolvedCellFontFamily(),
  };
}

function resolveContentAlign(
  snapshot: Pick<CellSnapshot, "value" | "format">,
  style?: CellStyleRecord,
): "left" | "center" | "right" {
  switch (style?.alignment?.horizontal) {
    case "left":
      return "left";
    case "center":
      return "center";
    case "right":
      return "right";
    case "general":
    case undefined:
      return shouldRightAlignCell(snapshot.value, snapshot.format) ? "right" : "left";
  }
}

export function getResolvedCellFontFamily(): string {
  return DEFAULT_FONT_FALLBACK;
}
