import {
  ValueTag,
  formatCellDisplayValue,
  formatErrorCode,
  shouldRightAlignCell,
  type CellSnapshot,
  type CellStyleRecord,
} from "@bilig/protocol";
import type { GridEngineLike } from "./grid-engine.js";

const DEFAULT_FONT_FALLBACK =
  '"Inter","SF Pro Text","SF Pro Display","Segoe UI","Helvetica Neue",Arial,sans-serif';
const DEFAULT_TEXT_COLOR = "#202124";
const TRANSPARENT_TEXT = "rgba(32, 33, 36, 0)";

interface GridCellOptions {
  readonly booleanSurfaceEnabled?: boolean;
  readonly textSurfaceEnabled?: boolean;
}

export const GridCellKind = {
  Number: "number",
  Boolean: "boolean",
  Text: "text",
} as const;

export type GridCellKind = (typeof GridCellKind)[keyof typeof GridCellKind];

export interface GridThemeOverride {
  textDark?: string;
  baseFontStyle?: string;
  fontFamily?: string;
}

export interface GridCell {
  readonly kind: GridCellKind;
  readonly allowOverlay: boolean;
  readonly data: number | boolean | string;
  readonly displayData?: string | undefined;
  readonly readonly: boolean;
  readonly copyData: string;
  readonly contentAlign?: "left" | "center" | "right" | undefined;
  readonly allowWrapping?: boolean | undefined;
  readonly themeOverride?: GridThemeOverride | undefined;
}

export interface RenderCellSnapshot {
  readonly kind: "number" | "boolean" | "error" | "string" | "empty";
  readonly displayText: string;
  readonly copyText: string;
  readonly align: "left" | "center" | "right";
  readonly wrap: boolean;
  readonly color: string;
  readonly font: string;
  readonly fontSize: number;
  readonly underline: boolean;
  readonly numberValue?: number | undefined;
  readonly booleanValue?: boolean | undefined;
  readonly stringValue?: string | undefined;
}

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
  options?: GridCellOptions,
): GridCell {
  const renderCell = snapshotToRenderCell(snapshot, style);

  switch (renderCell.kind) {
    case "number":
      return {
        kind: GridCellKind.Number,
        allowOverlay: false,
        data: renderCell.numberValue ?? 0,
        displayData: renderCell.displayText,
        readonly: false,
        copyData: renderCell.copyText,
        contentAlign: renderCell.align,
        ...(renderCell.wrap ? { allowWrapping: true } : {}),
        ...resolveTextThemeOverride(style, options),
      };
    case "boolean":
      if (options?.booleanSurfaceEnabled === true) {
        return {
          kind: GridCellKind.Text,
          allowOverlay: false,
          data: "",
          displayData: "",
          readonly: false,
          copyData: renderCell.copyText,
          contentAlign: "center",
        };
      }
      return {
        kind: GridCellKind.Boolean,
        allowOverlay: false,
        data: renderCell.booleanValue ?? false,
        readonly: false,
        copyData: renderCell.copyText,
      };
    case "error":
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: renderCell.displayText,
        displayData: renderCell.displayText,
        readonly: false,
        copyData: renderCell.copyText,
        ...(renderCell.wrap ? { allowWrapping: true } : {}),
        ...resolveTextThemeOverride(style, options),
      };
    case "string":
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: renderCell.stringValue ?? "",
        displayData: renderCell.displayText,
        readonly: false,
        copyData: renderCell.copyText,
        contentAlign: renderCell.align,
        ...(renderCell.wrap ? { allowWrapping: true } : {}),
        ...resolveTextThemeOverride(style, options),
      };
    case "empty":
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: "",
        displayData: renderCell.displayText,
        readonly: false,
        copyData: renderCell.copyText,
        contentAlign: renderCell.align,
        ...(renderCell.wrap ? { allowWrapping: true } : {}),
        ...resolveTextThemeOverride(style, options),
      };
  }
}

export function snapshotToRenderCell(
  snapshot: Pick<CellSnapshot, "formula" | "input" | "value" | "format">,
  style?: CellStyleRecord,
): RenderCellSnapshot {
  const copyText = cellToEditorSeed(snapshot);
  const displayText =
    snapshot.value.tag === ValueTag.Error
      ? formatErrorCode(snapshot.value.code)
      : snapshot.value.tag === ValueTag.Boolean
        ? snapshot.value.value
          ? "TRUE"
          : "FALSE"
        : formatCellDisplayValue(snapshot.value, snapshot.format);

  return {
    kind: resolveRenderCellKind(snapshot),
    displayText,
    copyText,
    align: resolveContentAlign(snapshot, style),
    wrap: style?.alignment?.wrap === true,
    color: style?.font?.color ?? DEFAULT_TEXT_COLOR,
    font: resolveCanvasFont(style),
    fontSize: style?.font?.size ?? 13,
    underline: style?.font?.underline === true,
    ...(snapshot.value.tag === ValueTag.Number ? { numberValue: snapshot.value.value } : {}),
    ...(snapshot.value.tag === ValueTag.Boolean ? { booleanValue: snapshot.value.value } : {}),
    ...(snapshot.value.tag === ValueTag.String ? { stringValue: snapshot.value.value } : {}),
  };
}

function resolveTextThemeOverride(
  style: CellStyleRecord | undefined,
  options?: GridCellOptions,
): { themeOverride?: GridThemeOverride } {
  const themeOverride = cellStyleToThemeOverride(style, options);
  return themeOverride ? { themeOverride } : {};
}

export function cellToGridCell(
  engine: GridEngineLike,
  sheetName: string,
  addr: string,
  options?: GridCellOptions,
): GridCell {
  const snapshot = engine.getCell(sheetName, addr);
  return snapshotToGridCell(snapshot, engine.getCellStyle(snapshot.styleId), options);
}

export function cellStyleToThemeOverride(
  style: CellStyleRecord | undefined,
  options?: GridCellOptions,
): GridThemeOverride | undefined {
  const textSurfaceEnabled = options?.textSurfaceEnabled === true;
  if (!style?.font) {
    return textSurfaceEnabled ? { textDark: TRANSPARENT_TEXT } : undefined;
  }
  const fontStyleParts: string[] = [];
  if (style.font?.italic) {
    fontStyleParts.push("italic");
  }
  fontStyleParts.push(style.font?.bold ? "700" : "400");
  fontStyleParts.push(`${style.font?.size ?? 13}px`);
  return {
    ...(textSurfaceEnabled || style.font?.color
      ? { textDark: textSurfaceEnabled ? TRANSPARENT_TEXT : style.font.color }
      : {}),
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

function resolveRenderCellKind(snapshot: Pick<CellSnapshot, "value">): RenderCellSnapshot["kind"] {
  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return "number";
    case ValueTag.Boolean:
      return "boolean";
    case ValueTag.Error:
      return "error";
    case ValueTag.String:
      return "string";
    case ValueTag.Empty:
      return "empty";
  }
}

function resolveCanvasFont(style: CellStyleRecord | undefined): string {
  const fontParts: string[] = [];
  if (style?.font?.italic) {
    fontParts.push("italic");
  }
  fontParts.push(style?.font?.bold ? "700" : "400");
  fontParts.push(`${style?.font?.size ?? 13}px`);
  fontParts.push(getResolvedCellFontFamily());
  return fontParts.join(" ");
}
