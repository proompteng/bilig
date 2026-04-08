import { isEngineReplicaSnapshot, type EngineReplicaSnapshot } from "@bilig/core";
import { parseCellAddress } from "@bilig/formula";
import {
  isWorkbookSnapshot,
  ValueTag,
  type CellBorderStyle,
  type CellBorderWeight,
  type CellHorizontalAlignment,
  type CellRangeRef,
  type CellStyleRecord,
  type CellValue,
  type CellVerticalAlignment,
  type WorkbookSnapshot,
} from "@bilig/protocol";
import type { DirtyRegion, WorkbookEventPayload } from "@bilig/zero-sync";
import type {
  AxisMetadataSourceRow,
  CellEvalRow,
  CellSourceRow,
  DefinedNameSourceRow,
  NumberFormatSourceRow,
  SheetSourceRow,
  StyleSourceRow,
  WorkbookMetadataSourceRow,
} from "./projection.js";

export type FocusedCellEventPayload = Extract<
  WorkbookEventPayload,
  { kind: "setCellValue" | "setCellFormula" | "clearCell" }
>;

export type StyleRangeEventPayload = Extract<
  WorkbookEventPayload,
  { kind: "setRangeStyle" | "clearRangeStyle" }
>;

export type NumberFormatRangeEventPayload = Extract<
  WorkbookEventPayload,
  { kind: "setRangeNumberFormat" | "clearRangeNumberFormat" }
>;

export type ColumnMetadataEventPayload = Extract<
  WorkbookEventPayload,
  { kind: "updateColumnWidth" }
>;

export function createEmptyWorkbookSnapshot(documentId: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: documentId,
    },
    sheets: [
      {
        id: 1,
        name: "Sheet1",
        order: 0,
        cells: [],
      },
    ],
  };
}

export function nowIso(): string {
  return new Date().toISOString();
}

export function parseInteger(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  if (typeof value === "string" && value.length > 0) {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return 0;
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function isDirtyRegion(value: unknown): value is DirtyRegion {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["rowStart"] === "number" &&
    typeof value["rowEnd"] === "number" &&
    typeof value["colStart"] === "number" &&
    typeof value["colEnd"] === "number"
  );
}

function isCellHorizontalAlignment(value: unknown): value is CellHorizontalAlignment {
  return value === "general" || value === "left" || value === "center" || value === "right";
}

function isCellVerticalAlignment(value: unknown): value is CellVerticalAlignment {
  return value === "top" || value === "middle" || value === "bottom";
}

function isCellBorderStyle(value: unknown): value is CellBorderStyle {
  return value === "solid" || value === "dashed" || value === "dotted" || value === "double";
}

function isCellBorderWeight(value: unknown): value is CellBorderWeight {
  return value === "thin" || value === "medium" || value === "thick";
}

export function parseCheckpointPayload(value: unknown, documentId: string): WorkbookSnapshot {
  return isWorkbookSnapshot(value) ? value : createEmptyWorkbookSnapshot(documentId);
}

export function parseCheckpointReplicaState(value: unknown): EngineReplicaSnapshot | null {
  return isEngineReplicaSnapshot(value) ? value : null;
}

export function isFocusedCellEventPayload(
  payload: WorkbookEventPayload,
): payload is FocusedCellEventPayload {
  return (
    payload.kind === "setCellValue" ||
    payload.kind === "setCellFormula" ||
    payload.kind === "clearCell"
  );
}

export function isStyleRangeEventPayload(
  payload: WorkbookEventPayload,
): payload is StyleRangeEventPayload {
  return payload.kind === "setRangeStyle" || payload.kind === "clearRangeStyle";
}

export function isNumberFormatRangeEventPayload(
  payload: WorkbookEventPayload,
): payload is NumberFormatRangeEventPayload {
  return payload.kind === "setRangeNumberFormat" || payload.kind === "clearRangeNumberFormat";
}

export function isColumnMetadataEventPayload(
  payload: WorkbookEventPayload,
): payload is ColumnMetadataEventPayload {
  return payload.kind === "updateColumnWidth";
}

export function eventRequiresRecalc(payload: WorkbookEventPayload): boolean {
  return !(
    payload.kind === "setRangeStyle" ||
    payload.kind === "clearRangeStyle" ||
    payload.kind === "setRangeNumberFormat" ||
    payload.kind === "clearRangeNumberFormat" ||
    payload.kind === "updateColumnWidth"
  );
}

function semanticSignature(value: unknown): string {
  return JSON.stringify(value);
}

export function sheetSignature(row: SheetSourceRow): string {
  return semanticSignature([row.name, row.sortOrder, row.freezeRows, row.freezeCols]);
}

export function cellSignature(row: CellSourceRow): string {
  return semanticSignature([
    row.sheetName,
    row.address,
    row.rowNum,
    row.colNum,
    row.inputValue ?? null,
    row.formula ?? null,
    row.format ?? null,
    row.styleId ?? null,
    row.explicitFormatId ?? null,
  ]);
}

export function axisSignature(row: AxisMetadataSourceRow): string {
  return semanticSignature([
    row.sheetName,
    row.startIndex,
    row.count,
    row.size ?? null,
    row.hidden ?? null,
  ]);
}

export function definedNameSignature(row: DefinedNameSourceRow): string {
  return semanticSignature([row.name, row.value]);
}

export function workbookMetadataSignature(row: WorkbookMetadataSourceRow): string {
  return semanticSignature([row.key, row.value]);
}

export function styleSignature(row: StyleSourceRow): string {
  return semanticSignature([row.id, row.recordJSON, row.hash]);
}

export function numberFormatSignature(row: NumberFormatSourceRow): string {
  return semanticSignature([row.id, row.code, row.kind]);
}

export function cellEvalSignature(row: CellEvalRow): string {
  return semanticSignature([
    row.sheetName,
    row.address,
    row.rowNum,
    row.colNum,
    row.value,
    row.flags,
    row.version,
    row.styleId,
    row.styleJson,
    row.formatId,
    row.formatCode,
  ]);
}

export function parseJsonKey(key: string): unknown[] {
  const parsed = JSON.parse(key) as unknown;
  if (!Array.isArray(parsed)) {
    throw new Error(`Invalid projection key: ${key}`);
  }
  return parsed;
}

function isCellValue(value: unknown): value is CellValue {
  if (!isRecord(value) || typeof value["tag"] !== "number") {
    return false;
  }
  const tag = value["tag"];
  if (tag === 0) {
    return true;
  }
  if (tag === 1) {
    return typeof value["value"] === "number";
  }
  if (tag === 2) {
    return typeof value["value"] === "boolean";
  }
  if (tag === 3) {
    return typeof value["value"] === "string";
  }
  if (tag === 4) {
    return typeof value["code"] === "number";
  }
  return false;
}

export function parseCellEvalValue(value: unknown): CellValue {
  return isCellValue(value) ? value : { tag: ValueTag.Empty };
}

export function parseCellStyleRecord(value: unknown): CellStyleRecord | null {
  if (!isRecord(value) || typeof value["id"] !== "string") {
    return null;
  }
  const record: CellStyleRecord = { id: value["id"] };
  if (isRecord(value["fill"])) {
    const fill = value["fill"];
    if (typeof fill["backgroundColor"] === "string") {
      record.fill = { backgroundColor: fill["backgroundColor"] };
    }
  }
  if (isRecord(value["font"])) {
    const font = value["font"];
    record.font = {
      ...(typeof font["family"] === "string" ? { family: font["family"] } : {}),
      ...(typeof font["size"] === "number" ? { size: font["size"] } : {}),
      ...(typeof font["bold"] === "boolean" ? { bold: font["bold"] } : {}),
      ...(typeof font["italic"] === "boolean" ? { italic: font["italic"] } : {}),
      ...(typeof font["underline"] === "boolean" ? { underline: font["underline"] } : {}),
      ...(typeof font["color"] === "string" ? { color: font["color"] } : {}),
    };
    if (Object.keys(record.font).length === 0) {
      delete record.font;
    }
  }
  if (isRecord(value["alignment"])) {
    const alignment = value["alignment"];
    const nextAlignment = {
      ...(isCellHorizontalAlignment(alignment["horizontal"])
        ? { horizontal: alignment["horizontal"] }
        : {}),
      ...(isCellVerticalAlignment(alignment["vertical"])
        ? { vertical: alignment["vertical"] }
        : {}),
      ...(typeof alignment["wrap"] === "boolean" ? { wrap: alignment["wrap"] } : {}),
      ...(typeof alignment["indent"] === "number" ? { indent: alignment["indent"] } : {}),
    };
    if (Object.keys(nextAlignment).length > 0) {
      record.alignment = nextAlignment;
    }
  }
  if (isRecord(value["borders"])) {
    const borders = value["borders"];
    const nextBorders: NonNullable<CellStyleRecord["borders"]> = {};
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const border = borders[side];
      if (!isRecord(border)) {
        continue;
      }
      if (
        isCellBorderStyle(border["style"]) &&
        isCellBorderWeight(border["weight"]) &&
        typeof border["color"] === "string"
      ) {
        nextBorders[side] = {
          style: border["style"],
          weight: border["weight"],
          color: border["color"],
        };
      }
    }
    if (Object.keys(nextBorders).length > 0) {
      record.borders = nextBorders;
    }
  }
  return record;
}

export function normalizeRangeBounds(range: CellRangeRef): {
  sheetName: string;
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
} {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    rowStart: Math.min(start.row, end.row),
    rowEnd: Math.max(start.row, end.row),
    colStart: Math.min(start.col, end.col),
    colEnd: Math.max(start.col, end.col),
  };
}

export function cellEvalRowInRange(
  row: Pick<CellEvalRow, "sheetName" | "rowNum" | "colNum">,
  range: CellRangeRef,
): boolean {
  const bounds = normalizeRangeBounds(range);
  return (
    row.sheetName === bounds.sheetName &&
    row.rowNum >= bounds.rowStart &&
    row.rowNum <= bounds.rowEnd &&
    row.colNum >= bounds.colStart &&
    row.colNum <= bounds.colEnd
  );
}

export function cellSourceRowInRange(
  row: Pick<CellSourceRow, "sheetName" | "rowNum" | "colNum">,
  range: CellRangeRef,
): boolean {
  const bounds = normalizeRangeBounds(range);
  return (
    row.sheetName === bounds.sheetName &&
    row.rowNum >= bounds.rowStart &&
    row.rowNum <= bounds.rowEnd &&
    row.colNum >= bounds.colStart &&
    row.colNum <= bounds.colEnd
  );
}
