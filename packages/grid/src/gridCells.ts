import { selectors, type SpreadsheetEngine } from "@bilig/core";
import { ValueTag, formatErrorCode, type CellSnapshot } from "@bilig/protocol";
import { GridCellKind, type GridCell } from "@glideapps/glide-data-grid";

export function cellToEditorSeed(snapshot: Pick<CellSnapshot, "formula" | "input" | "value">): string {
  if (snapshot.formula) {
    return `=${snapshot.formula}`;
  }
  if (snapshot.input === null || snapshot.input === undefined) {
    switch (snapshot.value.tag) {
      case ValueTag.Number:
        return String(snapshot.value.value);
      case ValueTag.Boolean:
        return snapshot.value.value ? "TRUE" : "FALSE";
      case ValueTag.String:
        return snapshot.value.value;
      case ValueTag.Error:
        return formatErrorCode(snapshot.value.code);
      default:
        return "";
    }
  }
  if (typeof snapshot.input === "boolean") {
    return snapshot.input ? "TRUE" : "FALSE";
  }
  return String(snapshot.input);
}

export function snapshotToGridCell(snapshot: Pick<CellSnapshot, "formula" | "input" | "value">): GridCell {
  const rawValue = cellToEditorSeed(snapshot);

  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return {
        kind: GridCellKind.Number,
        allowOverlay: false,
        data: snapshot.value.value,
        displayData: String(snapshot.value.value),
        readonly: false,
        copyData: snapshot.formula ? rawValue : String(snapshot.value.value),
        contentAlign: "right"
      };
    case ValueTag.Boolean:
      return {
        kind: GridCellKind.Boolean,
        allowOverlay: false,
        data: snapshot.value.value,
        readonly: false,
        copyData: snapshot.formula ? rawValue : snapshot.value.value ? "TRUE" : "FALSE"
      };
    case ValueTag.Error:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: formatErrorCode(snapshot.value.code),
        displayData: formatErrorCode(snapshot.value.code),
        readonly: false,
        copyData: snapshot.formula ? rawValue : formatErrorCode(snapshot.value.code)
      };
    case ValueTag.String:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: snapshot.value.value,
        displayData: snapshot.value.value,
        readonly: false,
        copyData: snapshot.formula ? rawValue : snapshot.value.value
      };
    default:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: "",
        displayData: "",
        readonly: false,
        copyData: snapshot.formula ? rawValue : ""
      };
  }
}

export function cellToGridCell(engine: SpreadsheetEngine, sheetName: string, addr: string): GridCell {
  return snapshotToGridCell(selectors.selectCellSnapshot(engine, sheetName, addr));
}
