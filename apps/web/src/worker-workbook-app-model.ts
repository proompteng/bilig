import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  formatErrorCode,
  type CellRangeRef,
  type CellSnapshot,
  type LiteralInput,
  type WorkbookDefinedNameSnapshot,
} from "@bilig/protocol";

export type EditingMode = "idle" | "cell" | "formula";

export type ParsedEditorInput =
  | { kind: "clear" }
  | { kind: "formula"; formula: string }
  | { kind: "value"; value: LiteralInput };

export interface WorkbookEditorConflict {
  readonly sheetName: string;
  readonly address: string;
  readonly phase: "badge" | "compare";
  readonly baseSnapshot: CellSnapshot;
  readonly authoritativeSnapshot: CellSnapshot;
}

export type ZeroConnectionState =
  | { name: "connected" }
  | { name: "connecting"; reason?: string }
  | { name: "disconnected"; reason: string }
  | {
      name: "needs-auth";
      reason:
        | { type: "mutate"; status: 401 | 403; body?: string }
        | { type: "query"; status: 401 | 403; body?: string }
        | { type: "zero-cache"; reason: string };
    }
  | { name: "error"; reason: string }
  | { name: "closed"; reason: string };

export function createNextSheetName(sheetNames: readonly string[]): string {
  const existing = new Set(sheetNames);
  let index = 1;
  while (existing.has(`Sheet${index}`)) {
    index += 1;
  }
  return `Sheet${index}`;
}

export function normalizeSheetNameKey(value: string): string {
  return value.trim().toUpperCase();
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    throw new Error(message);
  }
}

export function toResolvedValue(cell: CellSnapshot): string {
  switch (cell.value.tag) {
    case ValueTag.Number:
      return String(cell.value.value);
    case ValueTag.Boolean:
      return cell.value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return cell.value.value;
    case ValueTag.Error:
      return formatErrorCode(cell.value.code);
    case ValueTag.Empty:
      return "";
  }
  const exhaustiveValue: never = cell.value;
  return String(exhaustiveValue);
}

export function toEditorValue(cell: CellSnapshot): string {
  if (cell.value.tag === ValueTag.Error) {
    return formatErrorCode(cell.value.code);
  }
  if (cell.formula) {
    return `=${cell.formula}`;
  }
  if (cell.input === null || cell.input === undefined) {
    return toResolvedValue(cell);
  }
  if (typeof cell.input === "boolean") {
    return cell.input ? "TRUE" : "FALSE";
  }
  return String(cell.input);
}

export function parseEditorInput(rawValue: string): ParsedEditorInput {
  const normalized = rawValue.trim();
  if (normalized.startsWith("=")) {
    return { kind: "formula", formula: normalized.slice(1) };
  }
  if (normalized === "") {
    return { kind: "clear" };
  }
  if (normalized === "TRUE" || normalized === "FALSE") {
    return { kind: "value", value: normalized === "TRUE" };
  }
  const numeric = Number(normalized);
  if (!Number.isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(normalized)) {
    return { kind: "value", value: numeric };
  }
  return { kind: "value", value: normalized };
}

export function parsedEditorInputFromSnapshot(snapshot: CellSnapshot): ParsedEditorInput {
  if (typeof snapshot.formula === "string") {
    return { kind: "formula", formula: snapshot.formula };
  }
  if (snapshot.input !== undefined && snapshot.input !== null) {
    return { kind: "value", value: snapshot.input };
  }
  switch (snapshot.value.tag) {
    case ValueTag.Empty:
      return { kind: "clear" };
    case ValueTag.Number:
      return { kind: "value", value: snapshot.value.value };
    case ValueTag.Boolean:
      return { kind: "value", value: snapshot.value.value };
    case ValueTag.String:
      return { kind: "value", value: snapshot.value.value };
    case ValueTag.Error:
      return { kind: "value", value: formatErrorCode(snapshot.value.code) };
  }
}

export function parsedEditorInputEquals(
  left: ParsedEditorInput,
  right: ParsedEditorInput,
): boolean {
  if (left.kind !== right.kind) {
    return false;
  }
  switch (left.kind) {
    case "clear":
      return true;
    case "formula":
      return right.kind === "formula" && left.formula === right.formula;
    case "value":
      return right.kind === "value" && left.value === right.value;
    default: {
      const exhaustiveLeft: never = left;
      return exhaustiveLeft;
    }
  }
}

export function parsedEditorInputMatchesSnapshot(
  parsed: ParsedEditorInput,
  snapshot: CellSnapshot,
): boolean {
  return parsedEditorInputEquals(parsed, parsedEditorInputFromSnapshot(snapshot));
}

export function sameCellContent(left: CellSnapshot, right: CellSnapshot): boolean {
  return parsedEditorInputEquals(
    parsedEditorInputFromSnapshot(left),
    parsedEditorInputFromSnapshot(right),
  );
}

export function clampSelectionMovement(
  address: string,
  sheetName: string,
  movement: readonly [-1 | 0 | 1, -1 | 0 | 1],
): string {
  const parsed = parseCellAddress(address, sheetName);
  const nextRow = Math.min(MAX_ROWS - 1, Math.max(0, parsed.row + movement[1]));
  const nextCol = Math.min(MAX_COLS - 1, Math.max(0, parsed.col + movement[0]));
  return formatAddress(nextRow, nextCol);
}

export function parseSelectionTarget(
  input: string,
  fallbackSheet: string,
  definedNames?: readonly WorkbookDefinedNameSnapshot[],
): { sheetName: string; address: string } | null {
  const trimmed = input.trim();
  if (trimmed.length === 0) {
    return null;
  }

  const matchingDefinedName = definedNames?.find(
    (entry) => entry.name.trim().toUpperCase() === trimmed.toUpperCase(),
  );
  if (
    matchingDefinedName?.value &&
    typeof matchingDefinedName.value === "object" &&
    "kind" in matchingDefinedName.value &&
    matchingDefinedName.value.kind === "cell-ref"
  ) {
    return {
      sheetName: matchingDefinedName.value.sheetName,
      address: matchingDefinedName.value.address.toUpperCase(),
    };
  }

  const bangIndex = trimmed.lastIndexOf("!");
  const nextSheetName = bangIndex === -1 ? fallbackSheet : trimmed.slice(0, bangIndex);
  const nextAddress = bangIndex === -1 ? trimmed : trimmed.slice(bangIndex + 1);

  try {
    const parsed = parseCellAddress(nextAddress.toUpperCase(), nextSheetName || fallbackSheet);
    return {
      sheetName: nextSheetName || fallbackSheet,
      address: formatAddress(parsed.row, parsed.col),
    };
  } catch {
    return null;
  }
}

export function parseSelectionRangeLabel(
  label: string,
  sheetName: string,
): { sheetName: string; startAddress: string; endAddress: string } {
  const trimmed = label.trim().toUpperCase();
  if (trimmed === "ALL") {
    return {
      sheetName,
      startAddress: "A1",
      endAddress: formatAddress(MAX_ROWS - 1, MAX_COLS - 1),
    };
  }

  const rowSelection = /^(\d+):(\d+)$/.exec(trimmed);
  if (rowSelection) {
    const startRow = Math.min(Number(rowSelection[1]) - 1, Number(rowSelection[2]) - 1);
    const endRow = Math.max(Number(rowSelection[1]) - 1, Number(rowSelection[2]) - 1);
    return {
      sheetName,
      startAddress: formatAddress(startRow, 0),
      endAddress: formatAddress(endRow, MAX_COLS - 1),
    };
  }

  const columnSelection = /^([A-Z]+):([A-Z]+)$/.exec(trimmed);
  if (columnSelection) {
    const startColumn = parseCellAddress(`${columnSelection[1]}1`, sheetName).col;
    const endColumn = parseCellAddress(`${columnSelection[2]}1`, sheetName).col;
    return {
      sheetName,
      startAddress: formatAddress(0, Math.min(startColumn, endColumn)),
      endAddress: formatAddress(MAX_ROWS - 1, Math.max(startColumn, endColumn)),
    };
  }

  const [startAddress = label, endAddress = startAddress] = trimmed.includes(":")
    ? trimmed.split(":")
    : [trimmed, trimmed];
  return { sheetName, startAddress, endAddress };
}

export function getNormalizedRangeBounds(range: CellRangeRef): {
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

export function createRangeRef(
  sheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

export function formatConnectionStateLabel(state: ZeroConnectionState["name"]): string {
  switch (state) {
    case "connected":
      return "Live";
    case "connecting":
      return "Connecting";
    case "disconnected":
      return "Disconnected";
    case "needs-auth":
      return "Needs auth";
    case "error":
      return "Error";
    case "closed":
      return "Closed";
    default:
      return state;
  }
}

export function canAttemptRemoteSync(state: ZeroConnectionState["name"]): boolean {
  return state === "connected";
}

export function toErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return JSON.stringify(error);
}

export function isMutationErrorResult(value: unknown): value is {
  type: "error";
  error: { type: "app" | "zero"; message: string; details?: unknown };
} {
  return (
    isRecord(value) &&
    value["type"] === "error" &&
    isRecord(value["error"]) &&
    (value["error"]["type"] === "app" || value["error"]["type"] === "zero") &&
    typeof value["error"]["message"] === "string"
  );
}

export function emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  };
}

export function isTextEntryTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    (target instanceof HTMLElement && target.isContentEditable)
  );
}
