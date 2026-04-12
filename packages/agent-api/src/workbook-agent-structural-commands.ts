import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type {
  CellRangeRef,
  WorkbookAxisEntrySnapshot,
  WorkbookSortSnapshot,
} from "@bilig/protocol";
import type { WorkbookAgentCommand, WorkbookAgentPreviewRange } from "./workbook-agent-bundles.js";

export type WorkbookAgentStructuralCommand = Exclude<
  WorkbookAgentCommand,
  | Extract<WorkbookAgentCommand, { kind: "writeRange" }>
  | Extract<WorkbookAgentCommand, { kind: "setRangeFormulas" }>
  | Extract<WorkbookAgentCommand, { kind: "clearRange" }>
  | Extract<WorkbookAgentCommand, { kind: "formatRange" }>
  | Extract<WorkbookAgentCommand, { kind: "fillRange" }>
  | Extract<WorkbookAgentCommand, { kind: "copyRange" }>
  | Extract<WorkbookAgentCommand, { kind: "moveRange" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertDefinedName" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteDefinedName" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertTable" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteTable" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertPivotTable" }>
  | Extract<WorkbookAgentCommand, { kind: "deletePivotTable" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertChart" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteChart" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertImage" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteImage" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertShape" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteShape" }>
  | Extract<WorkbookAgentCommand, { kind: "setDataValidation" }>
  | Extract<WorkbookAgentCommand, { kind: "clearDataValidation" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertConditionalFormat" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteConditionalFormat" }>
  | Extract<WorkbookAgentCommand, { kind: "setSheetProtection" }>
  | Extract<WorkbookAgentCommand, { kind: "clearSheetProtection" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertRangeProtection" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteRangeProtection" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertCommentThread" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteCommentThread" }>
  | Extract<WorkbookAgentCommand, { kind: "upsertNote" }>
  | Extract<WorkbookAgentCommand, { kind: "deleteNote" }>
>;

const HIGH_RISK_STRUCTURAL_COMMAND_KINDS = new Set<WorkbookAgentStructuralCommand["kind"]>([
  "createSheet",
  "renameSheet",
  "deleteSheet",
  "insertRows",
  "deleteRows",
  "insertColumns",
  "deleteColumns",
]);

const WORKBOOK_SCOPE_STRUCTURAL_COMMAND_KINDS = new Set<WorkbookAgentStructuralCommand["kind"]>([
  "createSheet",
  "renameSheet",
  "deleteSheet",
  "insertRows",
  "deleteRows",
  "insertColumns",
  "deleteColumns",
]);

function hasOwnProperty(value: object, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(value, key);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["startAddress"] === "string" &&
    typeof value["endAddress"] === "string"
  );
}

function isSortKeys(value: unknown): value is WorkbookSortSnapshot["keys"] {
  return (
    Array.isArray(value) &&
    value.length > 0 &&
    value.every(
      (entry) =>
        isRecord(entry) &&
        typeof entry["keyAddress"] === "string" &&
        entry["keyAddress"].trim().length > 0 &&
        (entry["direction"] === "asc" || entry["direction"] === "desc"),
    )
  );
}

function formatColumnLabel(index: number): string {
  return formatAddress(0, index).replace(/\d+/gu, "");
}

function formatRowSpanLabel(startRow: number, count: number): string {
  const first = startRow + 1;
  const last = startRow + count;
  return count === 1 ? `row ${String(first)}` : `rows ${String(first)}-${String(last)}`;
}

function formatColumnSpanLabel(startCol: number, count: number): string {
  const first = formatColumnLabel(startCol);
  const last = formatColumnLabel(startCol + count - 1);
  return count === 1 ? `column ${first}` : `columns ${first}-${last}`;
}

function rangeLabel(
  range: Pick<WorkbookAgentPreviewRange, "sheetName" | "startAddress" | "endAddress">,
) {
  return `${range.sheetName}!${range.startAddress}:${range.endAddress}`;
}

function normalizeRangeBounds(range: CellRangeRef): CellRangeRef {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

function countRangeCells(range: CellRangeRef): number {
  const normalized = normalizeRangeBounds(range);
  const start = parseCellAddress(normalized.startAddress, normalized.sheetName);
  const end = parseCellAddress(normalized.endAddress, normalized.sheetName);
  return (end.row - start.row + 1) * (end.col - start.col + 1);
}

function describeAxisMetadataCommand(input: {
  axis: "row" | "column";
  sheetName: string;
  spanLabel: string;
  size?: number | null;
  hidden?: boolean | null;
}): string {
  const sizeKey = input.axis === "row" ? "height" : "width";
  const hasSize = hasOwnProperty(input, "size");
  const hasHidden = hasOwnProperty(input, "hidden");

  if (hasSize && !hasHidden) {
    return input.size === null
      ? `Reset ${input.axis} ${sizeKey} for ${input.spanLabel} in ${input.sheetName}`
      : `Resize ${input.spanLabel} in ${input.sheetName}`;
  }
  if (!hasSize && hasHidden) {
    if (input.hidden === true) {
      return `Hide ${input.spanLabel} in ${input.sheetName}`;
    }
    if (input.hidden === false) {
      return `Unhide ${input.spanLabel} in ${input.sheetName}`;
    }
    return `Reset ${input.axis} visibility metadata for ${input.spanLabel} in ${input.sheetName}`;
  }
  if (hasSize && hasHidden && input.size !== null) {
    if (input.hidden === true) {
      return `Resize and hide ${input.spanLabel} in ${input.sheetName}`;
    }
    if (input.hidden === false) {
      return `Resize and unhide ${input.spanLabel} in ${input.sheetName}`;
    }
  }
  return `Update ${input.axis} metadata for ${input.spanLabel} in ${input.sheetName}`;
}

function formatFreezePaneDescription(rows: number, cols: number): string {
  const rowLabel = rows === 1 ? "1 row" : `${String(rows)} rows`;
  const colLabel = cols === 1 ? "1 column" : `${String(cols)} columns`;
  if (rows > 0 && cols > 0) {
    return `${rowLabel} and ${colLabel}`;
  }
  if (rows > 0) {
    return rowLabel;
  }
  return colLabel;
}

function describeSortKeys(keys: WorkbookSortSnapshot["keys"]): string {
  return keys.map((key) => `${key.keyAddress} ${key.direction}`).join(", ");
}

function getConsistentAxisEntrySize(input: {
  entries: readonly WorkbookAxisEntrySnapshot[];
  start: number;
  count: number;
  spanLabel: string;
  propertyLabel: string;
}): number | null {
  const entryByIndex = new Map(input.entries.map((entry) => [entry.index, entry]));
  let resolved: number | null | undefined;
  for (let index = input.start; index < input.start + input.count; index += 1) {
    const next = entryByIndex.get(index)?.size;
    const normalized = typeof next === "number" ? next : null;
    if (resolved === undefined) {
      resolved = normalized;
      continue;
    }
    if (resolved !== normalized) {
      throw new Error(
        `Cannot preserve ${input.propertyLabel} for ${input.spanLabel} because the existing ${input.propertyLabel} state is mixed. Specify ${input.propertyLabel} explicitly.`,
      );
    }
  }
  return resolved ?? null;
}

function getConsistentAxisEntryHidden(input: {
  entries: readonly WorkbookAxisEntrySnapshot[];
  start: number;
  count: number;
  spanLabel: string;
  propertyLabel: string;
}): boolean | null {
  const entryByIndex = new Map(input.entries.map((entry) => [entry.index, entry]));
  let resolved: boolean | null | undefined;
  for (let index = input.start; index < input.start + input.count; index += 1) {
    const next = entryByIndex.get(index)?.hidden;
    const normalized = typeof next === "boolean" ? next : null;
    if (resolved === undefined) {
      resolved = normalized;
      continue;
    }
    if (resolved !== normalized) {
      throw new Error(
        `Cannot preserve ${input.propertyLabel} for ${input.spanLabel} because the existing ${input.propertyLabel} state is mixed. Specify ${input.propertyLabel} explicitly.`,
      );
    }
  }
  return resolved ?? null;
}

export function isWorkbookAgentStructuralCommandKind(
  kind: string,
): kind is WorkbookAgentStructuralCommand["kind"] {
  switch (kind) {
    case "createSheet":
    case "renameSheet":
    case "deleteSheet":
    case "insertRows":
    case "deleteRows":
    case "insertColumns":
    case "deleteColumns":
    case "setFreezePane":
    case "setFilter":
    case "clearFilter":
    case "setSort":
    case "clearSort":
    case "updateRowMetadata":
    case "updateColumnMetadata":
      return true;
    default:
      return false;
  }
}

export function isWorkbookAgentStructuralCommand(
  command: WorkbookAgentCommand,
): command is WorkbookAgentStructuralCommand {
  return isWorkbookAgentStructuralCommandKind(command.kind);
}

export function isWorkbookAgentStructuralCommandValue(
  value: unknown,
): value is WorkbookAgentStructuralCommand {
  if (!isRecord(value)) {
    return false;
  }
  const kind = value["kind"];
  if (typeof kind !== "string" || !isWorkbookAgentStructuralCommandKind(kind)) {
    return false;
  }
  switch (kind) {
    case "createSheet":
    case "deleteSheet":
      return typeof value["name"] === "string" && value["name"].trim().length > 0;
    case "renameSheet":
      return (
        typeof value["currentName"] === "string" &&
        value["currentName"].trim().length > 0 &&
        typeof value["nextName"] === "string" &&
        value["nextName"].trim().length > 0
      );
    case "insertRows":
    case "deleteRows":
    case "insertColumns":
    case "deleteColumns":
      return (
        typeof value["sheetName"] === "string" &&
        Number.isInteger(value["start"]) &&
        Number(value["start"]) >= 0 &&
        Number.isInteger(value["count"]) &&
        Number(value["count"]) > 0
      );
    case "setFreezePane":
      return (
        typeof value["sheetName"] === "string" &&
        Number.isInteger(value["rows"]) &&
        Number(value["rows"]) >= 0 &&
        Number.isInteger(value["cols"]) &&
        Number(value["cols"]) >= 0
      );
    case "setFilter":
    case "clearFilter":
    case "clearSort":
      return isCellRangeRef(value["range"]);
    case "setSort":
      return isCellRangeRef(value["range"]) && isSortKeys(value["keys"]);
    case "updateRowMetadata": {
      const hasHeight = value["height"] !== undefined;
      const hasHidden = value["hidden"] !== undefined;
      return (
        typeof value["sheetName"] === "string" &&
        Number.isInteger(value["startRow"]) &&
        Number(value["startRow"]) >= 0 &&
        Number.isInteger(value["count"]) &&
        Number(value["count"]) > 0 &&
        (hasHeight || hasHidden) &&
        (!hasHeight ||
          value["height"] === null ||
          (typeof value["height"] === "number" &&
            Number.isFinite(value["height"]) &&
            value["height"] > 0)) &&
        (!hasHidden || value["hidden"] === null || typeof value["hidden"] === "boolean")
      );
    }
    case "updateColumnMetadata": {
      const hasWidth = value["width"] !== undefined;
      const hasHidden = value["hidden"] !== undefined;
      return (
        typeof value["sheetName"] === "string" &&
        Number.isInteger(value["startCol"]) &&
        Number(value["startCol"]) >= 0 &&
        Number.isInteger(value["count"]) &&
        Number(value["count"]) > 0 &&
        (hasWidth || hasHidden) &&
        (!hasWidth ||
          value["width"] === null ||
          (typeof value["width"] === "number" &&
            Number.isFinite(value["width"]) &&
            value["width"] > 0)) &&
        (!hasHidden || value["hidden"] === null || typeof value["hidden"] === "boolean")
      );
    }
    default:
      return false;
  }
}

export function isHighRiskWorkbookAgentStructuralCommand(
  command: WorkbookAgentStructuralCommand,
): boolean {
  return HIGH_RISK_STRUCTURAL_COMMAND_KINDS.has(command.kind);
}

export function isWorkbookScopeStructuralCommand(command: WorkbookAgentStructuralCommand): boolean {
  return WORKBOOK_SCOPE_STRUCTURAL_COMMAND_KINDS.has(command.kind);
}

export function describeWorkbookAgentStructuralCommand(
  command: WorkbookAgentStructuralCommand,
): string {
  switch (command.kind) {
    case "createSheet":
      return `Create sheet ${command.name}`;
    case "renameSheet":
      return `Rename sheet ${command.currentName} to ${command.nextName}`;
    case "deleteSheet":
      return `Delete sheet ${command.name}`;
    case "insertRows":
      return `Insert ${formatRowSpanLabel(command.start, command.count)} in ${command.sheetName}`;
    case "deleteRows":
      return `Delete ${formatRowSpanLabel(command.start, command.count)} in ${command.sheetName}`;
    case "insertColumns":
      return `Insert ${formatColumnSpanLabel(command.start, command.count)} in ${command.sheetName}`;
    case "deleteColumns":
      return `Delete ${formatColumnSpanLabel(command.start, command.count)} in ${command.sheetName}`;
    case "setFreezePane":
      return command.rows === 0 && command.cols === 0
        ? `Clear freeze panes on ${command.sheetName}`
        : `Freeze ${formatFreezePaneDescription(command.rows, command.cols)} on ${command.sheetName}`;
    case "setFilter":
      return `Set filter on ${rangeLabel(normalizeRangeBounds(command.range))}`;
    case "clearFilter":
      return `Clear filter on ${rangeLabel(normalizeRangeBounds(command.range))}`;
    case "setSort":
      return `Sort ${rangeLabel(normalizeRangeBounds(command.range))} by ${describeSortKeys(command.keys)}`;
    case "clearSort":
      return `Clear sort on ${rangeLabel(normalizeRangeBounds(command.range))}`;
    case "updateRowMetadata":
      return describeAxisMetadataCommand({
        axis: "row",
        sheetName: command.sheetName,
        spanLabel: formatRowSpanLabel(command.startRow, command.count),
        ...(hasOwnProperty(command, "height") ? { size: command.height ?? null } : {}),
        ...(hasOwnProperty(command, "hidden") ? { hidden: command.hidden ?? null } : {}),
      });
    case "updateColumnMetadata":
      return describeAxisMetadataCommand({
        axis: "column",
        sheetName: command.sheetName,
        spanLabel: formatColumnSpanLabel(command.startCol, command.count),
        ...(hasOwnProperty(command, "width") ? { size: command.width ?? null } : {}),
        ...(hasOwnProperty(command, "hidden") ? { hidden: command.hidden ?? null } : {}),
      });
    default: {
      const exhaustive: never = command;
      return exhaustive;
    }
  }
}

export function estimateWorkbookAgentStructuralCommandAffectedCells(
  command: WorkbookAgentStructuralCommand,
): number | null {
  switch (command.kind) {
    case "createSheet":
    case "renameSheet":
    case "deleteSheet":
    case "insertRows":
    case "deleteRows":
    case "insertColumns":
    case "deleteColumns":
    case "setFreezePane":
    case "setFilter":
    case "clearFilter":
    case "setSort":
    case "clearSort":
      return "range" in command ? countRangeCells(command.range) : null;
    case "updateRowMetadata":
    case "updateColumnMetadata":
      return null;
    default: {
      const exhaustive: never = command;
      return exhaustive;
    }
  }
}

export function deriveWorkbookAgentStructuralCommandPreviewRanges(
  command: WorkbookAgentStructuralCommand,
): WorkbookAgentPreviewRange[] {
  switch (command.kind) {
    case "createSheet":
    case "renameSheet":
    case "deleteSheet":
    case "setFreezePane":
      return [];
    case "insertRows":
    case "deleteRows":
      return [
        {
          sheetName: command.sheetName,
          startAddress: formatAddress(command.start, 0),
          endAddress: formatAddress(command.start + command.count - 1, 0),
          role: "target",
        },
      ];
    case "insertColumns":
    case "deleteColumns":
      return [
        {
          sheetName: command.sheetName,
          startAddress: formatAddress(0, command.start),
          endAddress: formatAddress(0, command.start + command.count - 1),
          role: "target",
        },
      ];
    case "setFilter":
    case "clearFilter":
    case "setSort":
    case "clearSort":
      return [
        {
          ...normalizeRangeBounds(command.range),
          role: "target",
        },
      ];
    case "updateRowMetadata":
      return [
        {
          sheetName: command.sheetName,
          startAddress: formatAddress(command.startRow, 0),
          endAddress: formatAddress(command.startRow + command.count - 1, 0),
          role: "target",
        },
      ];
    case "updateColumnMetadata":
      return [
        {
          sheetName: command.sheetName,
          startAddress: formatAddress(0, command.startCol),
          endAddress: formatAddress(0, command.startCol + command.count - 1),
          role: "target",
        },
      ];
    default: {
      const exhaustive: never = command;
      return exhaustive;
    }
  }
}

export function resolveRowMetadataCommandState(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookAgentStructuralCommand, { kind: "updateRowMetadata" }>,
): {
  height: number | null;
  hidden: boolean | null;
} {
  const spanLabel = formatRowSpanLabel(command.startRow, command.count);
  const entries = engine.getRowAxisEntries(command.sheetName);
  return {
    height: hasOwnProperty(command, "height")
      ? (command.height ?? null)
      : getConsistentAxisEntrySize({
          entries,
          start: command.startRow,
          count: command.count,
          spanLabel,
          propertyLabel: "row height",
        }),
    hidden: hasOwnProperty(command, "hidden")
      ? (command.hidden ?? null)
      : getConsistentAxisEntryHidden({
          entries,
          start: command.startRow,
          count: command.count,
          spanLabel,
          propertyLabel: "row visibility",
        }),
  };
}

export function resolveColumnMetadataCommandState(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookAgentStructuralCommand, { kind: "updateColumnMetadata" }>,
): {
  width: number | null;
  hidden: boolean | null;
} {
  const spanLabel = formatColumnSpanLabel(command.startCol, command.count);
  const entries = engine.getColumnAxisEntries(command.sheetName);
  return {
    width: hasOwnProperty(command, "width")
      ? (command.width ?? null)
      : getConsistentAxisEntrySize({
          entries,
          start: command.startCol,
          count: command.count,
          spanLabel,
          propertyLabel: "column width",
        }),
    hidden: hasOwnProperty(command, "hidden")
      ? (command.hidden ?? null)
      : getConsistentAxisEntryHidden({
          entries,
          start: command.startCol,
          count: command.count,
          spanLabel,
          propertyLabel: "column visibility",
        }),
  };
}

export function applyWorkbookAgentStructuralCommand(
  engine: SpreadsheetEngine,
  command: WorkbookAgentStructuralCommand,
): void {
  switch (command.kind) {
    case "createSheet":
      engine.renderCommit([
        {
          kind: "upsertSheet",
          name: command.name,
          order: engine.exportSnapshot().sheets.length,
        },
      ]);
      return;
    case "renameSheet":
      engine.renderCommit([
        {
          kind: "renameSheet",
          oldName: command.currentName,
          newName: command.nextName,
        },
      ]);
      return;
    case "deleteSheet":
      engine.deleteSheet(command.name);
      return;
    case "insertRows":
      engine.insertRows(command.sheetName, command.start, command.count);
      return;
    case "deleteRows":
      engine.deleteRows(command.sheetName, command.start, command.count);
      return;
    case "insertColumns":
      engine.insertColumns(command.sheetName, command.start, command.count);
      return;
    case "deleteColumns":
      engine.deleteColumns(command.sheetName, command.start, command.count);
      return;
    case "setFreezePane":
      if (command.rows === 0 && command.cols === 0) {
        engine.clearFreezePane(command.sheetName);
        return;
      }
      engine.setFreezePane(command.sheetName, command.rows, command.cols);
      return;
    case "setFilter":
      engine.setFilter(command.range.sheetName, command.range);
      return;
    case "clearFilter":
      engine.clearFilter(command.range.sheetName, command.range);
      return;
    case "setSort":
      engine.setSort(command.range.sheetName, command.range, command.keys);
      return;
    case "clearSort":
      engine.clearSort(command.range.sheetName, command.range);
      return;
    case "updateRowMetadata": {
      const resolved = resolveRowMetadataCommandState(engine, command);
      engine.updateRowMetadata(
        command.sheetName,
        command.startRow,
        command.count,
        resolved.height,
        resolved.hidden,
      );
      return;
    }
    case "updateColumnMetadata": {
      const resolved = resolveColumnMetadataCommandState(engine, command);
      engine.updateColumnMetadata(
        command.sheetName,
        command.startCol,
        command.count,
        resolved.width,
        resolved.hidden,
      );
      return;
    }
  }
}
