import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type {
  CellRangeRef,
  WorkbookDefinedNameValueSnapshot,
  WorkbookPivotSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookTableSnapshot,
} from "@bilig/protocol";
import type { WorkbookAgentCommand, WorkbookAgentPreviewRange } from "./workbook-agent-bundles.js";

export type WorkbookAgentObjectCommand = Extract<
  WorkbookAgentCommand,
  | { kind: "upsertDefinedName" }
  | { kind: "deleteDefinedName" }
  | { kind: "upsertTable" }
  | { kind: "deleteTable" }
  | { kind: "upsertPivotTable" }
  | { kind: "deletePivotTable" }
>;

const HIGH_RISK_OBJECT_COMMAND_KINDS = new Set<WorkbookAgentObjectCommand["kind"]>([
  "upsertDefinedName",
  "deleteDefinedName",
  "upsertTable",
  "deleteTable",
  "deletePivotTable",
]);

const WORKBOOK_SCOPE_OBJECT_COMMAND_KINDS = new Set<WorkbookAgentObjectCommand["kind"]>([
  "upsertDefinedName",
  "deleteDefinedName",
  "upsertTable",
  "deleteTable",
]);

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

function isWorkbookDefinedNameValueSnapshot(
  value: unknown,
): value is WorkbookDefinedNameValueSnapshot {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return true;
  }
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "scalar":
      return (
        value["value"] === null ||
        typeof value["value"] === "string" ||
        typeof value["value"] === "number" ||
        typeof value["value"] === "boolean"
      );
    case "cell-ref":
      return typeof value["sheetName"] === "string" && typeof value["address"] === "string";
    case "range-ref":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["startAddress"] === "string" &&
        typeof value["endAddress"] === "string"
      );
    case "structured-ref":
      return typeof value["tableName"] === "string" && typeof value["columnName"] === "string";
    case "formula":
      return typeof value["formula"] === "string" && value["formula"].trim().length > 0;
    default:
      return false;
  }
}

function isWorkbookTableSnapshot(value: unknown): value is WorkbookTableSnapshot {
  return (
    isRecord(value) &&
    typeof value["name"] === "string" &&
    typeof value["sheetName"] === "string" &&
    typeof value["startAddress"] === "string" &&
    typeof value["endAddress"] === "string" &&
    Array.isArray(value["columnNames"]) &&
    value["columnNames"].every((entry) => typeof entry === "string") &&
    typeof value["headerRow"] === "boolean" &&
    typeof value["totalsRow"] === "boolean"
  );
}

function isWorkbookPivotValueSnapshots(
  value: unknown,
): value is readonly WorkbookPivotValueSnapshot[] {
  return (
    Array.isArray(value) &&
    value.every(
      (entry) =>
        isRecord(entry) &&
        typeof entry["sourceColumn"] === "string" &&
        (entry["summarizeBy"] === "sum" || entry["summarizeBy"] === "count") &&
        (entry["outputLabel"] === undefined || typeof entry["outputLabel"] === "string"),
    )
  );
}

function isWorkbookPivotSnapshot(value: unknown): value is WorkbookPivotSnapshot {
  return (
    isRecord(value) &&
    typeof value["name"] === "string" &&
    typeof value["sheetName"] === "string" &&
    typeof value["address"] === "string" &&
    isCellRangeRef(value["source"]) &&
    Array.isArray(value["groupBy"]) &&
    value["groupBy"].every((entry) => typeof entry === "string") &&
    isWorkbookPivotValueSnapshots(value["values"]) &&
    typeof value["rows"] === "number" &&
    typeof value["cols"] === "number"
  );
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

function rangeLabel(
  range: Pick<CellRangeRef, "sheetName" | "startAddress" | "endAddress">,
): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`;
}

function countRangeCells(range: CellRangeRef): number {
  const normalized = normalizeRangeBounds(range);
  const start = parseCellAddress(normalized.startAddress, normalized.sheetName);
  const end = parseCellAddress(normalized.endAddress, normalized.sheetName);
  return (end.row - start.row + 1) * (end.col - start.col + 1);
}

function pivotTargetRange(pivot: WorkbookPivotSnapshot): CellRangeRef {
  const anchor = parseCellAddress(pivot.address, pivot.sheetName);
  return {
    sheetName: pivot.sheetName,
    startAddress: pivot.address,
    endAddress: formatAddress(
      anchor.row + Math.max(pivot.rows, 1) - 1,
      anchor.col + Math.max(pivot.cols, 1) - 1,
    ),
  };
}

function describeDefinedNameValue(value: WorkbookDefinedNameValueSnapshot): string {
  if (typeof value === "string" && value.startsWith("=")) {
    return "formula";
  }
  if (typeof value !== "object" || value === null || !("kind" in value)) {
    return "scalar";
  }
  switch (value.kind) {
    case "scalar":
      return "scalar";
    case "cell-ref":
      return `${value.sheetName}!${value.address}`;
    case "range-ref":
      return `${value.sheetName}!${value.startAddress}:${value.endAddress}`;
    case "structured-ref":
      return `${value.tableName}[${value.columnName}]`;
    case "formula":
      return "formula";
  }
}

export function isWorkbookAgentObjectCommandKind(
  kind: string,
): kind is WorkbookAgentObjectCommand["kind"] {
  switch (kind) {
    case "upsertDefinedName":
    case "deleteDefinedName":
    case "upsertTable":
    case "deleteTable":
    case "upsertPivotTable":
    case "deletePivotTable":
      return true;
    default:
      return false;
  }
}

export function isWorkbookAgentObjectCommand(
  command: WorkbookAgentCommand,
): command is WorkbookAgentObjectCommand {
  return isWorkbookAgentObjectCommandKind(command.kind);
}

export function isWorkbookAgentObjectCommandValue(
  value: unknown,
): value is WorkbookAgentObjectCommand {
  if (!isRecord(value)) {
    return false;
  }
  const kind = value["kind"];
  if (typeof kind !== "string" || !isWorkbookAgentObjectCommandKind(kind)) {
    return false;
  }
  switch (kind) {
    case "upsertDefinedName":
      return (
        typeof value["name"] === "string" &&
        value["name"].trim().length > 0 &&
        isWorkbookDefinedNameValueSnapshot(value["value"])
      );
    case "deleteDefinedName":
    case "deleteTable":
      return typeof value["name"] === "string" && value["name"].trim().length > 0;
    case "upsertTable":
      return isWorkbookTableSnapshot(value["table"]);
    case "upsertPivotTable":
      return isWorkbookPivotSnapshot(value["pivot"]);
    case "deletePivotTable":
      return typeof value["sheetName"] === "string" && typeof value["address"] === "string";
  }
}

export function isHighRiskWorkbookAgentObjectCommand(command: WorkbookAgentObjectCommand): boolean {
  return HIGH_RISK_OBJECT_COMMAND_KINDS.has(command.kind);
}

export function isWorkbookScopeObjectCommand(command: WorkbookAgentObjectCommand): boolean {
  return WORKBOOK_SCOPE_OBJECT_COMMAND_KINDS.has(command.kind);
}

export function describeWorkbookAgentObjectCommand(command: WorkbookAgentObjectCommand): string {
  switch (command.kind) {
    case "upsertDefinedName":
      return `Set named range ${command.name} to ${describeDefinedNameValue(command.value)}`;
    case "deleteDefinedName":
      return `Delete named range ${command.name}`;
    case "upsertTable":
      return `Set table ${command.table.name} on ${rangeLabel(command.table)}`;
    case "deleteTable":
      return `Delete table ${command.name}`;
    case "upsertPivotTable":
      return `Set pivot ${command.pivot.name} at ${command.pivot.sheetName}!${command.pivot.address}`;
    case "deletePivotTable":
      return `Delete pivot at ${command.sheetName}!${command.address}`;
  }
}

export function estimateWorkbookAgentObjectCommandAffectedCells(
  command: WorkbookAgentObjectCommand,
): number | null {
  switch (command.kind) {
    case "upsertDefinedName":
    case "deleteDefinedName":
    case "deleteTable":
      return null;
    case "upsertTable":
      return countRangeCells(command.table);
    case "upsertPivotTable":
      return Math.max(command.pivot.rows, 1) * Math.max(command.pivot.cols, 1);
    case "deletePivotTable":
      return null;
  }
}

export function deriveWorkbookAgentObjectCommandPreviewRanges(
  command: WorkbookAgentObjectCommand,
): WorkbookAgentPreviewRange[] {
  switch (command.kind) {
    case "upsertDefinedName": {
      const value = command.value;
      if (typeof value !== "object" || value === null || !("kind" in value)) {
        return [];
      }
      switch (value.kind) {
        case "cell-ref":
          return [
            {
              sheetName: value.sheetName,
              startAddress: value.address,
              endAddress: value.address,
              role: "target",
            },
          ];
        case "range-ref":
          return [
            {
              ...normalizeRangeBounds({
                sheetName: value.sheetName,
                startAddress: value.startAddress,
                endAddress: value.endAddress,
              }),
              role: "target",
            },
          ];
        case "formula":
        case "scalar":
        case "structured-ref":
          return [];
        default:
          const exhaustive: never = value;
          return exhaustive;
      }
    }
    case "deleteDefinedName":
    case "deleteTable":
      return [];
    case "upsertTable":
      return [
        {
          ...normalizeRangeBounds(command.table),
          role: "target",
        },
      ];
    case "upsertPivotTable":
      return [
        {
          ...normalizeRangeBounds(command.pivot.source),
          role: "source",
        },
        {
          ...normalizeRangeBounds(pivotTargetRange(command.pivot)),
          role: "target",
        },
      ];
    case "deletePivotTable":
      return [
        {
          sheetName: command.sheetName,
          startAddress: command.address,
          endAddress: command.address,
          role: "target",
        },
      ];
  }
}

export function applyWorkbookAgentObjectCommand(
  engine: SpreadsheetEngine,
  command: WorkbookAgentObjectCommand,
): void {
  switch (command.kind) {
    case "upsertDefinedName":
      engine.setDefinedName(command.name, command.value);
      return;
    case "deleteDefinedName":
      engine.deleteDefinedName(command.name);
      return;
    case "upsertTable":
      engine.setTable(command.table);
      return;
    case "deleteTable":
      engine.deleteTable(command.name);
      return;
    case "upsertPivotTable":
      engine.setPivotTable(command.pivot.sheetName, command.pivot.address, {
        name: command.pivot.name,
        source: command.pivot.source,
        groupBy: command.pivot.groupBy,
        values: command.pivot.values,
      });
      return;
    case "deletePivotTable":
      engine.deletePivotTable(command.sheetName, command.address);
      return;
  }
}
