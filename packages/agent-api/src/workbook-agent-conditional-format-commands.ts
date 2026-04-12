import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type {
  CellRangeRef,
  WorkbookConditionalFormatRuleSnapshot,
  WorkbookConditionalFormatSnapshot,
} from "@bilig/protocol";
import type { WorkbookAgentCommand, WorkbookAgentPreviewRange } from "./workbook-agent-bundles.js";

export type WorkbookAgentConditionalFormatCommand = Extract<
  WorkbookAgentCommand,
  { kind: "upsertConditionalFormat" } | { kind: "deleteConditionalFormat" }
>;

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

function isConditionalFormatRule(value: unknown): value is WorkbookConditionalFormatRuleSnapshot {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "cellIs":
      return (
        typeof value["operator"] === "string" &&
        Array.isArray(value["values"]) &&
        value["values"].every(
          (entry) =>
            entry === null ||
            typeof entry === "string" ||
            typeof entry === "number" ||
            typeof entry === "boolean",
        ) &&
        (value["operator"] === "between" || value["operator"] === "notBetween"
          ? value["values"].length === 2
          : value["values"].length === 1)
      );
    case "textContains":
      return (
        typeof value["text"] === "string" &&
        value["text"].trim().length > 0 &&
        (value["caseSensitive"] === undefined || typeof value["caseSensitive"] === "boolean")
      );
    case "formula":
      return typeof value["formula"] === "string" && value["formula"].trim().length > 0;
    case "blanks":
    case "notBlanks":
      return true;
    default:
      return false;
  }
}

function isConditionalFormatSnapshot(value: unknown): value is WorkbookConditionalFormatSnapshot {
  return (
    isRecord(value) &&
    typeof value["id"] === "string" &&
    value["id"].trim().length > 0 &&
    isCellRangeRef(value["range"]) &&
    isConditionalFormatRule(value["rule"]) &&
    isRecord(value["style"]) &&
    (value["stopIfTrue"] === undefined || typeof value["stopIfTrue"] === "boolean") &&
    (value["priority"] === undefined || typeof value["priority"] === "number")
  );
}

function normalizeRangeBounds(range: CellRangeRef): CellRangeRef {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(Math.min(start.row, end.row), Math.min(start.col, end.col)),
    endAddress: formatAddress(Math.max(start.row, end.row), Math.max(start.col, end.col)),
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

export function isWorkbookAgentConditionalFormatCommandKind(
  kind: string,
): kind is WorkbookAgentConditionalFormatCommand["kind"] {
  return kind === "upsertConditionalFormat" || kind === "deleteConditionalFormat";
}

export function isWorkbookAgentConditionalFormatCommand(
  command: WorkbookAgentCommand,
): command is WorkbookAgentConditionalFormatCommand {
  return isWorkbookAgentConditionalFormatCommandKind(command.kind);
}

export function isWorkbookAgentConditionalFormatCommandValue(
  value: unknown,
): value is WorkbookAgentConditionalFormatCommand {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "upsertConditionalFormat":
      return isConditionalFormatSnapshot(value["format"]);
    case "deleteConditionalFormat":
      return typeof value["id"] === "string" && isCellRangeRef(value["range"]);
    default:
      return false;
  }
}

export function isHighRiskWorkbookAgentConditionalFormatCommand(
  _command: WorkbookAgentConditionalFormatCommand,
): boolean {
  return false;
}

export function isWorkbookScopeConditionalFormatCommand(
  _command: WorkbookAgentConditionalFormatCommand,
): boolean {
  return false;
}

export function describeWorkbookAgentConditionalFormatCommand(
  command: WorkbookAgentConditionalFormatCommand,
): string {
  switch (command.kind) {
    case "upsertConditionalFormat":
      return `Set conditional format on ${rangeLabel(normalizeRangeBounds(command.format.range))}`;
    case "deleteConditionalFormat":
      return `Remove conditional format from ${rangeLabel(normalizeRangeBounds(command.range))}`;
  }
}

export function estimateWorkbookAgentConditionalFormatCommandAffectedCells(
  command: WorkbookAgentConditionalFormatCommand,
): number {
  return countRangeCells(
    command.kind === "upsertConditionalFormat" ? command.format.range : command.range,
  );
}

export function deriveWorkbookAgentConditionalFormatCommandPreviewRanges(
  command: WorkbookAgentConditionalFormatCommand,
): WorkbookAgentPreviewRange[] {
  return [
    {
      ...normalizeRangeBounds(
        command.kind === "upsertConditionalFormat" ? command.format.range : command.range,
      ),
      role: "target",
    },
  ];
}

export function applyWorkbookAgentConditionalFormatCommand(
  engine: SpreadsheetEngine,
  command: WorkbookAgentConditionalFormatCommand,
): void {
  switch (command.kind) {
    case "upsertConditionalFormat":
      engine.setConditionalFormat(structuredClone(command.format));
      return;
    case "deleteConditionalFormat":
      engine.deleteConditionalFormat(command.id);
      return;
  }
}
