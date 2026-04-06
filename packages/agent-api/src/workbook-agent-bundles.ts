import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
} from "@bilig/protocol";

export interface WorkbookAgentUiSelectionRef {
  sheetName: string;
  address: string;
}

export interface WorkbookAgentViewportRef {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

export interface WorkbookAgentContextRef {
  selection: WorkbookAgentUiSelectionRef;
  viewport: WorkbookAgentViewportRef;
}

export type WorkbookAgentWriteCellInput =
  | LiteralInput
  | {
      value: LiteralInput;
    }
  | {
      formula: string;
    };

export type WorkbookAgentCommand =
  | {
      kind: "writeRange";
      sheetName: string;
      startAddress: string;
      values: WorkbookAgentWriteCellInput[][];
    }
  | {
      kind: "clearRange";
      range: CellRangeRef;
    }
  | {
      kind: "formatRange";
      range: CellRangeRef;
      patch?: CellStylePatch;
      numberFormat?: CellNumberFormatInput;
    }
  | {
      kind: "fillRange";
      source: CellRangeRef;
      target: CellRangeRef;
    }
  | {
      kind: "copyRange";
      source: CellRangeRef;
      target: CellRangeRef;
    }
  | {
      kind: "moveRange";
      source: CellRangeRef;
      target: CellRangeRef;
    }
  | {
      kind: "createSheet";
      name: string;
    }
  | {
      kind: "renameSheet";
      currentName: string;
      nextName: string;
    };

export type WorkbookAgentRiskClass = "low" | "medium" | "high";
export type WorkbookAgentBundleScope = "selection" | "sheet" | "workbook";
export type WorkbookAgentApprovalMode = "auto" | "preview" | "explicit";
export type WorkbookAgentAppliedBy = "user" | "auto";
export type WorkbookAgentAcceptedScope = "full";
export type WorkbookAgentPreviewRangeRole = "target" | "source";
export type WorkbookAgentPreviewChangeKind = "input" | "formula" | "style" | "numberFormat";

export interface WorkbookAgentPreviewRange {
  sheetName: string;
  startAddress: string;
  endAddress: string;
  role: WorkbookAgentPreviewRangeRole;
}

export interface WorkbookAgentPreviewCellDiff {
  sheetName: string;
  address: string;
  beforeInput: LiteralInput | null;
  beforeFormula: string | null;
  afterInput: LiteralInput | null;
  afterFormula: string | null;
  changeKinds: WorkbookAgentPreviewChangeKind[];
}

export interface WorkbookAgentPreviewEffectSummary {
  displayedCellDiffCount: number;
  truncatedCellDiffs: boolean;
  inputChangeCount: number;
  formulaChangeCount: number;
  styleChangeCount: number;
  numberFormatChangeCount: number;
  structuralChangeCount: number;
}

export interface WorkbookAgentPreviewSummary {
  ranges: WorkbookAgentPreviewRange[];
  structuralChanges: string[];
  cellDiffs: WorkbookAgentPreviewCellDiff[];
  effectSummary: WorkbookAgentPreviewEffectSummary;
}

export interface WorkbookAgentCommandBundle {
  id: string;
  documentId: string;
  threadId: string;
  turnId: string;
  goalText: string;
  summary: string;
  scope: WorkbookAgentBundleScope;
  riskClass: WorkbookAgentRiskClass;
  approvalMode: WorkbookAgentApprovalMode;
  baseRevision: number;
  createdAtUnixMs: number;
  context: WorkbookAgentContextRef | null;
  commands: WorkbookAgentCommand[];
  affectedRanges: WorkbookAgentPreviewRange[];
  estimatedAffectedCells: number | null;
}

export interface WorkbookAgentExecutionRecord {
  id: string;
  bundleId: string;
  documentId: string;
  threadId: string;
  turnId: string;
  actorUserId: string;
  goalText: string;
  planText: string | null;
  summary: string;
  scope: WorkbookAgentBundleScope;
  riskClass: WorkbookAgentRiskClass;
  approvalMode: WorkbookAgentApprovalMode;
  acceptedScope: WorkbookAgentAcceptedScope;
  appliedBy: WorkbookAgentAppliedBy;
  baseRevision: number;
  appliedRevision: number;
  createdAtUnixMs: number;
  appliedAtUnixMs: number;
  context: WorkbookAgentContextRef | null;
  commands: WorkbookAgentCommand[];
  preview: WorkbookAgentPreviewSummary | null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isLiteralInputValue(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  );
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["startAddress"] === "string" &&
    typeof value["endAddress"] === "string"
  );
}

function isWriteCellInput(value: unknown): value is WorkbookAgentWriteCellInput {
  return (
    isLiteralInputValue(value) ||
    (isRecord(value) && isLiteralInputValue(value["value"])) ||
    (isRecord(value) && typeof value["formula"] === "string" && value["formula"].length > 0)
  );
}

function isCommandArray(value: unknown): value is WorkbookAgentCommand[] {
  return Array.isArray(value) && value.every((entry) => isWorkbookAgentCommand(entry));
}

function isApprovalMode(value: unknown): value is WorkbookAgentApprovalMode {
  return value === "auto" || value === "preview" || value === "explicit";
}

function isAppliedBy(value: unknown): value is WorkbookAgentAppliedBy {
  return value === "user" || value === "auto";
}

function isAcceptedScope(value: unknown): value is WorkbookAgentAcceptedScope {
  return value === "full";
}

function isPreviewChangeKind(value: unknown): value is WorkbookAgentPreviewChangeKind {
  return value === "input" || value === "formula" || value === "style" || value === "numberFormat";
}

export function isWorkbookAgentContextRef(value: unknown): value is WorkbookAgentContextRef {
  return (
    isRecord(value) &&
    isRecord(value["selection"]) &&
    typeof value["selection"]["sheetName"] === "string" &&
    typeof value["selection"]["address"] === "string" &&
    isRecord(value["viewport"]) &&
    typeof value["viewport"]["rowStart"] === "number" &&
    typeof value["viewport"]["rowEnd"] === "number" &&
    typeof value["viewport"]["colStart"] === "number" &&
    typeof value["viewport"]["colEnd"] === "number"
  );
}

export function isWorkbookAgentPreviewRange(value: unknown): value is WorkbookAgentPreviewRange {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["startAddress"] === "string" &&
    typeof value["endAddress"] === "string" &&
    (value["role"] === "target" || value["role"] === "source")
  );
}

export function isWorkbookAgentPreviewCellDiff(
  value: unknown,
): value is WorkbookAgentPreviewCellDiff {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["address"] === "string" &&
    (value["beforeInput"] === null || isLiteralInputValue(value["beforeInput"])) &&
    (value["beforeFormula"] === null || typeof value["beforeFormula"] === "string") &&
    (value["afterInput"] === null || isLiteralInputValue(value["afterInput"])) &&
    (value["afterFormula"] === null || typeof value["afterFormula"] === "string") &&
    Array.isArray(value["changeKinds"]) &&
    value["changeKinds"].every((entry) => isPreviewChangeKind(entry))
  );
}

export function isWorkbookAgentPreviewEffectSummary(
  value: unknown,
): value is WorkbookAgentPreviewEffectSummary {
  return (
    isRecord(value) &&
    typeof value["displayedCellDiffCount"] === "number" &&
    Number.isFinite(value["displayedCellDiffCount"]) &&
    typeof value["truncatedCellDiffs"] === "boolean" &&
    typeof value["inputChangeCount"] === "number" &&
    Number.isFinite(value["inputChangeCount"]) &&
    typeof value["formulaChangeCount"] === "number" &&
    Number.isFinite(value["formulaChangeCount"]) &&
    typeof value["styleChangeCount"] === "number" &&
    Number.isFinite(value["styleChangeCount"]) &&
    typeof value["numberFormatChangeCount"] === "number" &&
    Number.isFinite(value["numberFormatChangeCount"]) &&
    typeof value["structuralChangeCount"] === "number" &&
    Number.isFinite(value["structuralChangeCount"])
  );
}

export function isWorkbookAgentPreviewSummary(
  value: unknown,
): value is WorkbookAgentPreviewSummary {
  return (
    isRecord(value) &&
    Array.isArray(value["ranges"]) &&
    value["ranges"].every((entry) => isWorkbookAgentPreviewRange(entry)) &&
    Array.isArray(value["structuralChanges"]) &&
    value["structuralChanges"].every((entry) => typeof entry === "string") &&
    Array.isArray(value["cellDiffs"]) &&
    value["cellDiffs"].every((entry) => isWorkbookAgentPreviewCellDiff(entry)) &&
    isWorkbookAgentPreviewEffectSummary(value["effectSummary"])
  );
}

function derivePreviewEffectSummary(input: {
  cellDiffs: readonly WorkbookAgentPreviewCellDiff[];
  structuralChanges: readonly string[];
  truncatedCellDiffs: boolean;
}): WorkbookAgentPreviewEffectSummary {
  return {
    displayedCellDiffCount: input.cellDiffs.length,
    truncatedCellDiffs: input.truncatedCellDiffs,
    inputChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes("input")).length,
    formulaChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes("formula"))
      .length,
    styleChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes("style")).length,
    numberFormatChangeCount: input.cellDiffs.filter((diff) =>
      diff.changeKinds.includes("numberFormat"),
    ).length,
    structuralChangeCount: input.structuralChanges.length,
  };
}

function decodeWorkbookAgentPreviewCellDiff(value: unknown): WorkbookAgentPreviewCellDiff | null {
  if (!isRecord(value)) {
    return null;
  }
  if (
    typeof value["sheetName"] !== "string" ||
    typeof value["address"] !== "string" ||
    (value["beforeInput"] !== null && !isLiteralInputValue(value["beforeInput"])) ||
    (value["beforeFormula"] !== null && typeof value["beforeFormula"] !== "string") ||
    (value["afterInput"] !== null && !isLiteralInputValue(value["afterInput"])) ||
    (value["afterFormula"] !== null && typeof value["afterFormula"] !== "string")
  ) {
    return null;
  }
  const explicitChangeKinds = Array.isArray(value["changeKinds"])
    ? value["changeKinds"].flatMap((entry) => (isPreviewChangeKind(entry) ? [entry] : []))
    : [];
  const derivedChangeKinds = explicitChangeKinds.length
    ? explicitChangeKinds
    : [
        ...(value["beforeFormula"] !== value["afterFormula"] ? (["formula"] as const) : []),
        ...(value["beforeInput"] !== value["afterInput"] ? (["input"] as const) : []),
      ];
  return {
    sheetName: value["sheetName"],
    address: value["address"],
    beforeInput: (value["beforeInput"] as LiteralInput | null | undefined) ?? null,
    beforeFormula: (value["beforeFormula"] as string | null | undefined) ?? null,
    afterInput: (value["afterInput"] as LiteralInput | null | undefined) ?? null,
    afterFormula: (value["afterFormula"] as string | null | undefined) ?? null,
    changeKinds: [...new Set(derivedChangeKinds)],
  };
}

export function decodeWorkbookAgentPreviewSummary(
  value: unknown,
): WorkbookAgentPreviewSummary | null {
  if (!isRecord(value)) {
    return null;
  }
  if (
    !Array.isArray(value["ranges"]) ||
    !value["ranges"].every((entry) => isWorkbookAgentPreviewRange(entry)) ||
    !Array.isArray(value["structuralChanges"]) ||
    !value["structuralChanges"].every((entry) => typeof entry === "string") ||
    !Array.isArray(value["cellDiffs"])
  ) {
    return null;
  }
  const cellDiffs = value["cellDiffs"].flatMap((entry) => {
    const decoded = decodeWorkbookAgentPreviewCellDiff(entry);
    return decoded ? [decoded] : [];
  });
  if (cellDiffs.length !== value["cellDiffs"].length) {
    return null;
  }
  const truncatedCellDiffs = isWorkbookAgentPreviewEffectSummary(value["effectSummary"])
    ? value["effectSummary"].truncatedCellDiffs
    : false;
  return {
    ranges: value["ranges"].map((range) => ({ ...range })),
    structuralChanges: [...value["structuralChanges"]],
    cellDiffs,
    effectSummary: isWorkbookAgentPreviewEffectSummary(value["effectSummary"])
      ? { ...value["effectSummary"] }
      : derivePreviewEffectSummary({
          cellDiffs,
          structuralChanges: value["structuralChanges"],
          truncatedCellDiffs,
        }),
  };
}

function samePreviewRange(
  left: WorkbookAgentPreviewRange,
  right: WorkbookAgentPreviewRange,
): boolean {
  return (
    left.sheetName === right.sheetName &&
    left.startAddress === right.startAddress &&
    left.endAddress === right.endAddress &&
    left.role === right.role
  );
}

function samePreviewCellDiff(
  left: WorkbookAgentPreviewCellDiff,
  right: WorkbookAgentPreviewCellDiff,
): boolean {
  return (
    left.sheetName === right.sheetName &&
    left.address === right.address &&
    left.beforeInput === right.beforeInput &&
    left.beforeFormula === right.beforeFormula &&
    left.afterInput === right.afterInput &&
    left.afterFormula === right.afterFormula &&
    left.changeKinds.length === right.changeKinds.length &&
    left.changeKinds.every((kind, index) => kind === right.changeKinds[index])
  );
}

export function areWorkbookAgentPreviewSummariesEqual(
  left: WorkbookAgentPreviewSummary,
  right: WorkbookAgentPreviewSummary,
): boolean {
  return (
    left.ranges.length === right.ranges.length &&
    left.ranges.every((range, index) => {
      const other = right.ranges[index];
      return other ? samePreviewRange(range, other) : false;
    }) &&
    left.structuralChanges.length === right.structuralChanges.length &&
    left.structuralChanges.every((change, index) => change === right.structuralChanges[index]) &&
    left.cellDiffs.length === right.cellDiffs.length &&
    left.cellDiffs.every((diff, index) => {
      const other = right.cellDiffs[index];
      return other ? samePreviewCellDiff(diff, other) : false;
    }) &&
    left.effectSummary.displayedCellDiffCount === right.effectSummary.displayedCellDiffCount &&
    left.effectSummary.truncatedCellDiffs === right.effectSummary.truncatedCellDiffs &&
    left.effectSummary.inputChangeCount === right.effectSummary.inputChangeCount &&
    left.effectSummary.formulaChangeCount === right.effectSummary.formulaChangeCount &&
    left.effectSummary.styleChangeCount === right.effectSummary.styleChangeCount &&
    left.effectSummary.numberFormatChangeCount === right.effectSummary.numberFormatChangeCount &&
    left.effectSummary.structuralChangeCount === right.effectSummary.structuralChangeCount
  );
}

export function isWorkbookAgentCommand(value: unknown): value is WorkbookAgentCommand {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "writeRange":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["startAddress"] === "string" &&
        Array.isArray(value["values"]) &&
        value["values"].every(
          (row) =>
            Array.isArray(row) &&
            row.length > 0 &&
            row.every((cellValue) => isWriteCellInput(cellValue)),
        )
      );
    case "clearRange":
      return isCellRangeRef(value["range"]);
    case "formatRange":
      return (
        isCellRangeRef(value["range"]) &&
        (value["patch"] === undefined || isRecord(value["patch"])) &&
        (value["numberFormat"] === undefined ||
          typeof value["numberFormat"] === "string" ||
          isRecord(value["numberFormat"]))
      );
    case "fillRange":
    case "copyRange":
    case "moveRange":
      return isCellRangeRef(value["source"]) && isCellRangeRef(value["target"]);
    case "createSheet":
      return typeof value["name"] === "string" && value["name"].trim().length > 0;
    case "renameSheet":
      return (
        typeof value["currentName"] === "string" &&
        value["currentName"].trim().length > 0 &&
        typeof value["nextName"] === "string" &&
        value["nextName"].trim().length > 0
      );
    default:
      return false;
  }
}

export function isWorkbookAgentCommandBundle(value: unknown): value is WorkbookAgentCommandBundle {
  return (
    isRecord(value) &&
    typeof value["id"] === "string" &&
    typeof value["documentId"] === "string" &&
    typeof value["threadId"] === "string" &&
    typeof value["turnId"] === "string" &&
    typeof value["goalText"] === "string" &&
    typeof value["summary"] === "string" &&
    (value["scope"] === "selection" ||
      value["scope"] === "sheet" ||
      value["scope"] === "workbook") &&
    (value["riskClass"] === "low" ||
      value["riskClass"] === "medium" ||
      value["riskClass"] === "high") &&
    isApprovalMode(value["approvalMode"]) &&
    typeof value["baseRevision"] === "number" &&
    typeof value["createdAtUnixMs"] === "number" &&
    (value["context"] === null || isWorkbookAgentContextRef(value["context"])) &&
    isCommandArray(value["commands"]) &&
    Array.isArray(value["affectedRanges"]) &&
    value["affectedRanges"].every((entry) => isWorkbookAgentPreviewRange(entry)) &&
    (value["estimatedAffectedCells"] === null ||
      typeof value["estimatedAffectedCells"] === "number")
  );
}

export function isWorkbookAgentExecutionRecord(
  value: unknown,
): value is WorkbookAgentExecutionRecord {
  return (
    isRecord(value) &&
    typeof value["id"] === "string" &&
    typeof value["bundleId"] === "string" &&
    typeof value["documentId"] === "string" &&
    typeof value["threadId"] === "string" &&
    typeof value["turnId"] === "string" &&
    typeof value["actorUserId"] === "string" &&
    typeof value["goalText"] === "string" &&
    (value["planText"] === null || typeof value["planText"] === "string") &&
    typeof value["summary"] === "string" &&
    (value["scope"] === "selection" ||
      value["scope"] === "sheet" ||
      value["scope"] === "workbook") &&
    (value["riskClass"] === "low" ||
      value["riskClass"] === "medium" ||
      value["riskClass"] === "high") &&
    isApprovalMode(value["approvalMode"]) &&
    isAcceptedScope(value["acceptedScope"]) &&
    isAppliedBy(value["appliedBy"]) &&
    typeof value["baseRevision"] === "number" &&
    typeof value["appliedRevision"] === "number" &&
    typeof value["createdAtUnixMs"] === "number" &&
    typeof value["appliedAtUnixMs"] === "number" &&
    (value["context"] === null || isWorkbookAgentContextRef(value["context"])) &&
    isCommandArray(value["commands"]) &&
    (value["preview"] === null || isWorkbookAgentPreviewSummary(value["preview"]))
  );
}

function normalizeFormula(formula: string): string {
  return formula.startsWith("=") ? formula.slice(1) : formula;
}

function normalizeRangeBounds(range: CellRangeRef): {
  sheetName: string;
  startAddress: string;
  endAddress: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
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
    startRow,
    endRow,
    startCol,
    endCol,
  };
}

function countRangeCells(range: CellRangeRef): number {
  const bounds = normalizeRangeBounds(range);
  return (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1);
}

function writeRangeToRange(
  command: Extract<WorkbookAgentCommand, { kind: "writeRange" }>,
): CellRangeRef {
  const start = parseCellAddress(command.startAddress, command.sheetName);
  const width = command.values.reduce((maxWidth, row) => Math.max(maxWidth, row.length), 0);
  return {
    sheetName: command.sheetName,
    startAddress: command.startAddress,
    endAddress: formatAddress(start.row + command.values.length - 1, start.col + width - 1),
  };
}

export function estimateWorkbookAgentCommandAffectedCells(
  command: WorkbookAgentCommand,
): number | null {
  switch (command.kind) {
    case "writeRange":
      return command.values.reduce((sum, row) => sum + row.length, 0);
    case "clearRange":
    case "formatRange":
      return countRangeCells(command.range);
    case "fillRange":
    case "copyRange":
    case "moveRange":
      return countRangeCells(command.target);
    case "createSheet":
    case "renameSheet":
      return null;
    default: {
      const exhaustive: never = command;
      return exhaustive;
    }
  }
}

export function deriveWorkbookAgentCommandPreviewRanges(
  command: WorkbookAgentCommand,
): WorkbookAgentPreviewRange[] {
  switch (command.kind) {
    case "writeRange":
      return [
        {
          ...writeRangeToRange(command),
          role: "target",
        },
      ];
    case "clearRange":
    case "formatRange":
      return [
        {
          ...normalizeRangeBounds(command.range),
          role: "target",
        },
      ];
    case "fillRange":
    case "copyRange":
    case "moveRange":
      return [
        {
          ...normalizeRangeBounds(command.source),
          role: "source",
        },
        {
          ...normalizeRangeBounds(command.target),
          role: "target",
        },
      ];
    case "createSheet":
    case "renameSheet":
      return [];
    default: {
      const exhaustive: never = command;
      return exhaustive;
    }
  }
}

export function applyWorkbookAgentCommand(
  engine: SpreadsheetEngine,
  command: WorkbookAgentCommand,
): void {
  switch (command.kind) {
    case "writeRange": {
      const start = parseCellAddress(command.startAddress, command.sheetName);
      command.values.forEach((rowValues, rowOffset) => {
        rowValues.forEach((cellInput, colOffset) => {
          const address = formatAddress(start.row + rowOffset, start.col + colOffset);
          if (cellInput === null) {
            engine.clearCell(command.sheetName, address);
            return;
          }
          if (
            typeof cellInput === "string" ||
            typeof cellInput === "number" ||
            typeof cellInput === "boolean"
          ) {
            engine.setCellValue(command.sheetName, address, cellInput);
            return;
          }
          if ("formula" in cellInput) {
            engine.setCellFormula(command.sheetName, address, normalizeFormula(cellInput.formula));
            return;
          }
          engine.setCellValue(command.sheetName, address, cellInput.value);
        });
      });
      return;
    }
    case "clearRange":
      engine.clearRange(command.range);
      return;
    case "formatRange":
      if (command.patch !== undefined) {
        engine.setRangeStyle(command.range, command.patch);
      }
      if (command.numberFormat !== undefined) {
        engine.setRangeNumberFormat(command.range, command.numberFormat);
      }
      return;
    case "fillRange":
      engine.fillRange(command.source, command.target);
      return;
    case "copyRange":
      engine.copyRange(command.source, command.target);
      return;
    case "moveRange":
      engine.moveRange(command.source, command.target);
      return;
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
    default: {
      const exhaustive: never = command;
      throw new Error(`Unhandled workbook agent command: ${JSON.stringify(exhaustive)}`);
    }
  }
}

export function applyWorkbookAgentCommandBundle(
  engine: SpreadsheetEngine,
  bundle: Pick<WorkbookAgentCommandBundle, "commands">,
): void {
  bundle.commands.forEach((command) => {
    applyWorkbookAgentCommand(engine, command);
  });
}
