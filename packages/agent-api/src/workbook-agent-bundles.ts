import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  WorkbookCommentThreadSnapshot,
  LiteralInput,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookNoteSnapshot,
  WorkbookPivotSnapshot,
  WorkbookTableSnapshot,
} from "@bilig/protocol";
import {
  applyWorkbookAgentAnnotationCommand,
  deriveWorkbookAgentAnnotationCommandPreviewRanges,
  describeWorkbookAgentAnnotationCommand,
  estimateWorkbookAgentAnnotationCommandAffectedCells,
  isHighRiskWorkbookAgentAnnotationCommand,
  isWorkbookAgentAnnotationCommand,
  isWorkbookAgentAnnotationCommandValue,
  isWorkbookScopeAnnotationCommand,
} from "./workbook-agent-annotation-commands.js";
import {
  applyWorkbookAgentObjectCommand,
  deriveWorkbookAgentObjectCommandPreviewRanges,
  describeWorkbookAgentObjectCommand,
  estimateWorkbookAgentObjectCommandAffectedCells,
  isHighRiskWorkbookAgentObjectCommand,
  isWorkbookAgentObjectCommand,
  isWorkbookAgentObjectCommandValue,
  isWorkbookScopeObjectCommand,
} from "./workbook-agent-object-commands.js";
import {
  applyWorkbookAgentStructuralCommand,
  describeWorkbookAgentStructuralCommand,
  deriveWorkbookAgentStructuralCommandPreviewRanges,
  estimateWorkbookAgentStructuralCommandAffectedCells,
  isHighRiskWorkbookAgentStructuralCommand,
  isWorkbookAgentStructuralCommand,
  isWorkbookAgentStructuralCommandValue,
  isWorkbookScopeStructuralCommand,
} from "./workbook-agent-structural-commands.js";
import {
  applyWorkbookAgentValidationCommand,
  deriveWorkbookAgentValidationCommandPreviewRanges,
  describeWorkbookAgentValidationCommand,
  estimateWorkbookAgentValidationCommandAffectedCells,
  isHighRiskWorkbookAgentValidationCommand,
  isWorkbookAgentValidationCommand,
  isWorkbookAgentValidationCommandValue,
  isWorkbookScopeValidationCommand,
} from "./workbook-agent-validation-commands.js";

export interface WorkbookAgentUiSelectionRef {
  sheetName: string;
  address: string;
  range?: {
    startAddress: string;
    endAddress: string;
  };
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
      kind: "setRangeFormulas";
      range: CellRangeRef;
      formulas: string[][];
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
      kind: "upsertDefinedName";
      name: string;
      value: WorkbookDefinedNameValueSnapshot;
    }
  | {
      kind: "deleteDefinedName";
      name: string;
    }
  | {
      kind: "upsertTable";
      table: WorkbookTableSnapshot;
    }
  | {
      kind: "deleteTable";
      name: string;
    }
  | {
      kind: "upsertPivotTable";
      pivot: WorkbookPivotSnapshot;
    }
  | {
      kind: "deletePivotTable";
      sheetName: string;
      address: string;
    }
  | {
      kind: "createSheet";
      name: string;
    }
  | {
      kind: "renameSheet";
      currentName: string;
      nextName: string;
    }
  | {
      kind: "deleteSheet";
      name: string;
    }
  | {
      kind: "insertRows";
      sheetName: string;
      start: number;
      count: number;
    }
  | {
      kind: "deleteRows";
      sheetName: string;
      start: number;
      count: number;
    }
  | {
      kind: "insertColumns";
      sheetName: string;
      start: number;
      count: number;
    }
  | {
      kind: "deleteColumns";
      sheetName: string;
      start: number;
      count: number;
    }
  | {
      kind: "setFreezePane";
      sheetName: string;
      rows: number;
      cols: number;
    }
  | {
      kind: "setFilter";
      range: CellRangeRef;
    }
  | {
      kind: "clearFilter";
      range: CellRangeRef;
    }
  | {
      kind: "setSort";
      range: CellRangeRef;
      keys: {
        keyAddress: string;
        direction: "asc" | "desc";
      }[];
    }
  | {
      kind: "clearSort";
      range: CellRangeRef;
    }
  | {
      kind: "setDataValidation";
      validation: WorkbookDataValidationSnapshot;
    }
  | {
      kind: "clearDataValidation";
      range: CellRangeRef;
    }
  | {
      kind: "upsertCommentThread";
      thread: WorkbookCommentThreadSnapshot;
    }
  | {
      kind: "deleteCommentThread";
      sheetName: string;
      address: string;
    }
  | {
      kind: "upsertNote";
      note: WorkbookNoteSnapshot;
    }
  | {
      kind: "deleteNote";
      sheetName: string;
      address: string;
    }
  | {
      kind: "updateRowMetadata";
      sheetName: string;
      startRow: number;
      count: number;
      height?: number | null;
      hidden?: boolean | null;
    }
  | {
      kind: "updateColumnMetadata";
      sheetName: string;
      startCol: number;
      count: number;
      width?: number | null;
      hidden?: boolean | null;
    };

export type WorkbookAgentRiskClass = "low" | "medium" | "high";
export type WorkbookAgentBundleScope = "selection" | "sheet" | "workbook";
export type WorkbookAgentAppliedBy = "user" | "auto";
export type WorkbookAgentAcceptedScope = "full" | "partial";
export type WorkbookAgentSharedReviewStatus = "pending" | "approved" | "rejected";
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
  baseRevision: number;
  createdAtUnixMs: number;
  context: WorkbookAgentContextRef | null;
  commands: WorkbookAgentCommand[];
  affectedRanges: WorkbookAgentPreviewRange[];
  estimatedAffectedCells: number | null;
  sharedReview?: WorkbookAgentSharedReviewState | null;
}

export interface WorkbookAgentSharedReviewState {
  ownerUserId: string;
  status: WorkbookAgentSharedReviewStatus;
  decidedByUserId: string | null;
  decidedAtUnixMs: number | null;
  recommendations: WorkbookAgentSharedReviewRecommendation[];
}

export interface WorkbookAgentSharedReviewRecommendation {
  userId: string;
  decision: Extract<WorkbookAgentSharedReviewStatus, "approved" | "rejected">;
  decidedAtUnixMs: number;
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

function isAppliedBy(value: unknown): value is WorkbookAgentAppliedBy {
  return value === "user" || value === "auto";
}

function isSharedReviewStatus(value: unknown): value is WorkbookAgentSharedReviewStatus {
  return value === "pending" || value === "approved" || value === "rejected";
}

function isSharedReviewRecommendation(
  value: unknown,
): value is WorkbookAgentSharedReviewRecommendation {
  return (
    isRecord(value) &&
    typeof value["userId"] === "string" &&
    (value["decision"] === "approved" || value["decision"] === "rejected") &&
    typeof value["decidedAtUnixMs"] === "number"
  );
}

function isSharedReviewState(value: unknown): value is WorkbookAgentSharedReviewState {
  return (
    isRecord(value) &&
    typeof value["ownerUserId"] === "string" &&
    isSharedReviewStatus(value["status"]) &&
    (value["decidedByUserId"] === null || typeof value["decidedByUserId"] === "string") &&
    (value["decidedAtUnixMs"] === null || typeof value["decidedAtUnixMs"] === "number") &&
    Array.isArray(value["recommendations"]) &&
    value["recommendations"].every((entry) => isSharedReviewRecommendation(entry))
  );
}

function isAcceptedScope(value: unknown): value is WorkbookAgentAcceptedScope {
  return value === "full" || value === "partial";
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
    (value["selection"]["range"] === undefined ||
      (isRecord(value["selection"]["range"]) &&
        typeof value["selection"]["range"]["startAddress"] === "string" &&
        typeof value["selection"]["range"]["endAddress"] === "string")) &&
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

function rangeLabel(range: WorkbookAgentPreviewRange): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`;
}

export function describeWorkbookAgentCommand(command: WorkbookAgentCommand): string {
  if (isWorkbookAgentStructuralCommand(command)) {
    return describeWorkbookAgentStructuralCommand(command);
  }
  if (isWorkbookAgentObjectCommand(command)) {
    return describeWorkbookAgentObjectCommand(command);
  }
  if (isWorkbookAgentValidationCommand(command)) {
    return describeWorkbookAgentValidationCommand(command);
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    return describeWorkbookAgentAnnotationCommand(command);
  }
  switch (command.kind) {
    case "writeRange": {
      const ranges = deriveWorkbookAgentCommandPreviewRanges(command);
      return ranges[0] ? `Write cells in ${rangeLabel(ranges[0])}` : "Write cells";
    }
    case "setRangeFormulas":
      return `Set formulas in ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[0]!)}`;
    case "clearRange":
      return `Clear ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[0]!)}`;
    case "formatRange":
      return `Format ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[0]!)}`;
    case "fillRange":
      return `Fill ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[1]!)}`;
    case "copyRange":
      return `Copy into ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[1]!)}`;
    case "moveRange":
      return `Move cells to ${rangeLabel(deriveWorkbookAgentCommandPreviewRanges(command)[1]!)}`;
    default: {
      const exhaustive: never = command;
      return String(exhaustive);
    }
  }
}

function dedupePreviewRanges(
  ranges: readonly WorkbookAgentPreviewRange[],
): WorkbookAgentPreviewRange[] {
  const nextRanges: WorkbookAgentPreviewRange[] = [];
  ranges.forEach((range) => {
    if (!nextRanges.some((existing) => samePreviewRange(existing, range))) {
      nextRanges.push(range);
    }
  });
  return nextRanges;
}

function summarizeCommands(commands: readonly WorkbookAgentCommand[]): string {
  if (commands.length === 0) {
    return "No workbook changes staged";
  }
  if (commands.length === 1) {
    return describeWorkbookAgentCommand(commands[0]!);
  }
  const firstSummary = describeWorkbookAgentCommand(commands[0]!);
  return `${firstSummary} and ${String(commands.length - 1)} more change${
    commands.length === 2 ? "" : "s"
  }`;
}

function isSelectionOnlyCommand(
  command: WorkbookAgentCommand,
  context: WorkbookAgentContextRef | null,
): boolean {
  if (!context) {
    return false;
  }
  const selectionSheet = context.selection.sheetName;
  const selectionRange = context.selection.range ?? {
    startAddress: context.selection.address,
    endAddress: context.selection.address,
  };
  const ranges = deriveWorkbookAgentCommandPreviewRanges(command);
  if (ranges.length !== 1) {
    return false;
  }
  const range = ranges[0];
  if (!range) {
    return false;
  }
  return (
    range.role === "target" &&
    range.sheetName === selectionSheet &&
    range.startAddress === selectionRange.startAddress &&
    range.endAddress === selectionRange.endAddress
  );
}

function deriveWorkbookAgentRiskClass(
  commands: readonly WorkbookAgentCommand[],
  context: WorkbookAgentContextRef | null,
): WorkbookAgentRiskClass {
  if (
    commands.some(
      (command) =>
        (isWorkbookAgentStructuralCommand(command) &&
          isHighRiskWorkbookAgentStructuralCommand(command)) ||
        (isWorkbookAgentObjectCommand(command) && isHighRiskWorkbookAgentObjectCommand(command)) ||
        (isWorkbookAgentValidationCommand(command) &&
          isHighRiskWorkbookAgentValidationCommand(command)) ||
        (isWorkbookAgentAnnotationCommand(command) &&
          isHighRiskWorkbookAgentAnnotationCommand(command)),
    )
  ) {
    return "high";
  }
  if (
    commands.every(
      (command) => command.kind === "formatRange" && isSelectionOnlyCommand(command, context),
    )
  ) {
    return "low";
  }
  return "medium";
}

function deriveWorkbookAgentBundleScope(
  commands: readonly WorkbookAgentCommand[],
  context: WorkbookAgentContextRef | null,
): WorkbookAgentBundleScope {
  if (
    commands.some(
      (command) =>
        (isWorkbookAgentStructuralCommand(command) && isWorkbookScopeStructuralCommand(command)) ||
        (isWorkbookAgentObjectCommand(command) && isWorkbookScopeObjectCommand(command)) ||
        (isWorkbookAgentValidationCommand(command) && isWorkbookScopeValidationCommand(command)) ||
        (isWorkbookAgentAnnotationCommand(command) && isWorkbookScopeAnnotationCommand(command)),
    )
  ) {
    return "workbook";
  }

  const ranges = commands.flatMap((command) => deriveWorkbookAgentCommandPreviewRanges(command));
  if (ranges.length === 0) {
    return "sheet";
  }

  const distinctSheets = new Set(ranges.map((range) => range.sheetName));
  if (distinctSheets.size > 1) {
    return "workbook";
  }

  if (context && commands.every((command) => isSelectionOnlyCommand(command, context))) {
    return "selection";
  }

  return "sheet";
}

function estimateWorkbookAgentAffectedCells(
  commands: readonly WorkbookAgentCommand[],
): number | null {
  let total = 0;
  let sawCount = false;
  commands.forEach((command) => {
    const next = estimateWorkbookAgentCommandAffectedCells(command);
    if (next !== null) {
      total += next;
      sawCount = true;
    }
  });
  return sawCount ? total : null;
}

export function createWorkbookAgentCommandBundle(input: {
  bundleId?: string;
  documentId: string;
  threadId: string;
  turnId: string;
  goalText: string;
  baseRevision: number;
  context: WorkbookAgentContextRef | null;
  commands: readonly WorkbookAgentCommand[];
  now: number;
  sharedReview?: WorkbookAgentSharedReviewState | null;
}): WorkbookAgentCommandBundle {
  const commands = [...input.commands];
  const scope = deriveWorkbookAgentBundleScope(commands, input.context);
  const riskClass = deriveWorkbookAgentRiskClass(commands, input.context);
  return {
    id: input.bundleId ?? crypto.randomUUID(),
    documentId: input.documentId,
    threadId: input.threadId,
    turnId: input.turnId,
    goalText: input.goalText,
    summary: summarizeCommands(commands),
    scope,
    riskClass,
    baseRevision: input.baseRevision,
    createdAtUnixMs: input.now,
    context: input.context ? structuredClone(input.context) : null,
    commands: commands.map((command) => structuredClone(command)),
    affectedRanges: dedupePreviewRanges(
      commands.flatMap((command) => deriveWorkbookAgentCommandPreviewRanges(command)),
    ),
    estimatedAffectedCells: estimateWorkbookAgentAffectedCells(commands),
    sharedReview: input.sharedReview ? structuredClone(input.sharedReview) : null,
  };
}

export function appendWorkbookAgentCommandToBundle(input: {
  previousBundle: WorkbookAgentCommandBundle | null;
  documentId: string;
  threadId: string;
  turnId: string;
  goalText: string;
  baseRevision: number;
  context: WorkbookAgentContextRef | null;
  command: WorkbookAgentCommand;
  now: number;
}): WorkbookAgentCommandBundle {
  const previousBundle =
    input.previousBundle &&
    input.previousBundle.threadId === input.threadId &&
    input.previousBundle.turnId === input.turnId
      ? input.previousBundle
      : null;
  return createWorkbookAgentCommandBundle({
    ...(previousBundle ? { bundleId: previousBundle.id } : {}),
    documentId: input.documentId,
    threadId: input.threadId,
    turnId: input.turnId,
    goalText: input.goalText,
    baseRevision: input.baseRevision,
    context: input.context,
    commands: [...(previousBundle?.commands ?? []), input.command],
    now: previousBundle ? previousBundle.createdAtUnixMs : input.now,
    sharedReview: previousBundle?.sharedReview ?? null,
  });
}

export function describeWorkbookAgentBundle(bundle: WorkbookAgentCommandBundle): string {
  const affectedCells =
    bundle.estimatedAffectedCells === null
      ? "unknown affected cell count"
      : `${String(bundle.estimatedAffectedCells)} affected cell${
          bundle.estimatedAffectedCells === 1 ? "" : "s"
        }`;
  return [
    `Prepared workbook review item: ${bundle.summary}.`,
    `Risk: ${bundle.riskClass}. Scope: ${bundle.scope}.`,
    `Change target: ${affectedCells}.`,
  ].join(" ");
}

export function normalizeWorkbookAgentCommandIndexes(
  bundle: Pick<WorkbookAgentCommandBundle, "commands">,
  commandIndexes: readonly number[] | null | undefined,
): number[] {
  if (commandIndexes === null || commandIndexes === undefined) {
    return bundle.commands.map((_command, index) => index);
  }
  const requested = new Set<number>();
  commandIndexes.forEach((index) => {
    if (Number.isInteger(index) && index >= 0 && index < bundle.commands.length) {
      requested.add(index);
    }
  });
  return bundle.commands.flatMap((_command, index) => (requested.has(index) ? [index] : []));
}

export function isFullWorkbookAgentCommandSelection(input: {
  bundle: Pick<WorkbookAgentCommandBundle, "commands">;
  commandIndexes: readonly number[] | null | undefined;
}): boolean {
  return (
    normalizeWorkbookAgentCommandIndexes(input.bundle, input.commandIndexes).length ===
    input.bundle.commands.length
  );
}

export function projectWorkbookAgentBundle(input: {
  bundle: WorkbookAgentCommandBundle;
  commandIndexes: readonly number[] | null | undefined;
  bundleId?: string;
  baseRevision?: number;
  now?: number;
}): WorkbookAgentCommandBundle | null {
  const selectedIndexes = normalizeWorkbookAgentCommandIndexes(input.bundle, input.commandIndexes);
  if (selectedIndexes.length === 0) {
    return null;
  }
  if (selectedIndexes.length === input.bundle.commands.length) {
    return {
      ...structuredClone(input.bundle),
      ...(input.bundleId ? { id: input.bundleId } : {}),
      ...(input.baseRevision !== undefined ? { baseRevision: input.baseRevision } : {}),
      ...(input.now !== undefined ? { createdAtUnixMs: input.now } : {}),
    };
  }
  return createWorkbookAgentCommandBundle({
    ...(input.bundleId ? { bundleId: input.bundleId } : {}),
    documentId: input.bundle.documentId,
    threadId: input.bundle.threadId,
    turnId: input.bundle.turnId,
    goalText: input.bundle.goalText,
    baseRevision: input.baseRevision ?? input.bundle.baseRevision,
    context: input.bundle.context,
    commands: selectedIndexes.map((index) => input.bundle.commands[index]!),
    now: input.now ?? input.bundle.createdAtUnixMs,
    sharedReview: input.bundle.sharedReview ?? null,
  });
}

export function splitWorkbookAgentCommandBundle(input: {
  bundle: WorkbookAgentCommandBundle;
  acceptedCommandIndexes: readonly number[] | null | undefined;
  remainingBaseRevision?: number;
  remainingBundleId?: string;
  now?: number;
}): {
  acceptedBundle: WorkbookAgentCommandBundle | null;
  remainingBundle: WorkbookAgentCommandBundle | null;
  acceptedScope: WorkbookAgentAcceptedScope | null;
  acceptedCommandIndexes: number[];
} {
  const acceptedCommandIndexes = normalizeWorkbookAgentCommandIndexes(
    input.bundle,
    input.acceptedCommandIndexes,
  );
  if (acceptedCommandIndexes.length === 0) {
    return {
      acceptedBundle: null,
      remainingBundle: structuredClone(input.bundle),
      acceptedScope: null,
      acceptedCommandIndexes,
    };
  }
  const acceptedSet = new Set(acceptedCommandIndexes);
  const remainingCommandIndexes = input.bundle.commands.flatMap((_command, index) =>
    acceptedSet.has(index) ? [] : [index],
  );
  return {
    acceptedBundle: projectWorkbookAgentBundle({
      bundle: input.bundle,
      commandIndexes: acceptedCommandIndexes,
      bundleId: input.bundle.id,
    }),
    remainingBundle: projectWorkbookAgentBundle({
      bundle: input.bundle,
      commandIndexes: remainingCommandIndexes,
      ...(input.remainingBundleId ? { bundleId: input.remainingBundleId } : {}),
      ...(input.remainingBaseRevision !== undefined
        ? { baseRevision: input.remainingBaseRevision }
        : {}),
      ...(input.now !== undefined ? { now: input.now } : {}),
    }),
    acceptedScope:
      acceptedCommandIndexes.length === input.bundle.commands.length ? "full" : "partial",
    acceptedCommandIndexes,
  };
}

export function buildWorkbookAgentExecutionRecord(input: {
  bundle: WorkbookAgentCommandBundle;
  actorUserId: string;
  planText: string | null;
  preview: WorkbookAgentExecutionRecord["preview"];
  appliedRevision: number;
  appliedBy: WorkbookAgentAppliedBy;
  acceptedScope: WorkbookAgentAcceptedScope;
  now: number;
}): WorkbookAgentExecutionRecord {
  return {
    id: crypto.randomUUID(),
    bundleId: input.bundle.id,
    documentId: input.bundle.documentId,
    threadId: input.bundle.threadId,
    turnId: input.bundle.turnId,
    actorUserId: input.actorUserId,
    goalText: input.bundle.goalText,
    planText: input.planText,
    summary: input.bundle.summary,
    scope: input.bundle.scope,
    riskClass: input.bundle.riskClass,
    acceptedScope: input.acceptedScope,
    appliedBy: input.appliedBy,
    baseRevision: input.bundle.baseRevision,
    appliedRevision: input.appliedRevision,
    createdAtUnixMs: input.bundle.createdAtUnixMs,
    appliedAtUnixMs: input.now,
    context: input.bundle.context ? structuredClone(input.bundle.context) : null,
    commands: input.bundle.commands.map((command) => structuredClone(command)),
    preview: input.preview ? structuredClone(input.preview) : null,
  };
}

export function isWorkbookAgentCommand(value: unknown): value is WorkbookAgentCommand {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  if (isWorkbookAgentStructuralCommandValue(value)) {
    return true;
  }
  if (isWorkbookAgentObjectCommandValue(value)) {
    return true;
  }
  if (isWorkbookAgentValidationCommandValue(value)) {
    return true;
  }
  if (isWorkbookAgentAnnotationCommandValue(value)) {
    return true;
  }
  switch (value["kind"]) {
    case "writeRange":
    case "setRangeFormulas":
      return value["kind"] === "writeRange"
        ? typeof value["sheetName"] === "string" &&
            typeof value["startAddress"] === "string" &&
            Array.isArray(value["values"]) &&
            value["values"].every(
              (row) =>
                Array.isArray(row) &&
                row.length > 0 &&
                row.every((cellValue) => isWriteCellInput(cellValue)),
            )
        : isCellRangeRef(value["range"]) &&
            Array.isArray(value["formulas"]) &&
            value["formulas"].every(
              (row) =>
                Array.isArray(row) &&
                row.length > 0 &&
                row.every((formula) => typeof formula === "string" && formula.trim().length > 0),
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
    typeof value["baseRevision"] === "number" &&
    typeof value["createdAtUnixMs"] === "number" &&
    (value["context"] === null || isWorkbookAgentContextRef(value["context"])) &&
    isCommandArray(value["commands"]) &&
    Array.isArray(value["affectedRanges"]) &&
    value["affectedRanges"].every((entry) => isWorkbookAgentPreviewRange(entry)) &&
    (value["sharedReview"] === undefined ||
      value["sharedReview"] === null ||
      isSharedReviewState(value["sharedReview"])) &&
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
  if (isWorkbookAgentStructuralCommand(command)) {
    return estimateWorkbookAgentStructuralCommandAffectedCells(command);
  }
  if (isWorkbookAgentObjectCommand(command)) {
    return estimateWorkbookAgentObjectCommandAffectedCells(command);
  }
  if (isWorkbookAgentValidationCommand(command)) {
    return estimateWorkbookAgentValidationCommandAffectedCells(command);
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    return estimateWorkbookAgentAnnotationCommandAffectedCells(command);
  }
  switch (command.kind) {
    case "writeRange":
      return command.values.reduce((sum, row) => sum + row.length, 0);
    case "setRangeFormulas":
      return command.formulas.reduce((sum, row) => sum + row.length, 0);
    case "clearRange":
    case "formatRange":
      return countRangeCells(command.range);
    case "fillRange":
    case "copyRange":
    case "moveRange":
      return countRangeCells(command.target);
    default: {
      const exhaustive: never = command;
      return exhaustive;
    }
  }
}

export function deriveWorkbookAgentCommandPreviewRanges(
  command: WorkbookAgentCommand,
): WorkbookAgentPreviewRange[] {
  if (isWorkbookAgentStructuralCommand(command)) {
    return deriveWorkbookAgentStructuralCommandPreviewRanges(command);
  }
  if (isWorkbookAgentObjectCommand(command)) {
    return deriveWorkbookAgentObjectCommandPreviewRanges(command);
  }
  if (isWorkbookAgentValidationCommand(command)) {
    return deriveWorkbookAgentValidationCommandPreviewRanges(command);
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    return deriveWorkbookAgentAnnotationCommandPreviewRanges(command);
  }
  switch (command.kind) {
    case "writeRange":
      return [
        {
          ...writeRangeToRange(command),
          role: "target",
        },
      ];
    case "setRangeFormulas":
      return [
        {
          ...normalizeRangeBounds(command.range),
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
  if (isWorkbookAgentStructuralCommand(command)) {
    applyWorkbookAgentStructuralCommand(engine, command);
    return;
  }
  if (isWorkbookAgentObjectCommand(command)) {
    applyWorkbookAgentObjectCommand(engine, command);
    return;
  }
  if (isWorkbookAgentValidationCommand(command)) {
    applyWorkbookAgentValidationCommand(engine, command);
    return;
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    applyWorkbookAgentAnnotationCommand(engine, command);
    return;
  }
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
    case "setRangeFormulas":
      engine.setRangeFormulas(
        command.range,
        command.formulas.map((row) => row.map((formula) => normalizeFormula(formula))),
      );
      return;
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
