import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
  WorkbookAxisEntrySnapshot,
} from "@bilig/protocol";

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
export type WorkbookAgentApprovalMode = "auto" | "preview" | "explicit";
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
  approvalMode: WorkbookAgentApprovalMode;
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

function hasOwnProperty(value: object, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(value, key);
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

function resolveRowMetadataCommandState(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookAgentCommand, { kind: "updateRowMetadata" }>,
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

function resolveColumnMetadataCommandState(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookAgentCommand, { kind: "updateColumnMetadata" }>,
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

export function describeWorkbookAgentCommand(command: WorkbookAgentCommand): string {
  switch (command.kind) {
    case "writeRange": {
      const ranges = deriveWorkbookAgentCommandPreviewRanges(command);
      return ranges[0] ? `Write cells in ${rangeLabel(ranges[0])}` : "Write cells";
    }
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
    case "createSheet":
      return `Create sheet ${command.name}`;
    case "renameSheet":
      return `Rename sheet ${command.currentName} to ${command.nextName}`;
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
    commands.some((command) => command.kind === "createSheet" || command.kind === "renameSheet")
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
    commands.some((command) => command.kind === "createSheet" || command.kind === "renameSheet")
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

function deriveWorkbookAgentApprovalMode(
  riskClass: WorkbookAgentRiskClass,
  scope: WorkbookAgentBundleScope,
): WorkbookAgentApprovalMode {
  if (riskClass === "low" && scope === "selection") {
    return "auto";
  }
  if (riskClass === "high") {
    return "explicit";
  }
  return "preview";
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
    approvalMode: deriveWorkbookAgentApprovalMode(riskClass, scope),
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
    `Impact: ${affectedCells}.`,
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
    approvalMode: input.bundle.approvalMode,
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
    case "updateRowMetadata":
    case "updateColumnMetadata":
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
