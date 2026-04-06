import {
  deriveWorkbookAgentCommandPreviewRanges,
  estimateWorkbookAgentCommandAffectedCells,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentContextRef,
  type WorkbookAgentExecutionRecord,
  type WorkbookAgentAppliedBy,
  type WorkbookAgentPreviewRange,
  type WorkbookAgentRiskClass,
  type WorkbookAgentApprovalMode,
  type WorkbookAgentBundleScope,
} from "@bilig/agent-api";
import type { WorkbookAgentUiContext } from "@bilig/contracts";

function rangeLabel(range: WorkbookAgentPreviewRange): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`;
}

function describeCommand(command: WorkbookAgentCommand): string {
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
    default: {
      const exhaustive: never = command;
      return String(exhaustive);
    }
  }
}

function sameRange(left: WorkbookAgentPreviewRange, right: WorkbookAgentPreviewRange): boolean {
  return (
    left.sheetName === right.sheetName &&
    left.startAddress === right.startAddress &&
    left.endAddress === right.endAddress &&
    left.role === right.role
  );
}

function dedupeRanges(ranges: readonly WorkbookAgentPreviewRange[]): WorkbookAgentPreviewRange[] {
  const nextRanges: WorkbookAgentPreviewRange[] = [];
  ranges.forEach((range) => {
    if (!nextRanges.some((existing) => sameRange(existing, range))) {
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
    return describeCommand(commands[0]!);
  }
  const firstSummary = describeCommand(commands[0]!);
  return `${firstSummary} and ${String(commands.length - 1)} more change${
    commands.length === 2 ? "" : "s"
  }`;
}

function toContextRef(context: WorkbookAgentUiContext | null): WorkbookAgentContextRef | null {
  return context
    ? {
        selection: {
          sheetName: context.selection.sheetName,
          address: context.selection.address,
        },
        viewport: { ...context.viewport },
      }
    : null;
}

function isSelectionOnlyCommand(
  command: WorkbookAgentCommand,
  context: WorkbookAgentContextRef | null,
): boolean {
  if (!context) {
    return false;
  }
  const selectionSheet = context.selection.sheetName;
  const selectionAddress = context.selection.address;
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
    range.startAddress === selectionAddress &&
    range.endAddress === selectionAddress
  );
}

function deriveRiskClass(
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

function deriveScope(
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

function deriveApprovalMode(
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

function estimateAffectedCells(commands: readonly WorkbookAgentCommand[]): number | null {
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

export function createWorkbookAgentBundle(input: {
  bundleId?: string;
  documentId: string;
  threadId: string;
  turnId: string;
  goalText: string;
  baseRevision: number;
  context: WorkbookAgentUiContext | null;
  commands: readonly WorkbookAgentCommand[];
  now: number;
}): WorkbookAgentCommandBundle {
  const context = toContextRef(input.context);
  const commands = [...input.commands];
  const scope = deriveScope(commands, context);
  const riskClass = deriveRiskClass(commands, context);
  const affectedRanges = dedupeRanges(
    commands.flatMap((command) => deriveWorkbookAgentCommandPreviewRanges(command)),
  );
  return {
    id: input.bundleId ?? crypto.randomUUID(),
    documentId: input.documentId,
    threadId: input.threadId,
    turnId: input.turnId,
    goalText: input.goalText,
    summary: summarizeCommands(commands),
    scope,
    riskClass,
    approvalMode: deriveApprovalMode(riskClass, scope),
    baseRevision: input.baseRevision,
    createdAtUnixMs: input.now,
    context,
    commands,
    affectedRanges,
    estimatedAffectedCells: estimateAffectedCells(commands),
  };
}

export function appendWorkbookAgentBundleCommand(input: {
  previousBundle: WorkbookAgentCommandBundle | null;
  documentId: string;
  threadId: string;
  turnId: string;
  goalText: string;
  baseRevision: number;
  context: WorkbookAgentUiContext | null;
  command: WorkbookAgentCommand;
  now: number;
}): WorkbookAgentCommandBundle {
  const previousBundle =
    input.previousBundle &&
    input.previousBundle.threadId === input.threadId &&
    input.previousBundle.turnId === input.turnId
      ? input.previousBundle
      : null;
  const previousCommands = previousBundle ? previousBundle.commands : [];
  return createWorkbookAgentBundle({
    ...(previousBundle ? { bundleId: previousBundle.id } : {}),
    documentId: input.documentId,
    threadId: input.threadId,
    turnId: input.turnId,
    goalText: input.goalText,
    baseRevision: input.baseRevision,
    context: input.context,
    commands: [...previousCommands, input.command],
    now: previousBundle ? previousBundle.createdAtUnixMs : input.now,
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
    `Staged preview bundle: ${bundle.summary}.`,
    `Risk: ${bundle.riskClass}. Scope: ${bundle.scope}.`,
    `Preview target: ${affectedCells}.`,
  ].join(" ");
}

export function buildWorkbookAgentExecutionRecord(input: {
  bundle: WorkbookAgentCommandBundle;
  actorUserId: string;
  planText: string | null;
  preview: WorkbookAgentExecutionRecord["preview"];
  appliedRevision: number;
  appliedBy: WorkbookAgentAppliedBy;
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
    acceptedScope: "full",
    appliedBy: input.appliedBy,
    baseRevision: input.bundle.baseRevision,
    appliedRevision: input.appliedRevision,
    createdAtUnixMs: input.bundle.createdAtUnixMs,
    appliedAtUnixMs: input.now,
    context: input.bundle.context,
    commands: input.bundle.commands.map((command) => structuredClone(command)),
    preview: input.preview ? structuredClone(input.preview) : null,
  };
}
