import { isWorkbookChangeUndoBundle, type WorkbookChangeUndoBundle } from "@bilig/zero-sync";
import { formatWorkbookCollaboratorLabel } from "./workbook-presence-model.js";

export interface WorkbookChangeRange {
  readonly sheetName: string;
  readonly startAddress: string;
  readonly endAddress: string;
}

export interface WorkbookChangeRow {
  readonly revision: number;
  readonly actorUserId: string;
  readonly clientMutationId: string | null;
  readonly eventKind: string;
  readonly summary: string;
  readonly sheetId: number | null;
  readonly sheetName: string | null;
  readonly anchorAddress: string | null;
  readonly rangeJson: WorkbookChangeRange | null;
  readonly undoBundleJson: WorkbookChangeUndoBundle | null;
  readonly revertedByRevision: number | null;
  readonly revertsRevision: number | null;
  readonly createdAt: number;
}

export interface WorkbookChangeEntry {
  readonly revision: number;
  readonly actorUserId: string;
  readonly actorLabel: string;
  readonly clientMutationId: string | null;
  readonly eventKind: string;
  readonly summary: string;
  readonly sheetName: string | null;
  readonly address: string | null;
  readonly targetLabel: string | null;
  readonly createdAt: number;
  readonly isJumpable: boolean;
  readonly canRevert: boolean;
  readonly canRedo: boolean;
  readonly revertedByRevision: number | null;
  readonly revertsRevision: number | null;
}

export interface WorkbookHistoryState {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly undoRevision: number | null;
  readonly redoRevision: number | null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function normalizeWorkbookChangeRange(value: unknown): WorkbookChangeRange | null {
  if (!isRecord(value)) {
    return null;
  }
  const sheetName = value["sheetName"];
  const startAddress = value["startAddress"];
  const endAddress = value["endAddress"];
  if (
    typeof sheetName !== "string" ||
    typeof startAddress !== "string" ||
    typeof endAddress !== "string"
  ) {
    return null;
  }
  return {
    sheetName,
    startAddress,
    endAddress,
  };
}

function normalizeWorkbookChangeRow(value: unknown): WorkbookChangeRow | null {
  if (!isRecord(value)) {
    return null;
  }
  const revision = value["revision"];
  const actorUserId = value["actorUserId"];
  const eventKind = value["eventKind"];
  const summary = value["summary"];
  const createdAt = value["createdAt"];
  if (
    typeof revision !== "number" ||
    typeof actorUserId !== "string" ||
    typeof eventKind !== "string" ||
    typeof summary !== "string" ||
    typeof createdAt !== "number"
  ) {
    return null;
  }
  const clientMutationId = value["clientMutationId"];
  const sheetId = value["sheetId"];
  const sheetName = value["sheetName"];
  const anchorAddress = value["anchorAddress"];
  const revertedByRevision = value["revertedByRevision"];
  const revertsRevision = value["revertsRevision"];
  return {
    revision,
    actorUserId,
    clientMutationId: typeof clientMutationId === "string" ? clientMutationId : null,
    eventKind,
    summary,
    sheetId: typeof sheetId === "number" ? sheetId : null,
    sheetName: typeof sheetName === "string" ? sheetName : null,
    anchorAddress: typeof anchorAddress === "string" ? anchorAddress : null,
    rangeJson: normalizeWorkbookChangeRange(value["rangeJson"]),
    undoBundleJson: isWorkbookChangeUndoBundle(value["undoBundleJson"])
      ? value["undoBundleJson"]
      : null,
    revertedByRevision: typeof revertedByRevision === "number" ? revertedByRevision : null,
    revertsRevision: typeof revertsRevision === "number" ? revertsRevision : null,
    createdAt,
  };
}

export function normalizeWorkbookChangeRows(value: unknown): readonly WorkbookChangeRow[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.flatMap((entry) => {
    const row = normalizeWorkbookChangeRow(entry);
    return row ? [row] : [];
  });
}

function formatChangeTarget(
  range: WorkbookChangeRange | null,
  fallbackAddress: string | null,
): string | null {
  if (range) {
    return range.startAddress === range.endAddress
      ? `${range.sheetName}!${range.startAddress}`
      : `${range.sheetName}!${range.startAddress}:${range.endAddress}`;
  }
  return fallbackAddress ? fallbackAddress : null;
}

export function formatWorkbookChangeTimestamp(createdAt: number): string {
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(createdAt));
}

export function selectWorkbookChangeEntries(input: {
  readonly rows: readonly WorkbookChangeRow[];
  readonly knownSheetNames: readonly string[];
}): readonly WorkbookChangeEntry[] {
  const knownSheetNames = new Set(input.knownSheetNames);
  return input.rows.map((row) => {
    const targetLabel = formatChangeTarget(row.rangeJson, row.anchorAddress);
    const sheetName = row.sheetName ?? row.rangeJson?.sheetName ?? null;
    const address = row.anchorAddress ?? row.rangeJson?.startAddress ?? null;
    return {
      revision: row.revision,
      actorUserId: row.actorUserId,
      actorLabel: formatWorkbookCollaboratorLabel(row.actorUserId),
      clientMutationId: row.clientMutationId,
      eventKind: row.eventKind,
      summary: row.summary,
      sheetName,
      address,
      targetLabel,
      createdAt: row.createdAt,
      isJumpable:
        typeof sheetName === "string" &&
        typeof address === "string" &&
        knownSheetNames.has(sheetName),
      canRevert:
        row.undoBundleJson !== null &&
        row.revertedByRevision === null &&
        row.eventKind !== "revertChange",
      canRedo:
        row.undoBundleJson !== null &&
        row.revertedByRevision === null &&
        row.eventKind === "revertChange",
      revertedByRevision: row.revertedByRevision,
      revertsRevision: row.revertsRevision,
    } satisfies WorkbookChangeEntry;
  });
}

export function selectWorkbookHistoryState(input: {
  readonly entries: readonly WorkbookChangeEntry[];
  readonly currentUserId: string;
}): WorkbookHistoryState {
  const ownEntries = input.entries.filter((entry) => entry.actorUserId === input.currentUserId);
  const latestOwnHistoryAction = ownEntries.find(
    (entry) => entry.canRedo || entry.canRevert || entry.eventKind === "redoChange",
  );
  const latestUndoableChange = ownEntries.find((entry) => entry.canRevert);

  return {
    canUndo: latestUndoableChange !== undefined,
    canRedo: latestOwnHistoryAction?.canRedo === true,
    undoRevision: latestUndoableChange?.revision ?? null,
    redoRevision: latestOwnHistoryAction?.canRedo ? latestOwnHistoryAction.revision : null,
  };
}
