import type { Viewport } from "@bilig/protocol";
import { formatWorkbookCollaboratorLabel } from "./workbook-presence-model.js";

export type WorkbookSheetViewVisibility = "private" | "shared";

export interface WorkbookSheetViewRow {
  readonly id: string;
  readonly ownerUserId: string;
  readonly name: string;
  readonly visibility: WorkbookSheetViewVisibility;
  readonly sheetId: number | null;
  readonly sheetName: string | null;
  readonly address: string;
  readonly viewportJson: unknown;
  readonly updatedAt: number;
}

export interface WorkbookSheetViewEntry {
  readonly id: string;
  readonly ownerUserId: string;
  readonly ownerLabel: string;
  readonly name: string;
  readonly visibility: WorkbookSheetViewVisibility;
  readonly sheetId: number | null;
  readonly sheetName: string | null;
  readonly address: string;
  readonly viewport: Viewport;
  readonly updatedAt: number;
  readonly targetLabel: string;
  readonly isApplyable: boolean;
  readonly canManage: boolean;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function normalizeViewport(value: unknown): Viewport | null {
  if (!isRecord(value)) {
    return null;
  }
  const rowStart = value["rowStart"];
  const rowEnd = value["rowEnd"];
  const colStart = value["colStart"];
  const colEnd = value["colEnd"];
  if (
    typeof rowStart !== "number" ||
    typeof rowEnd !== "number" ||
    typeof colStart !== "number" ||
    typeof colEnd !== "number" ||
    rowEnd < rowStart ||
    colEnd < colStart
  ) {
    return null;
  }
  return {
    rowStart,
    rowEnd,
    colStart,
    colEnd,
  };
}

function normalizeVisibility(value: unknown): WorkbookSheetViewVisibility | null {
  return value === "private" || value === "shared" ? value : null;
}

function normalizeWorkbookSheetViewRow(value: unknown): WorkbookSheetViewRow | null {
  if (!isRecord(value)) {
    return null;
  }
  const id = value["id"];
  const ownerUserId = value["ownerUserId"];
  const name = value["name"];
  const visibility = normalizeVisibility(value["visibility"]);
  const address = value["address"];
  const updatedAt = value["updatedAt"];
  if (
    typeof id !== "string" ||
    typeof ownerUserId !== "string" ||
    typeof name !== "string" ||
    visibility === null ||
    typeof address !== "string" ||
    typeof updatedAt !== "number"
  ) {
    return null;
  }
  const sheetId = value["sheetId"];
  const sheetName = value["sheetName"];
  return {
    id,
    ownerUserId,
    name,
    visibility,
    sheetId: typeof sheetId === "number" ? sheetId : null,
    sheetName: typeof sheetName === "string" ? sheetName : null,
    address,
    viewportJson: value["viewportJson"],
    updatedAt,
  };
}

export function normalizeWorkbookSheetViewRows(value: unknown): readonly WorkbookSheetViewRow[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.flatMap((entry) => {
    const row = normalizeWorkbookSheetViewRow(entry);
    return row ? [row] : [];
  });
}

export function selectWorkbookSheetViewEntries(input: {
  readonly rows: readonly WorkbookSheetViewRow[];
  readonly knownSheetNames: readonly string[];
  readonly currentUserId: string;
}): readonly WorkbookSheetViewEntry[] {
  const knownSheets = new Set(input.knownSheetNames);
  return input.rows
    .flatMap((row) => {
      const viewport = normalizeViewport(row.viewportJson);
      if (!viewport) {
        return [];
      }
      const ownerLabel = formatWorkbookCollaboratorLabel(row.ownerUserId);
      const isApplyable = typeof row.sheetName === "string" && knownSheets.has(row.sheetName);
      return [
        {
          id: row.id,
          ownerUserId: row.ownerUserId,
          ownerLabel,
          name: row.name,
          visibility: row.visibility,
          sheetId: row.sheetId,
          sheetName: row.sheetName,
          address: row.address,
          viewport,
          updatedAt: row.updatedAt,
          targetLabel:
            typeof row.sheetName === "string" ? `${row.sheetName}!${row.address}` : "Unknown sheet",
          isApplyable,
          canManage: row.ownerUserId === input.currentUserId,
        } satisfies WorkbookSheetViewEntry,
      ];
    })
    .toSorted((left, right) => {
      if (left.visibility !== right.visibility) {
        return left.visibility === "private" ? -1 : 1;
      }
      if (left.updatedAt !== right.updatedAt) {
        return right.updatedAt - left.updatedAt;
      }
      return left.name.localeCompare(right.name);
    });
}

export function formatWorkbookSheetViewTimestamp(updatedAt: number): string {
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(updatedAt);
}
