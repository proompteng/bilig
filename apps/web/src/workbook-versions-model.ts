import type { Viewport } from "@bilig/protocol";
import { formatWorkbookCollaboratorLabel } from "./workbook-presence-model.js";

export interface WorkbookVersionRow {
  readonly id: string;
  readonly ownerUserId: string;
  readonly name: string;
  readonly revision: number;
  readonly sheetId: number | null;
  readonly sheetName: string | null;
  readonly address: string | null;
  readonly viewportJson: unknown;
  readonly createdAt: number;
  readonly updatedAt: number;
}

export interface WorkbookVersionEntry {
  readonly id: string;
  readonly ownerUserId: string;
  readonly ownerLabel: string;
  readonly name: string;
  readonly revision: number;
  readonly sheetId: number | null;
  readonly sheetName: string | null;
  readonly address: string | null;
  readonly viewport: Viewport | null;
  readonly createdAt: number;
  readonly updatedAt: number;
  readonly targetLabel: string | null;
  readonly canDelete: boolean;
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

function normalizeWorkbookVersionRow(value: unknown): WorkbookVersionRow | null {
  if (!isRecord(value)) {
    return null;
  }
  const id = value["id"];
  const ownerUserId = value["ownerUserId"];
  const name = value["name"];
  const revision = value["revision"];
  const createdAt = value["createdAt"];
  const updatedAt = value["updatedAt"];
  if (
    typeof id !== "string" ||
    typeof ownerUserId !== "string" ||
    typeof name !== "string" ||
    typeof revision !== "number" ||
    typeof createdAt !== "number" ||
    typeof updatedAt !== "number"
  ) {
    return null;
  }
  const sheetId = value["sheetId"];
  const sheetName = value["sheetName"];
  const address = value["address"];
  return {
    id,
    ownerUserId,
    name,
    revision,
    sheetId: typeof sheetId === "number" ? sheetId : null,
    sheetName: typeof sheetName === "string" ? sheetName : null,
    address: typeof address === "string" ? address : null,
    viewportJson: value["viewportJson"],
    createdAt,
    updatedAt,
  };
}

export function normalizeWorkbookVersionRows(value: unknown): readonly WorkbookVersionRow[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.flatMap((entry) => {
    const row = normalizeWorkbookVersionRow(entry);
    return row ? [row] : [];
  });
}

export function selectWorkbookVersionEntries(input: {
  readonly rows: readonly WorkbookVersionRow[];
  readonly currentUserId: string;
}): readonly WorkbookVersionEntry[] {
  return input.rows
    .map((row) => {
      return {
        id: row.id,
        ownerUserId: row.ownerUserId,
        ownerLabel: formatWorkbookCollaboratorLabel(row.ownerUserId),
        name: row.name,
        revision: row.revision,
        sheetId: row.sheetId,
        sheetName: row.sheetName,
        address: row.address,
        viewport: normalizeViewport(row.viewportJson),
        createdAt: row.createdAt,
        updatedAt: row.updatedAt,
        targetLabel:
          row.sheetName && row.address ? `${row.sheetName}!${row.address}` : row.sheetName,
        canDelete: row.ownerUserId === input.currentUserId,
      } satisfies WorkbookVersionEntry;
    })
    .toSorted((left, right) => {
      if (left.updatedAt !== right.updatedAt) {
        return right.updatedAt - left.updatedAt;
      }
      return left.name.localeCompare(right.name);
    });
}

export function formatWorkbookVersionTimestamp(updatedAt: number): string {
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(updatedAt);
}
