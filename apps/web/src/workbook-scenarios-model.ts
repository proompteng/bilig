import type { Viewport } from "@bilig/protocol";
import { formatWorkbookCollaboratorLabel } from "./workbook-presence-model.js";

export interface WorkbookScenarioRow {
  readonly documentId: string;
  readonly workbookId: string;
  readonly ownerUserId: string;
  readonly name: string;
  readonly baseRevision: number;
  readonly sheetId: number | null;
  readonly sheetName: string | null;
  readonly address: string | null;
  readonly viewportJson: unknown;
  readonly createdAt: number;
  readonly updatedAt: number;
}

export interface WorkbookScenarioEntry {
  readonly documentId: string;
  readonly workbookId: string;
  readonly ownerUserId: string;
  readonly ownerLabel: string;
  readonly name: string;
  readonly baseRevision: number;
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

function normalizeWorkbookScenarioRow(value: unknown): WorkbookScenarioRow | null {
  if (!isRecord(value)) {
    return null;
  }
  const documentId = value["documentId"];
  const workbookId = value["workbookId"];
  const ownerUserId = value["ownerUserId"];
  const name = value["name"];
  const baseRevision = value["baseRevision"];
  const createdAt = value["createdAt"];
  const updatedAt = value["updatedAt"];
  if (
    typeof documentId !== "string" ||
    typeof workbookId !== "string" ||
    typeof ownerUserId !== "string" ||
    typeof name !== "string" ||
    typeof baseRevision !== "number" ||
    typeof createdAt !== "number" ||
    typeof updatedAt !== "number"
  ) {
    return null;
  }
  const sheetId = value["sheetId"];
  const sheetName = value["sheetName"];
  const address = value["address"];
  return {
    documentId,
    workbookId,
    ownerUserId,
    name,
    baseRevision,
    sheetId: typeof sheetId === "number" ? sheetId : null,
    sheetName: typeof sheetName === "string" ? sheetName : null,
    address: typeof address === "string" ? address : null,
    viewportJson: value["viewportJson"],
    createdAt,
    updatedAt,
  };
}

export function normalizeWorkbookScenarioRows(value: unknown): readonly WorkbookScenarioRow[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.flatMap((entry) => {
    const row = normalizeWorkbookScenarioRow(entry);
    return row ? [row] : [];
  });
}

export function normalizeWorkbookScenarioEntry(value: unknown): WorkbookScenarioEntry | null {
  const row = normalizeWorkbookScenarioRow(value);
  if (!row) {
    return null;
  }
  return {
    documentId: row.documentId,
    workbookId: row.workbookId,
    ownerUserId: row.ownerUserId,
    ownerLabel: formatWorkbookCollaboratorLabel(row.ownerUserId),
    name: row.name,
    baseRevision: row.baseRevision,
    sheetId: row.sheetId,
    sheetName: row.sheetName,
    address: row.address,
    viewport: normalizeViewport(row.viewportJson),
    createdAt: row.createdAt,
    updatedAt: row.updatedAt,
    targetLabel: row.sheetName && row.address ? `${row.sheetName}!${row.address}` : row.sheetName,
    canDelete: true,
  };
}

export function selectWorkbookScenarioEntries(input: {
  readonly rows: readonly WorkbookScenarioRow[];
  readonly currentUserId: string;
}): readonly WorkbookScenarioEntry[] {
  return input.rows
    .map((row) => {
      return {
        documentId: row.documentId,
        workbookId: row.workbookId,
        ownerUserId: row.ownerUserId,
        ownerLabel: formatWorkbookCollaboratorLabel(row.ownerUserId),
        name: row.name,
        baseRevision: row.baseRevision,
        sheetId: row.sheetId,
        sheetName: row.sheetName,
        address: row.address,
        viewport: normalizeViewport(row.viewportJson),
        createdAt: row.createdAt,
        updatedAt: row.updatedAt,
        targetLabel:
          row.sheetName && row.address ? `${row.sheetName}!${row.address}` : row.sheetName,
        canDelete: row.ownerUserId === input.currentUserId,
      } satisfies WorkbookScenarioEntry;
    })
    .toSorted((left, right) => {
      if (left.updatedAt !== right.updatedAt) {
        return right.updatedAt - left.updatedAt;
      }
      return left.name.localeCompare(right.name);
    });
}

export function formatWorkbookScenarioTimestamp(updatedAt: number): string {
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(updatedAt);
}
