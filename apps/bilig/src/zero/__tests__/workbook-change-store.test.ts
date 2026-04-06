import { describe, expect, it } from "vitest";
import {
  appendWorkbookChange,
  backfillWorkbookChanges,
  buildWorkbookChangeDescriptor,
  loadWorkbookChange,
} from "../workbook-change-store.js";
import type { QueryResultRow, Queryable } from "../store.js";

interface RecordedQuery {
  readonly text: string;
  readonly values: readonly unknown[] | undefined;
}

class FakeQueryable implements Queryable {
  readonly calls: RecordedQuery[] = [];

  constructor(
    private readonly responders: readonly ((
      text: string,
      values: readonly unknown[] | undefined,
    ) => QueryResultRow[] | null)[] = [],
  ) {}

  async query<T extends QueryResultRow = QueryResultRow>(
    text: string,
    values?: unknown[],
  ): Promise<{ rows: T[] }> {
    this.calls.push({ text, values });
    for (const responder of this.responders) {
      const rows = responder(text, values);
      if (rows) {
        return {
          rows: rows.filter((row): row is T => row !== null),
        };
      }
    }
    return { rows: [] };
  }
}

function latestQuery(queryable: FakeQueryable): RecordedQuery {
  const query = queryable.calls.at(-1);
  if (!query) {
    throw new Error("Expected at least one query");
  }
  return query;
}

describe("workbook-change-store", () => {
  it("summarizes renderCommit cell bundles as authoritative range changes", () => {
    expect(
      buildWorkbookChangeDescriptor({
        kind: "renderCommit",
        ops: [
          { kind: "upsertCell", sheetName: "Sheet1", addr: "B2", value: 1 },
          { kind: "upsertCell", sheetName: "Sheet1", addr: "C4", formula: "=SUM(B2:B3)" },
          { kind: "upsertCell", sheetName: "Sheet1", addr: "B4", value: 3 },
        ],
      }),
    ).toEqual({
      eventKind: "renderCommit",
      summary: "Updated 3 cells in Sheet1!B2:C4",
      sheetName: "Sheet1",
      anchorAddress: "B2",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "C4",
      },
    });
  });

  it("records workbook changes with resolved sheet ids and serialized ranges", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 4, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
    ]);

    await appendWorkbookChange(queryable, {
      documentId: "doc-1",
      revision: 7,
      actorUserId: "amy@example.com",
      clientMutationId: "mutation-7",
      payload: {
        kind: "fillRange",
        source: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A2",
        },
        target: {
          sheetName: "Sheet1",
          startAddress: "B1",
          endAddress: "B2",
        },
      },
      undoBundle: {
        kind: "engineOps",
        ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "B1" }],
      },
      createdAtUnixMs: 123_456,
    });

    const query = latestQuery(queryable);
    expect(query.text).toContain("INSERT INTO workbook_change");
    expect(query.values).toEqual([
      "doc-1",
      7,
      "amy@example.com",
      "mutation-7",
      "fillRange",
      "Filled Sheet1!B1:B2",
      4,
      "Sheet1",
      "B1",
      JSON.stringify({
        sheetName: "Sheet1",
        startAddress: "B1",
        endAddress: "B2",
      }),
      JSON.stringify({
        kind: "engineOps",
        ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "B1" }],
      }),
      null,
      null,
      123_456,
    ]);
  });

  it("backfills missing workbook_change rows from authoritative workbook events", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM workbook_event AS event")
          ? [
              {
                workbookId: "doc-1",
                revision: 9,
                actorUserId: "sam@example.com",
                clientMutationId: "mutation-9",
                payload: {
                  kind: "setCellFormula",
                  sheetName: "Sheet1",
                  address: "D5",
                  formula: "=SUM(A1:A4)",
                },
                createdAtUnixMs: 987_654,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 11, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
    ]);

    await backfillWorkbookChanges(queryable);

    const insertQuery = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_change"),
    );
    expect(insertQuery?.values).toEqual([
      "doc-1",
      9,
      "sam@example.com",
      "mutation-9",
      "setCellFormula",
      "Set formula in Sheet1!D5",
      11,
      "Sheet1",
      "D5",
      JSON.stringify({
        sheetName: "Sheet1",
        startAddress: "D5",
        endAddress: "D5",
      }),
      null,
      null,
      null,
      987_654,
    ]);
  });

  it("summarizes restoreVersion events as named workbook restores", () => {
    expect(
      buildWorkbookChangeDescriptor({
        kind: "restoreVersion",
        versionId: "version-1",
        versionName: "Month close",
        sheetName: "Sheet1",
        address: "D5",
        snapshot: {
          version: 1,
          workbook: {
            name: "doc-1",
          },
          sheets: [],
        },
      }),
    ).toEqual({
      eventKind: "restoreVersion",
      summary: "Restored version Month close",
      sheetName: "Sheet1",
      anchorAddress: "D5",
      range: {
        sheetName: "Sheet1",
        startAddress: "D5",
        endAddress: "D5",
      },
    });
  });

  it("records revert changes with target metadata and marks the original revision as reverted", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 4, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
    ]);

    await appendWorkbookChange(queryable, {
      documentId: "doc-1",
      revision: 8,
      actorUserId: "amy@example.com",
      clientMutationId: "mutation-8",
      payload: {
        kind: "revertChange",
        targetRevision: 7,
        targetSummary: "Filled Sheet1!B1:B2",
        sheetName: "Sheet1",
        address: "B1",
        range: {
          sheetName: "Sheet1",
          startAddress: "B1",
          endAddress: "B2",
        },
        appliedBundle: {
          kind: "engineOps",
          ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "B1" }],
        },
      },
      undoBundle: {
        kind: "engineOps",
        ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "B1", value: 1 }],
      },
      createdAtUnixMs: 456_789,
    });

    const insertQuery = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_change"),
    );
    expect(insertQuery?.values).toEqual([
      "doc-1",
      8,
      "amy@example.com",
      "mutation-8",
      "revertChange",
      "Reverted r7: Filled Sheet1!B1:B2",
      4,
      "Sheet1",
      "B1",
      JSON.stringify({
        sheetName: "Sheet1",
        startAddress: "B1",
        endAddress: "B2",
      }),
      JSON.stringify({
        kind: "engineOps",
        ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "B1", value: 1 }],
      }),
      null,
      7,
      456_789,
    ]);

    const updateQuery = latestQuery(queryable);
    expect(updateQuery.text).toContain("UPDATE workbook_change");
    expect(updateQuery.values).toEqual(["doc-1", 7, 8]);
  });

  it("loads persisted undo metadata and revert markers for a workbook change row", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM workbook_change")
          ? [
              {
                revision: 12,
                actorUserId: "alex@example.com",
                clientMutationId: "mutation-12",
                eventKind: "setCellValue",
                summary: "Updated Sheet1!A1",
                sheetId: 1,
                sheetName: "Sheet1",
                anchorAddress: "A1",
                rangeJson: { sheetName: "Sheet1", startAddress: "A1", endAddress: "A1" },
                undoBundleJson: {
                  kind: "engineOps",
                  ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "A1" }],
                },
                revertedByRevision: 13,
                revertsRevision: null,
                createdAtUnixMs: 123_000,
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    await expect(loadWorkbookChange(queryable, "doc-1", 12)).resolves.toEqual({
      revision: 12,
      actorUserId: "alex@example.com",
      clientMutationId: "mutation-12",
      eventKind: "setCellValue",
      summary: "Updated Sheet1!A1",
      sheetId: 1,
      sheetName: "Sheet1",
      anchorAddress: "A1",
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "A1",
      },
      undoBundle: {
        kind: "engineOps",
        ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "A1" }],
      },
      revertedByRevision: 13,
      revertsRevision: null,
      createdAtUnixMs: 123_000,
    });
  });
});
