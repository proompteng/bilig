import { describe, expect, it } from "vitest";
import { WorkbookRuntimeManager } from "../../workbook-runtime/runtime-manager.js";
import { handleServerMutator } from "../server-mutators.js";
import { reconcileWorkbookSheetViews, upsertWorkbookSheetView } from "../sheet-view-store.js";
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

describe("sheet-view-store", () => {
  it("upserts named workbook views with resolved stable sheet ids", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 7, sheetName: "Revenue" } satisfies QueryResultRow]
          : null,
    ]);

    await upsertWorkbookSheetView(queryable, {
      documentId: "doc-1",
      id: "view-1",
      ownerUserId: "alex@example.com",
      name: "My revenue focus",
      visibility: "private",
      sheetName: "Revenue",
      address: "D8",
      viewport: {
        rowStart: 10,
        rowEnd: 32,
        colStart: 3,
        colEnd: 12,
      },
    });

    const query = latestQuery(queryable);
    expect(query.text).toContain("INSERT INTO sheet_view");
    expect(query.values?.slice(0, 9)).toEqual([
      "doc-1",
      "view-1",
      "alex@example.com",
      "My revenue focus",
      "private",
      7,
      "Revenue",
      "D8",
      JSON.stringify({
        rowStart: 10,
        rowEnd: 32,
        colStart: 3,
        colEnd: 12,
      }),
    ]);
    expect(typeof query.values?.[9]).toBe("number");
  });

  it("rejects ownership changes for an existing workbook view id", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheet_view")
          ? [{ ownerUserId: "other@example.com" } satisfies QueryResultRow]
          : null,
    ]);

    await expect(
      upsertWorkbookSheetView(queryable, {
        documentId: "doc-1",
        id: "view-1",
        ownerUserId: "alex@example.com",
        name: "Revenue",
        visibility: "shared",
        sheetId: 7,
        address: "A1",
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 8,
        },
      }),
    ).rejects.toThrow("owned by another user");
  });

  it("reconciles renamed and deleted sheet anchors", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 12, sheetName: "Forecast" } satisfies QueryResultRow]
          : null,
    ]);

    await reconcileWorkbookSheetViews({
      db: queryable,
      documentId: "doc-1",
      payload: {
        kind: "renderCommit",
        ops: [
          { kind: "renameSheet", oldName: "Plan", newName: "Forecast" },
          { kind: "deleteSheet", name: "Archive" },
        ],
      },
    });

    expect(queryable.calls.some((call) => call.text.includes("UPDATE sheet_view"))).toBe(true);
    expect(queryable.calls.some((call) => call.text.includes("DELETE FROM sheet_view"))).toBe(true);
  });

  it("routes workbook.upsertSheetView and workbook.deleteSheetView through the server mutator path", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 5, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
    ]);

    await handleServerMutator(
      {
        dbTransaction: {
          wrappedTransaction: queryable,
        },
      },
      "workbook.upsertSheetView",
      {
        documentId: "doc-1",
        id: "view-1",
        name: "Shared QA",
        visibility: "shared",
        sheetName: "Sheet1",
        address: "C4",
        viewport: {
          rowStart: 2,
          rowEnd: 14,
          colStart: 2,
          colEnd: 9,
        },
      },
      new WorkbookRuntimeManager(),
      {
        userID: "sam@example.com",
        roles: ["editor"],
      },
    );

    expect(latestQuery(queryable).text).toContain("INSERT INTO sheet_view");

    await handleServerMutator(
      {
        dbTransaction: {
          wrappedTransaction: queryable,
        },
      },
      "workbook.deleteSheetView",
      {
        documentId: "doc-1",
        id: "view-1",
      },
      new WorkbookRuntimeManager(),
      {
        userID: "sam@example.com",
        roles: ["editor"],
      },
    );

    expect(latestQuery(queryable).text).toContain("DELETE FROM sheet_view");
  });
});
