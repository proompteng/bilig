import { describe, expect, it } from "vitest";
import {
  createWorkbookScenario,
  deleteWorkbookScenario,
  loadWorkbookScenarioByDocument,
} from "../workbook-scenario-store.js";
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

describe("workbook-scenario-store", () => {
  it("creates authoritative workbook scenario rows with source revision and anchor context", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 3, sheetName: "Revenue" } satisfies QueryResultRow]
          : null,
    ]);

    await createWorkbookScenario(queryable, {
      workbookId: "doc-1",
      documentId: "scenario:1",
      ownerUserId: "alex@example.com",
      name: "What-if margin case",
      baseRevision: 17,
      sheetName: "Revenue",
      address: "D12",
      viewport: {
        rowStart: 4,
        rowEnd: 22,
        colStart: 2,
        colEnd: 10,
      },
    });

    const query = latestQuery(queryable);
    expect(query.text).toContain("INSERT INTO workbook_scenario");
    expect(query.values?.slice(0, 9)).toEqual([
      "scenario:1",
      "doc-1",
      "alex@example.com",
      "What-if margin case",
      17,
      3,
      "Revenue",
      "D12",
      JSON.stringify({
        rowStart: 4,
        rowEnd: 22,
        colStart: 2,
        colEnd: 10,
      }),
    ]);
    expect(typeof query.values?.[9]).toBe("number");
    expect(typeof query.values?.[10]).toBe("number");
  });

  it("loads workbook scenarios by scenario document id", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM workbook_scenario")
          ? [
              {
                document_id: "scenario:2",
                workbook_id: "doc-2",
                owner_user_id: "sam@example.com",
                name: "Pricing branch",
                base_revision: 9,
                sheet_id: 5,
                sheet_name: "Pricing",
                address: "C7",
                viewport_json: {
                  rowStart: 1,
                  rowEnd: 18,
                  colStart: 1,
                  colEnd: 8,
                },
                created_at: 1_775_456_000_000,
                updated_at: 1_775_456_100_000,
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    await expect(loadWorkbookScenarioByDocument(queryable, "scenario:2")).resolves.toEqual({
      documentId: "scenario:2",
      workbookId: "doc-2",
      ownerUserId: "sam@example.com",
      name: "Pricing branch",
      baseRevision: 9,
      sheetId: 5,
      sheetName: "Pricing",
      address: "C7",
      viewport: {
        rowStart: 1,
        rowEnd: 18,
        colStart: 1,
        colEnd: 8,
      },
      createdAt: 1_775_456_000_000,
      updatedAt: 1_775_456_100_000,
    });
  });

  it("deletes only scenario rows owned by the current user", async () => {
    const queryable = new FakeQueryable();

    await deleteWorkbookScenario(queryable, {
      workbookId: "doc-1",
      documentId: "scenario:3",
      ownerUserId: "alex@example.com",
    });

    const query = latestQuery(queryable);
    expect(query.text).toContain("DELETE FROM workbook_scenario");
    expect(query.values).toEqual(["doc-1", "scenario:3", "alex@example.com"]);
  });
});
