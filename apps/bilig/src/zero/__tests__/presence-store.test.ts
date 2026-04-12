import { describe, expect, it } from "vitest";
import { WorkbookRuntimeManager } from "../../workbook-runtime/runtime-manager.js";
import { handleServerMutator } from "../server-mutators.js";
import {
  resolveWorkbookPresenceSheetRef,
  upsertWorkbookPresence,
  type UpsertWorkbookPresenceInput,
} from "../presence-store.js";
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

describe("presence-store", () => {
  it("resolves a workbook presence sheet ref by name", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 7, sheetName: "Revenue" } satisfies QueryResultRow]
          : null,
    ]);

    await expect(
      resolveWorkbookPresenceSheetRef(queryable, {
        documentId: "doc-1",
        sheetName: "Revenue",
      }),
    ).resolves.toEqual({
      sheetId: 7,
      sheetName: "Revenue",
    });
  });

  it("preserves the provided coarse location when the sheet lookup misses", async () => {
    const queryable = new FakeQueryable();

    await expect(
      resolveWorkbookPresenceSheetRef(queryable, {
        documentId: "doc-1",
        sheetName: "Ad hoc",
      }),
    ).resolves.toEqual({
      sheetId: null,
      sheetName: "Ad hoc",
    });
  });

  it("upserts coarse presence rows with resolved sheet ids and serialized selection", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 3, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
    ]);

    const input = {
      documentId: "doc-1",
      sessionId: "doc-1:browser:test",
      userId: "alex@example.com",
      presenceClientId: "presence:self",
      sheetName: "Sheet1",
      address: "B2",
      selection: { sheetName: "Sheet1", address: "B2" },
    } satisfies UpsertWorkbookPresenceInput;

    await upsertWorkbookPresence(queryable, input);

    const query = latestQuery(queryable);
    expect(query.text).toContain("INSERT INTO presence_coarse");
    expect(query.values?.slice(0, 8)).toEqual([
      "doc-1",
      "doc-1:browser:test",
      "alex@example.com",
      "presence:self",
      3,
      "Sheet1",
      "B2",
      JSON.stringify({ sheetName: "Sheet1", address: "B2" }),
    ]);
    expect(typeof query.values?.[8]).toBe("number");
  });

  it("routes workbook.updatePresence through the server mutator path", async () => {
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
      "workbook.updatePresence",
      {
        documentId: "doc-1",
        sessionId: "doc-1:browser:test",
        presenceClientId: "presence:self",
        sheetName: "Sheet1",
        address: "C4",
        selection: { sheetName: "Sheet1", address: "C4" },
      },
      new WorkbookRuntimeManager(),
      {
        userID: "sam@example.com",
        roles: ["editor"],
      },
    );

    const query = latestQuery(queryable);
    expect(query.text).toContain("INSERT INTO presence_coarse");
    expect(query.values?.slice(0, 8)).toEqual([
      "doc-1",
      "doc-1:browser:test",
      "sam@example.com",
      "presence:self",
      5,
      "Sheet1",
      "C4",
      JSON.stringify({ sheetName: "Sheet1", address: "C4" }),
    ]);
  });

  it("bootstraps a missing workbook before upserting presence", async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("INSERT INTO workbooks") ? [{ id: "doc-1" } satisfies QueryResultRow] : null,
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 1, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
    ]);

    await handleServerMutator(
      {
        dbTransaction: {
          wrappedTransaction: queryable,
        },
      },
      "workbook.updatePresence",
      {
        documentId: "doc-1",
        sessionId: "doc-1:browser:test",
        presenceClientId: "presence:self",
        sheetName: "Sheet1",
        address: "A1",
        selection: { sheetName: "Sheet1", address: "A1" },
      },
      new WorkbookRuntimeManager(),
      {
        userID: "guest@example.com",
        roles: ["editor"],
      },
    );

    const workbookInsertIndex = queryable.calls.findIndex((call) =>
      call.text.includes("INSERT INTO workbooks"),
    );
    const sheetInsertIndex = queryable.calls.findIndex((call) =>
      call.text.includes("INSERT INTO sheets"),
    );
    const presenceInsertIndex = queryable.calls.findIndex((call) =>
      call.text.includes("INSERT INTO presence_coarse"),
    );

    expect(workbookInsertIndex).toBeGreaterThanOrEqual(0);
    expect(sheetInsertIndex).toBeGreaterThan(workbookInsertIndex);
    expect(presenceInsertIndex).toBeGreaterThan(sheetInsertIndex);
  });
});
