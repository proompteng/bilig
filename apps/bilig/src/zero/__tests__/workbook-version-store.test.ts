import { SpreadsheetEngine } from "@bilig/core";
import { describe, expect, it } from "vitest";
import { WorkbookRuntimeManager } from "../../workbook-runtime/runtime-manager.js";
import { handleServerMutator } from "../server-mutators.js";
import {
  createWorkbookVersion,
  deleteWorkbookVersion,
  loadWorkbookVersion,
} from "../workbook-version-store.js";
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

async function createSnapshot(
  cellAddress: string,
  value: string,
): Promise<ReturnType<SpreadsheetEngine["exportSnapshot"]>> {
  const engine = new SpreadsheetEngine({ workbookName: "doc-1", replicaId: "seed" });
  await engine.ready();
  engine.setCellValue("Sheet1", cellAddress, value);
  return engine.exportSnapshot();
}

describe("workbook-version-store", () => {
  it("creates authoritative workbook versions with snapshot payloads and sheet anchors", async () => {
    const snapshot = await createSnapshot("B4", "ready");
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 3, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
    ]);

    await createWorkbookVersion(queryable, {
      documentId: "doc-1",
      id: "version-1",
      ownerUserId: "alex@example.com",
      name: "Daily checkpoint",
      revision: 7,
      snapshot,
      sheetName: "Sheet1",
      address: "B4",
      viewport: {
        rowStart: 1,
        rowEnd: 12,
        colStart: 1,
        colEnd: 8,
      },
    });

    const query = latestQuery(queryable);
    expect(query.text).toContain("INSERT INTO workbook_version");
    expect(query.values?.slice(0, 10)).toEqual([
      "doc-1",
      "version-1",
      "alex@example.com",
      "Daily checkpoint",
      7,
      JSON.stringify(snapshot),
      3,
      "Sheet1",
      "B4",
      JSON.stringify({
        rowStart: 1,
        rowEnd: 12,
        colStart: 1,
        colEnd: 8,
      }),
    ]);
    expect(typeof query.values?.[10]).toBe("number");
    expect(typeof query.values?.[11]).toBe("number");
  });

  it("loads stored workbook versions with parsed snapshots", async () => {
    const snapshot = await createSnapshot("C2", "restored");
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM workbook_version")
          ? [
              {
                id: "version-2",
                owner_user_id: "sam@example.com",
                name: "Month close",
                revision: 11,
                snapshot_json: snapshot,
                sheet_id: 5,
                sheet_name: "Sheet1",
                address: "C2",
                viewport_json: {
                  rowStart: 2,
                  rowEnd: 20,
                  colStart: 2,
                  colEnd: 10,
                },
              } satisfies QueryResultRow,
            ]
          : null,
    ]);

    await expect(loadWorkbookVersion(queryable, "doc-1", "version-2")).resolves.toEqual({
      id: "version-2",
      ownerUserId: "sam@example.com",
      name: "Month close",
      revision: 11,
      snapshot,
      sheetId: 5,
      sheetName: "Sheet1",
      address: "C2",
      viewport: {
        rowStart: 2,
        rowEnd: 20,
        colStart: 2,
        colEnd: 10,
      },
    });
  });

  it("deletes only versions owned by the current user", async () => {
    const queryable = new FakeQueryable();

    await deleteWorkbookVersion(queryable, {
      documentId: "doc-1",
      id: "version-3",
      ownerUserId: "alex@example.com",
    });

    const query = latestQuery(queryable);
    expect(query.text).toContain("DELETE FROM workbook_version");
    expect(query.values).toEqual(["doc-1", "version-3", "alex@example.com"]);
  });

  it("routes create, delete, and restore version mutators through the authoritative server path", async () => {
    const currentSnapshot = await createSnapshot("A1", "before");
    const restoredSnapshot = await createSnapshot("D5", "after");
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 1, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
      (text) =>
        text.includes("FROM workbook_version")
          ? [
              {
                id: "version-restore",
                owner_user_id: "alex@example.com",
                name: "Restore me",
                revision: 4,
                snapshot_json: restoredSnapshot,
                sheet_id: 1,
                sheet_name: "Sheet1",
                address: "D5",
                viewport_json: {
                  rowStart: 4,
                  rowEnd: 18,
                  colStart: 3,
                  colEnd: 12,
                },
              } satisfies QueryResultRow,
            ]
          : null,
    ]);
    const runtimeManager = new WorkbookRuntimeManager({
      loadMetadata: async () => ({
        headRevision: 3,
        calculatedRevision: 3,
        ownerUserId: "alex@example.com",
      }),
      loadState: async () => ({
        snapshot: currentSnapshot,
        replicaSnapshot: null,
        headRevision: 3,
        calculatedRevision: 3,
        ownerUserId: "alex@example.com",
      }),
    });

    await handleServerMutator(
      {
        dbTransaction: {
          wrappedTransaction: queryable,
        },
      },
      "workbook.createVersion",
      {
        documentId: "doc-1",
        id: "version-1",
        name: "Checkpoint",
        sheetName: "Sheet1",
        address: "A1",
        viewport: {
          rowStart: 0,
          rowEnd: 10,
          colStart: 0,
          colEnd: 8,
        },
      },
      runtimeManager,
      {
        userID: "alex@example.com",
        roles: ["editor"],
      },
    );

    expect(queryable.calls.some((call) => call.text.includes("INSERT INTO workbook_version"))).toBe(
      true,
    );

    await handleServerMutator(
      {
        dbTransaction: {
          wrappedTransaction: queryable,
        },
      },
      "workbook.restoreVersion",
      {
        documentId: "doc-1",
        id: "version-restore",
      },
      runtimeManager,
      {
        userID: "alex@example.com",
        roles: ["editor"],
      },
    );

    const workbookEventInsert = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_event"),
    );
    expect(workbookEventInsert?.values?.[4]).toContain('"kind":"restoreVersion"');
    expect(workbookEventInsert?.values?.[4]).toContain('"versionName":"Restore me"');

    await handleServerMutator(
      {
        dbTransaction: {
          wrappedTransaction: queryable,
        },
      },
      "workbook.deleteVersion",
      {
        documentId: "doc-1",
        id: "version-1",
      },
      runtimeManager,
      {
        userID: "alex@example.com",
        roles: ["editor"],
      },
    );

    expect(latestQuery(queryable).text).toContain("DELETE FROM workbook_version");
  });

  it("routes revertChange through the authoritative mutation path and records a revert event", async () => {
    const currentSnapshot = await createSnapshot("A1", "before");
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes("FROM workbook_change") && values?.[1] === 7
          ? [
              {
                revision: 7,
                actorUserId: "sam@example.com",
                clientMutationId: "mutation-7",
                eventKind: "setCellValue",
                summary: "Updated Sheet1!A1",
                sheetId: 1,
                sheetName: "Sheet1",
                anchorAddress: "A1",
                rangeJson: {
                  sheetName: "Sheet1",
                  startAddress: "A1",
                  endAddress: "A1",
                },
                undoBundleJson: {
                  kind: "engineOps",
                  ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "A1" }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_000,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) =>
        text.includes("FROM sheets")
          ? [{ sheetId: 1, sheetName: "Sheet1" } satisfies QueryResultRow]
          : null,
    ]);
    const runtimeManager = new WorkbookRuntimeManager({
      loadMetadata: async () => ({
        headRevision: 7,
        calculatedRevision: 7,
        ownerUserId: "alex@example.com",
      }),
      loadState: async () => ({
        snapshot: currentSnapshot,
        replicaSnapshot: null,
        headRevision: 7,
        calculatedRevision: 7,
        ownerUserId: "alex@example.com",
      }),
    });

    await handleServerMutator(
      {
        dbTransaction: {
          wrappedTransaction: queryable,
        },
      },
      "workbook.revertChange",
      {
        documentId: "doc-1",
        revision: 7,
      },
      runtimeManager,
      {
        userID: "alex@example.com",
        roles: ["editor"],
      },
    );

    const workbookEventInsert = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_event"),
    );
    expect(workbookEventInsert?.values?.[4]).toContain('"kind":"revertChange"');
    expect(workbookEventInsert?.values?.[4]).toContain('"targetRevision":7');

    const workbookChangeInsert = queryable.calls.find((call) =>
      call.text.includes("INSERT INTO workbook_change"),
    );
    expect(workbookChangeInsert?.values?.[4]).toBe("revertChange");

    const revertMarkQuery = queryable.calls.find((call) =>
      call.text.includes("UPDATE workbook_change"),
    );
    expect(revertMarkQuery?.values).toEqual(["doc-1", 7, 8]);
  });
});
