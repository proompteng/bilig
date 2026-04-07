import { describe, expect, it } from "vitest";
import {
  DEFAULT_ZERO_PUBLICATION,
  ensureZeroPublication,
  resolveZeroPublicationName,
  ZERO_PUBLICATION_TABLES,
} from "../publication-store.js";
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

function isPublicationLookup(text: string): boolean {
  return text.includes("FROM pg_publication\n") && !text.includes("pg_publication_tables");
}

function isPublicationTableLookup(text: string): boolean {
  return text.includes("FROM pg_publication_tables");
}

function latestQuery(queryable: FakeQueryable): RecordedQuery {
  const query = queryable.calls.at(-1);
  if (!query) {
    throw new Error("Expected at least one query");
  }
  return query;
}

describe("publication-store", () => {
  it("creates the publication with every replicated table when it is missing", async () => {
    const queryable = new FakeQueryable();

    await ensureZeroPublication(queryable);

    const query = latestQuery(queryable);
    expect(query.text).toContain(`CREATE PUBLICATION "${DEFAULT_ZERO_PUBLICATION}" FOR TABLE`);
    ZERO_PUBLICATION_TABLES.forEach((tableName) => {
      expect(query.text).toContain(`public."${tableName}"`);
    });
  });

  it("adds only the missing replicated tables when the publication already exists", async () => {
    const queryable = new FakeQueryable([
      (text) => (isPublicationLookup(text) ? [{ present: 1 } satisfies QueryResultRow] : null),
      (text) =>
        isPublicationTableLookup(text)
          ? [
              { tableName: "workbooks" } satisfies QueryResultRow,
              { tableName: "sheets" } satisfies QueryResultRow,
              { tableName: "cell_styles" } satisfies QueryResultRow,
              { tableName: "cell_number_formats" } satisfies QueryResultRow,
              { tableName: "cells" } satisfies QueryResultRow,
              { tableName: "row_metadata" } satisfies QueryResultRow,
              { tableName: "column_metadata" } satisfies QueryResultRow,
              { tableName: "cell_eval" } satisfies QueryResultRow,
              { tableName: "defined_names" } satisfies QueryResultRow,
            ]
          : null,
    ]);

    await ensureZeroPublication(queryable);

    const query = latestQuery(queryable);
    expect(query.text).toContain(`ALTER PUBLICATION "${DEFAULT_ZERO_PUBLICATION}" ADD TABLE`);
    expect(query.text).toContain(`public."presence_coarse"`);
    expect(query.text).toContain(`public."workbook_change"`);
    expect(query.text).not.toContain(`public."workbooks"`);
    expect(query.text).not.toContain(`public."defined_names"`);
  });

  it("does not mutate the publication when every replicated table is already present", async () => {
    const queryable = new FakeQueryable([
      (text) => (isPublicationLookup(text) ? [{ present: 1 } satisfies QueryResultRow] : null),
      (text) =>
        isPublicationTableLookup(text)
          ? ZERO_PUBLICATION_TABLES.map((tableName) => ({ tableName }) satisfies QueryResultRow)
          : null,
    ]);

    await ensureZeroPublication(queryable);

    expect(queryable.calls).toHaveLength(2);
  });

  it("rejects invalid publication names from the environment", () => {
    expect(() => resolveZeroPublicationName({ BILIG_ZERO_PUBLICATION: "bad-name!" })).toThrow(
      "Invalid Zero publication name: bad-name!",
    );
  });
});
