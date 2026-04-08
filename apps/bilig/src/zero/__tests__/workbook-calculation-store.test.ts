import { describe, expect, it, vi } from "vitest";
import { backfillWorkbookSnapshotsFromInlineState } from "../workbook-calculation-store.js";
import type { Queryable } from "../store.js";

describe("workbook calculation store", () => {
  it("backfills json-v1 workbook snapshots from inline state", async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] });
    const db: Queryable = { query };

    await backfillWorkbookSnapshotsFromInlineState(db);

    expect(query).toHaveBeenCalledWith(expect.stringContaining("INSERT INTO workbook_snapshot"), [
      "json-v1",
    ]);
  });
});
