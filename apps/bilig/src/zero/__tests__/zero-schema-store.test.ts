import { describe, expect, it, vi } from "vitest";

const repairFns = vi.hoisted(() => ({
  repairWorkbookSheetIds: vi.fn(),
}));

vi.mock("../sheet-id-repair.js", () => ({
  repairWorkbookSheetIds: repairFns.repairWorkbookSheetIds,
}));

import { ensureZeroSyncSchema } from "../zero-schema-store.js";
import type { Queryable } from "../store.js";

describe("zero schema store", () => {
  it("repairs sheet ids and backfills workbook snapshots using json-v1", async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] });
    const db: Queryable = { query };

    await ensureZeroSyncSchema(db);

    expect(repairFns.repairWorkbookSheetIds).toHaveBeenCalledWith(db);
    expect(query).toHaveBeenCalledWith(expect.stringContaining("INSERT INTO workbook_snapshot"), [
      "json-v1",
    ]);
  });
});
