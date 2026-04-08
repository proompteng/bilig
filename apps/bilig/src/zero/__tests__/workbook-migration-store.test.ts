import { describe, expect, it, vi } from "vitest";

const storeFns = vi.hoisted(() => ({
  applyAxisMetadataDiff: vi.fn(),
  applyCalculationSettings: vi.fn(),
  applyCellDiff: vi.fn(),
  applyDefinedNameDiff: vi.fn(),
  applyNumberFormatDiff: vi.fn(),
  applySheetDiff: vi.fn(),
  applyStyleDiff: vi.fn(),
  applyWorkbookMetadataDiff: vi.fn(),
  insertWorkbookHeaderIfMissing: vi.fn(),
  persistCellEvalRows: vi.fn(),
  upsertWorkbookHeader: vi.fn(),
}));

vi.mock("../store.js", () => ({
  applyAxisMetadataDiff: storeFns.applyAxisMetadataDiff,
  applyCalculationSettings: storeFns.applyCalculationSettings,
  applyCellDiff: storeFns.applyCellDiff,
  applyDefinedNameDiff: storeFns.applyDefinedNameDiff,
  applyNumberFormatDiff: storeFns.applyNumberFormatDiff,
  applySheetDiff: storeFns.applySheetDiff,
  applyStyleDiff: storeFns.applyStyleDiff,
  applyWorkbookMetadataDiff: storeFns.applyWorkbookMetadataDiff,
  insertWorkbookHeaderIfMissing: storeFns.insertWorkbookHeaderIfMissing,
  upsertWorkbookHeader: storeFns.upsertWorkbookHeader,
}));

vi.mock("../workbook-calculation-store.js", () => ({
  persistCellEvalRows: storeFns.persistCellEvalRows,
}));

import {
  backfillAuthoritativeCellEval,
  dropLegacyZeroSyncSchemaObjects,
  ensureWorkbookDocumentExists,
} from "../workbook-migration-store.js";
import type { Queryable } from "../store.js";

describe("workbook migration store", () => {
  it("skips projection replacement when the workbook already exists", async () => {
    storeFns.insertWorkbookHeaderIfMissing.mockResolvedValueOnce(false);
    const query = vi.fn();
    const db: Queryable = { query };

    await ensureWorkbookDocumentExists(db, "book-1", "owner-1");

    expect(storeFns.insertWorkbookHeaderIfMissing).toHaveBeenCalledOnce();
    expect(storeFns.applySheetDiff).not.toHaveBeenCalled();
    expect(query).not.toHaveBeenCalled();
  });

  it("drops the legacy zero-sync schema objects", async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] });
    const db: Queryable = { query };

    await dropLegacyZeroSyncSchemaObjects(db);

    expect(query.mock.calls).toEqual([
      [`DROP INDEX IF EXISTS sheet_style_ranges_workbook_sheet_idx`],
      [`DROP INDEX IF EXISTS sheet_format_ranges_workbook_sheet_idx`],
      [`DROP TABLE IF EXISTS sheet_style_ranges`],
      [`DROP TABLE IF EXISTS sheet_format_ranges`],
    ]);
  });

  it("returns early from backfill when no legacy workbook ids are found", async () => {
    const query = vi
      .fn()
      .mockResolvedValueOnce({ rows: [{ relation: null }] })
      .mockResolvedValueOnce({ rows: [{ relation: null }] })
      .mockResolvedValueOnce({ rows: [] })
      .mockResolvedValueOnce({ rows: [] });
    const db: Queryable = { query };

    await backfillAuthoritativeCellEval(db);

    expect(query).toHaveBeenCalledTimes(4);
    expect(storeFns.upsertWorkbookHeader).not.toHaveBeenCalled();
    expect(storeFns.persistCellEvalRows).not.toHaveBeenCalled();
  });
});
