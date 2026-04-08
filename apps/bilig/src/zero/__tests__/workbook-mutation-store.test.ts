import { describe, expect, it, vi, beforeEach } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";

const projectionFns = vi.hoisted(() => ({
  buildCalculationSettingsRowFromEngine: vi.fn(),
  buildSheetCellSourceRowsFromEngine: vi.fn(),
  buildSheetColumnMetadataRowsFromEngine: vi.fn(),
  buildSingleCellSourceRowFromEngine: vi.fn(),
  buildWorkbookHeaderRowFromEngine: vi.fn(),
  buildWorkbookNumberFormatRowsFromEngine: vi.fn(),
  buildWorkbookSourceProjectionFromEngine: vi.fn(),
  buildWorkbookStyleRowsFromEngine: vi.fn(),
  materializeCellEvalProjection: vi.fn(),
}));

const storeFns = vi.hoisted(() => ({
  applyAxisMetadataDiff: vi.fn(),
  applyCalculationSettings: vi.fn(),
  applyCellDiff: vi.fn(),
  applyNumberFormatDiff: vi.fn(),
  applySourceProjectionDiff: vi.fn(),
  applyStyleDiff: vi.fn(),
  persistCellEvalRangeDiff: vi.fn(),
  persistCellSourceRange: vi.fn(),
  upsertWorkbookHeader: vi.fn(),
}));

const changeStoreFns = vi.hoisted(() => ({
  appendWorkbookChange: vi.fn(),
}));

vi.mock("../projection.js", () => ({
  buildCalculationSettingsRowFromEngine: projectionFns.buildCalculationSettingsRowFromEngine,
  buildSheetCellSourceRowsFromEngine: projectionFns.buildSheetCellSourceRowsFromEngine,
  buildSheetColumnMetadataRowsFromEngine: projectionFns.buildSheetColumnMetadataRowsFromEngine,
  buildSingleCellSourceRowFromEngine: projectionFns.buildSingleCellSourceRowFromEngine,
  buildWorkbookHeaderRowFromEngine: projectionFns.buildWorkbookHeaderRowFromEngine,
  buildWorkbookNumberFormatRowsFromEngine: projectionFns.buildWorkbookNumberFormatRowsFromEngine,
  buildWorkbookSourceProjectionFromEngine: projectionFns.buildWorkbookSourceProjectionFromEngine,
  buildWorkbookStyleRowsFromEngine: projectionFns.buildWorkbookStyleRowsFromEngine,
  materializeCellEvalProjection: projectionFns.materializeCellEvalProjection,
}));

vi.mock("../store.js", () => ({
  applyAxisMetadataDiff: storeFns.applyAxisMetadataDiff,
  applyCalculationSettings: storeFns.applyCalculationSettings,
  applyCellDiff: storeFns.applyCellDiff,
  applyNumberFormatDiff: storeFns.applyNumberFormatDiff,
  applySourceProjectionDiff: storeFns.applySourceProjectionDiff,
  applyStyleDiff: storeFns.applyStyleDiff,
  persistCellEvalRangeDiff: storeFns.persistCellEvalRangeDiff,
  persistCellSourceRange: storeFns.persistCellSourceRange,
  upsertWorkbookHeader: storeFns.upsertWorkbookHeader,
}));

vi.mock("../workbook-change-store.js", () => ({
  appendWorkbookChange: changeStoreFns.appendWorkbookChange,
}));

import { persistWorkbookMutation } from "../workbook-mutation-store.js";
import type { PersistWorkbookMutationOptions, Queryable } from "../store.js";

function makeBaseOptions(
  overrides: Partial<PersistWorkbookMutationOptions> = {},
): PersistWorkbookMutationOptions {
  return {
    previousState: {
      projection: {
        workbook: {
          id: "book-1",
          name: "Book",
          ownerUserId: "owner-1",
          headRevision: 1,
          calculatedRevision: 1,
          sourceProjectionVersion: 2,
          calcMode: "auto",
          compatibilityMode: "excel",
          recalcEpoch: 0,
          updatedAt: "2026-01-01T00:00:00.000Z",
        },
        sheets: [],
        cells: [],
        rowMetadata: [],
        columnMetadata: [],
        definedNames: [],
        workbookMetadataEntries: [],
        calculationSettings: {
          workbookId: "book-1",
          mode: "auto",
          recalcEpoch: 0,
        },
        styles: [],
        numberFormats: [],
      },
      headRevision: 1,
      calculatedRevision: 1,
      ownerUserId: "owner-1",
    },
    nextEngine: new SpreadsheetEngine({
      workbookName: "book-1",
      replicaId: "mutation-test",
    }),
    updatedBy: "user-1",
    ownerUserId: "owner-1",
    eventPayload: {
      kind: "setCellValue",
      sheetName: "Sheet1",
      address: "A1",
      value: 123,
    },
    undoBundle: null,
    ...overrides,
  };
}

describe("workbook mutation store", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    projectionFns.buildWorkbookHeaderRowFromEngine.mockReturnValue({
      workbookId: "book-1",
      id: "book-1",
      name: "Book",
      ownerUserId: "owner-1",
      headRevision: 2,
      calculatedRevision: 2,
      sourceProjectionVersion: 2,
      calcMode: "auto",
      compatibilityMode: "excel",
      recalcEpoch: 0,
      updatedAt: "2026-01-01T00:00:00.000Z",
    });
    projectionFns.buildCalculationSettingsRowFromEngine.mockReturnValue({
      workbookId: "book-1",
      mode: "auto",
      recalcEpoch: 0,
    });
    projectionFns.buildSingleCellSourceRowFromEngine.mockReturnValue({
      workbookId: "book-1",
      sheetName: "Sheet1",
      address: "A1",
      rowNum: 0,
      colNum: 0,
      inputValue: 123,
      formula: null,
      format: null,
      styleId: null,
      explicitFormatId: null,
      updatedAt: "2026-01-01T00:00:00.000Z",
    });
    projectionFns.buildWorkbookStyleRowsFromEngine.mockReturnValue([
      {
        workbookId: "book-1",
        id: "style-1",
        recordJSON: { id: "style-1" },
        hash: "hash-1",
        createdAt: "2026-01-01T00:00:00.000Z",
      },
    ]);
    projectionFns.buildSheetCellSourceRowsFromEngine.mockReturnValue([
      {
        workbookId: "book-1",
        sheetName: "Sheet1",
        address: "A1",
        rowNum: 0,
        colNum: 0,
        inputValue: 123,
        formula: null,
        format: null,
        styleId: "style-1",
        explicitFormatId: null,
        updatedAt: "2026-01-01T00:00:00.000Z",
      },
    ]);
    projectionFns.materializeCellEvalProjection.mockReturnValue([]);
  });

  it("queues recalculation work for value mutations", async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] });
    const db: Queryable = { query };

    const result = await persistWorkbookMutation(db, "book-1", makeBaseOptions());

    expect(result.revision).toBe(2);
    expect(result.recalcJobId).toBe("book-1:recalc:2");
    expect(storeFns.applyCellDiff).toHaveBeenCalledOnce();
    expect(changeStoreFns.appendWorkbookChange).toHaveBeenCalledOnce();
    expect(
      query.mock.calls.some(
        ([sql]) => typeof sql === "string" && sql.includes("INSERT INTO recalc_job"),
      ),
    ).toBe(true);
  });

  it("does not queue recalculation work for formatting-only mutations when revisions are current", async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] });
    const db: Queryable = { query };

    const result = await persistWorkbookMutation(
      db,
      "book-1",
      makeBaseOptions({
        eventPayload: {
          kind: "setRangeStyle",
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "B2",
          },
          patch: { font: { bold: true } },
        },
      }),
    );

    expect(result.recalcJobId).toBeNull();
    expect(storeFns.applyStyleDiff).toHaveBeenCalledOnce();
    expect(storeFns.persistCellSourceRange).toHaveBeenCalledOnce();
    expect(
      query.mock.calls.some(
        ([sql]) => typeof sql === "string" && sql.includes("INSERT INTO recalc_job"),
      ),
    ).toBe(false);
  });
});
