import { describe, expect, it, vi } from "vitest";
import { createEmptyWorkbookSnapshot } from "../store-support.js";

const storeFns = vi.hoisted(() => ({
  persistCellEvalDiff: vi.fn(),
  persistCellEvalIncremental: vi.fn(),
  persistWorkbookCheckpoint: vi.fn(),
  shouldPersistWorkbookCheckpointRevision: vi.fn((revision: number) => revision === 64),
}));

vi.mock("../store.js", () => ({
  shouldPersistWorkbookCheckpointRevision: storeFns.shouldPersistWorkbookCheckpointRevision,
}));

vi.mock("../workbook-calculation-store.js", () => ({
  persistCellEvalDiff: storeFns.persistCellEvalDiff,
  persistCellEvalIncremental: storeFns.persistCellEvalIncremental,
  persistWorkbookCheckpoint: storeFns.persistWorkbookCheckpoint,
}));

import {
  leaseNextRecalcJob,
  markRecalcJobCompleted,
  markRecalcJobFailed,
} from "../recalc-job-store.js";
import type { Queryable } from "../store.js";

describe("recalc job store", () => {
  it("filters invalid dirty regions when leasing a job", async () => {
    const db: Queryable = {
      query: vi.fn().mockResolvedValue({
        rows: [
          {
            id: "job-1",
            workbook_id: "book-1",
            from_revision: "3",
            to_revision: "7",
            dirty_regions_json: [
              {
                sheetName: "Sheet1",
                rowStart: 1,
                rowEnd: 2,
                colStart: 3,
                colEnd: 4,
              },
              { nope: true },
            ],
            attempts: "2",
          },
        ],
      }),
    };

    await expect(leaseNextRecalcJob(db, "worker-1")).resolves.toEqual({
      id: "job-1",
      workbookId: "book-1",
      fromRevision: 3,
      toRevision: 7,
      dirtyRegions: [
        {
          sheetName: "Sheet1",
          rowStart: 1,
          rowEnd: 2,
          colStart: 3,
          colEnd: 4,
        },
      ],
      attempts: 2,
    });
  });

  it("persists incremental results and checkpoints completed lease revisions", async () => {
    const db: Queryable = {
      query: vi
        .fn()
        .mockResolvedValueOnce({ rows: [{ head_revision: 64 }] })
        .mockResolvedValue({ rows: [] }),
    };
    const snapshot = createEmptyWorkbookSnapshot("book-1");

    await expect(
      markRecalcJobCompleted(
        db,
        {
          id: "job-1",
          workbookId: "book-1",
          fromRevision: 1,
          toRevision: 64,
          dirtyRegions: null,
          attempts: 1,
        },
        [],
        snapshot,
        null,
        true,
      ),
    ).resolves.toBe(true);

    expect(storeFns.persistCellEvalIncremental).toHaveBeenCalledWith(db, "book-1", []);
    expect(storeFns.persistCellEvalDiff).not.toHaveBeenCalled();
    expect(storeFns.persistWorkbookCheckpoint).toHaveBeenCalledWith(
      db,
      "book-1",
      64,
      snapshot,
      null,
    );
  });

  it("marks exhausted failures as failed instead of pending", async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] });
    const db: Queryable = { query };

    await markRecalcJobFailed(
      db,
      {
        id: "job-9",
        workbookId: "book-1",
        fromRevision: 1,
        toRevision: 2,
        dirtyRegions: null,
        attempts: 3,
      },
      new Error("boom"),
    );

    expect(query).toHaveBeenCalledWith(expect.any(String), [
      "job-9",
      "failed",
      expect.stringContaining("boom"),
    ]);
  });
});
