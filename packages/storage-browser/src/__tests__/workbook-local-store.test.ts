import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import {
  createMemoryWorkbookLocalStoreFactory,
  type WorkbookLocalAuthoritativeBase,
  type WorkbookLocalMutationRecord,
  type WorkbookLocalProjectionOverlay,
} from "../index.js";

function createBase(value: number): WorkbookLocalAuthoritativeBase {
  return {
    sheets: [
      {
        sheetId: 1,
        name: "Sheet1",
        sortOrder: 0,
        freezeRows: 0,
        freezeCols: 0,
      },
    ],
    cellInputs: [
      {
        sheetId: 1,
        sheetName: "Sheet1",
        address: "A1",
        rowNum: 0,
        colNum: 0,
        input: value,
        formula: undefined,
        format: undefined,
      },
    ],
    cellRenders: [
      {
        sheetId: 1,
        sheetName: "Sheet1",
        address: "A1",
        rowNum: 0,
        colNum: 0,
        value: { tag: ValueTag.Number, value },
        flags: 0,
        version: 1,
        styleId: undefined,
        numberFormatId: undefined,
      },
    ],
    rowAxisEntries: [],
    columnAxisEntries: [],
    styles: [],
  };
}

function createOverlay(value: number): WorkbookLocalProjectionOverlay {
  return {
    cells: [
      {
        sheetId: 1,
        sheetName: "Sheet1",
        address: "A1",
        rowNum: 0,
        colNum: 0,
        value: { tag: ValueTag.Number, value },
        flags: 0,
        version: 2,
        input: value,
        formula: undefined,
        format: undefined,
        styleId: undefined,
        numberFormatId: undefined,
      },
    ],
    rowAxisEntries: [],
    columnAxisEntries: [],
    styles: [],
  };
}

function createMutation(
  overrides: Partial<WorkbookLocalMutationRecord> = {},
): WorkbookLocalMutationRecord {
  return {
    id: "memory-doc:pending:1",
    localSeq: 1,
    baseRevision: 0,
    method: "setCellValue",
    args: ["Sheet1", "A1", 17],
    enqueuedAtUnixMs: 100,
    submittedAtUnixMs: null,
    lastAttemptedAtUnixMs: null,
    ackedAtUnixMs: null,
    rebasedAtUnixMs: null,
    failedAtUnixMs: null,
    attemptCount: 0,
    failureMessage: null,
    status: "local",
    ...overrides,
  };
}

describe("memory workbook local store", () => {
  it("persists runtime state and normalized projection data across reopen", async () => {
    const factory = createMemoryWorkbookLocalStoreFactory();
    const store = await factory.open("memory-doc");

    await store.persistProjectionState({
      state: {
        snapshot: { version: 1, workbook: { name: "memory-doc" }, sheets: [] },
        replica: { replica: { id: "seed", clock: 0 }, entityVersions: [], sheetDeleteVersions: [] },
        authoritativeRevision: 7,
        appliedPendingLocalSeq: 3,
      },
      authoritativeBase: createBase(11),
      projectionOverlay: createOverlay(17),
    });
    store.close();

    const reopened = await factory.open("memory-doc");
    await expect(reopened.loadBootstrapState()).resolves.toEqual({
      workbookName: "memory-doc",
      sheetNames: ["Sheet1"],
      materializedCellCount: 1,
      authoritativeRevision: 7,
      appliedPendingLocalSeq: 3,
    });
    expect(await reopened.loadState()).toMatchObject({
      authoritativeRevision: 7,
      appliedPendingLocalSeq: 3,
    });
    expect(
      reopened.readViewportProjection("Sheet1", {
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      }),
    ).toMatchObject({
      cells: [
        {
          snapshot: {
            sheetName: "Sheet1",
            address: "A1",
            value: { tag: ValueTag.Number, value: 17 },
          },
        },
      ],
    });
  });

  it("keeps acked mutations in the journal while filtering them from the active list", async () => {
    const factory = createMemoryWorkbookLocalStoreFactory();
    const store = await factory.open("memory-doc");
    const local = createMutation();
    const acked = {
      ...local,
      submittedAtUnixMs: 120,
      lastAttemptedAtUnixMs: 120,
      ackedAtUnixMs: 180,
      attemptCount: 1,
      status: "acked" as const,
    };

    await store.appendPendingMutation(local);
    await store.updatePendingMutation(acked);
    store.close();

    const reopened = await factory.open("memory-doc");
    await expect(reopened.listPendingMutations()).resolves.toEqual([]);
    await expect(reopened.listMutationJournalEntries()).resolves.toEqual([acked]);
  });
});
