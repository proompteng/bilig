import { describe, expect, it } from "vitest";
import { formatAddress } from "@bilig/formula";
import { ValueTag } from "@bilig/protocol";
import { WorkerViewportTileStore, listViewportTileBounds } from "../worker-viewport-tile-store.js";

function createViewportValue(
  sheetId: number,
  sheetName: string,
  input: {
    rowStart: number;
    rowEnd: number;
    colStart: number;
    colEnd: number;
  },
) {
  const address = formatAddress(input.rowStart, input.colStart);
  return {
    sheetId,
    sheetName,
    cells: [
      {
        row: input.rowStart,
        col: input.colStart,
        snapshot: {
          sheetName,
          address,
          value: { tag: ValueTag.Number, value: input.rowStart * 1000 + input.colStart },
          flags: 0,
          version: 1,
        },
      },
    ],
    rowAxisEntries: [
      {
        index: input.rowStart,
        id: `${sheetName}:row:${String(input.rowStart)}`,
        size: 22,
        hidden: false,
      },
    ],
    columnAxisEntries: [
      {
        index: input.colStart,
        id: `${sheetName}:col:${String(input.colStart)}`,
        size: 104,
        hidden: false,
      },
    ],
    styles: [{ id: "style-0" }],
  };
}

describe("WorkerViewportTileStore", () => {
  it("covers wide viewports with 128x32 tile bounds", () => {
    expect(
      listViewportTileBounds({
        rowStart: 0,
        rowEnd: 40,
        colStart: 0,
        colEnd: 140,
      }),
    ).toEqual([
      { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
      { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 },
      { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
    ]);
  });

  it("reuses clean cached tiles across overlapping viewport reads", () => {
    const tileStore = new WorkerViewportTileStore();
    const reads: Array<{ sheetName: string; viewport: Record<string, number> }> = [];
    const localStore = {
      readViewportProjection(sheetName: string, viewport: Record<string, number>) {
        reads.push({ sheetName, viewport: { ...viewport } });
        return createViewportValue(7, sheetName, {
          rowStart: viewport.rowStart,
          rowEnd: viewport.rowEnd,
          colStart: viewport.colStart,
          colEnd: viewport.colEnd,
        });
      },
    };

    const first = tileStore.readViewport({
      localStore,
      sheetName: "Sheet1",
      viewport: {
        rowStart: 0,
        rowEnd: 40,
        colStart: 0,
        colEnd: 140,
      },
    });

    expect(first?.cells).toHaveLength(4);
    expect(reads).toHaveLength(4);

    const second = tileStore.readViewport({
      localStore,
      sheetName: "Sheet1",
      viewport: {
        rowStart: 0,
        rowEnd: 12,
        colStart: 0,
        colEnd: 12,
      },
    });

    expect(second?.sheetId).toBe(7);
    expect(second?.cells).toHaveLength(1);
    expect(reads).toHaveLength(4);
  });

  it("invalidates only the targeted sheet tiles", () => {
    const tileStore = new WorkerViewportTileStore();
    const reads: string[] = [];
    const localStore = {
      readViewportProjection(sheetName: string, viewport: Record<string, number>) {
        reads.push(`${sheetName}:${viewport.rowStart}:${viewport.colStart}`);
        return createViewportValue(sheetName === "Sheet1" ? 1 : 2, sheetName, {
          rowStart: viewport.rowStart,
          rowEnd: viewport.rowEnd,
          colStart: viewport.colStart,
          colEnd: viewport.colEnd,
        });
      },
    };

    tileStore.readViewport({
      localStore,
      sheetName: "Sheet1",
      viewport: { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
    });
    tileStore.readViewport({
      localStore,
      sheetName: "Sheet2",
      viewport: { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
    });
    tileStore.invalidateSheet(1);

    tileStore.readViewport({
      localStore,
      sheetName: "Sheet2",
      viewport: { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
    });
    tileStore.readViewport({
      localStore,
      sheetName: "Sheet1",
      viewport: { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
    });

    expect(reads).toEqual(["Sheet1:0:0", "Sheet2:0:0", "Sheet1:0:0"]);
  });
});
