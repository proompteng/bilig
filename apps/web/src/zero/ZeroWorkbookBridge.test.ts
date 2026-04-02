/* oxlint-disable @typescript-eslint/no-unsafe-type-assertion */
import { describe, expect, it, vi } from "vitest";
import type { TypedView, Zero } from "@rocicorp/zero";
import type { WorkerViewportCache } from "../viewport-cache.js";
import { ZeroWorkbookBridge } from "./ZeroWorkbookBridge.js";

function createTypedView<T>(initial: T) {
  let data = initial;
  const listeners = new Set<(value: T) => void>();
  return {
    get data() {
      return data;
    },
    addListener(listener: (value: T) => void) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    destroy() {},
    emit(next: T) {
      data = next;
      for (const listener of listeners) {
        listener(next);
      }
    },
  } as TypedView<T> & { emit(next: T): void };
}

describe("ZeroWorkbookBridge", () => {
  it("repaints subscribed viewports when authoritative cell_eval rows update", () => {
    const workbookView = createTypedView({
      id: "bilig-demo",
      name: "bilig-demo",
      headRevision: 1,
      calculatedRevision: 1,
    });
    const sheetListView = createTypedView([
      {
        workbookId: "bilig-demo",
        name: "Sheet1",
        sortOrder: 0,
      },
    ]);
    const styleRegistryView = createTypedView<readonly unknown[]>([]);
    const numberFormatRegistryView = createTypedView<readonly unknown[]>([]);

    const selectionSourceView = createTypedView({
      workbookId: "bilig-demo",
      sheetId: "Sheet1",
      address: "B2",
      rowNum: 1,
      colNum: 1,
      inputValue: "relay",
    });
    const selectionEvalView = createTypedView({
      workbookId: "bilig-demo",
      sheetId: "Sheet1",
      address: "B2",
      rowNum: 1,
      colNum: 1,
      value: { tag: 0 },
      flags: 0,
      version: 1,
    });

    const viewportInputView = createTypedView<readonly unknown[]>([]);
    const viewportEvalView = createTypedView<readonly unknown[]>([]);
    const viewportRowView = createTypedView<readonly unknown[]>([]);
    const viewportColumnView = createTypedView<readonly unknown[]>([]);

    // Sequence based on ZeroWorkbookBridge usage
    const views = [
      workbookView, // constructor: queries.workbook.get
      sheetListView, // constructor: queries.sheet.byWorkbook
      styleRegistryView, // constructor: queries.cellStyle.byWorkbook
      numberFormatRegistryView, // constructor: queries.numberFormat.byWorkbook
      selectionSourceView, // constructor -> setSelection: queries.cellInput.one
      selectionEvalView, // constructor -> setSelection: queries.cellEval.one
      viewportInputView, // subscribeViewport -> attachTile: queries.cellInput.tile
      viewportEvalView, // subscribeViewport -> attachTile: queries.cellEval.tile
      viewportRowView, // subscribeViewport -> attachTile: queries.sheetRow.tile
      viewportColumnView, // subscribeViewport -> attachTile: queries.sheetCol.tile
    ];
    let index = 0;

    const zero = {
      materialize(_query: unknown) {
        const view = views[index];
        index += 1;
        return view ?? createTypedView([]);
      },
    } as unknown as Zero;

    const appliedPatches = new Array<{ styles: readonly { id: string }[] }>();
    const cache = {
      setKnownSheets: vi.fn(),
      peekCell: vi.fn(() => undefined),
      subscribeCells: vi.fn(() => () => {}),
      applyViewportPatch: vi.fn((patch) => {
        appliedPatches.push({ styles: patch.styles });
        return patch.cells.map((cell: { col: number; row: number }) => ({
          cell: [cell.col, cell.row] as const,
        }));
      }),
    } as unknown as WorkerViewportCache;

    const bridge = new ZeroWorkbookBridge(zero, "bilig-demo", cache, () => {});
    const listener = vi.fn();
    const unsubscribe = bridge.subscribeViewport(
      "Sheet1",
      {
        rowStart: 1,
        rowEnd: 2,
        colStart: 1,
        colEnd: 2,
      },
      listener,
    );

    expect(listener).toHaveBeenCalledTimes(1);
    expect(appliedPatches.at(-1)?.styles).toEqual([]);
    const initialPatchCount = appliedPatches.length;

    // Emit to the viewportEvalView since that's what subscribeViewport listens to via tile notification
    viewportEvalView.emit([
      {
        workbookId: "bilig-demo",
        sheetId: "Sheet1",
        address: "B2",
        rowNum: 1,
        colNum: 1,
        value: { tag: 0 },
        flags: 0,
        version: 2,
        styleId: "style-border",
        styleJson: {
          id: "style-border",
          fill: { backgroundColor: "#dbeafe" },
        },
      },
    ]);

    expect(listener).toHaveBeenCalledTimes(2);
    expect(
      appliedPatches
        .slice(initialPatchCount)
        .some((patch) => patch.styles.some((style) => style.id === "style-border")),
    ).toBe(true);

    unsubscribe();
    bridge.dispose();
  });
});
