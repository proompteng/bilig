/* oxlint-disable @typescript-eslint/no-unsafe-type-assertion */
import { describe, expect, it, vi } from "vitest";
import { ValueTag } from "@bilig/protocol";
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

  it("repaints visible viewports when selected-cell materialization updates the cache", () => {
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
    const styleRegistryView = createTypedView<readonly unknown[]>([
      {
        workbookId: "bilig-demo",
        styleId: "style-fill",
        styleJson: {
          id: "style-fill",
          fill: { backgroundColor: "#c9daf8" },
        },
      },
    ]);
    const numberFormatRegistryView = createTypedView<readonly unknown[]>([]);

    const selectionSourceView = createTypedView<readonly unknown[]>([]);
    const selectionEvalView = createTypedView<readonly unknown[]>([]);

    const viewportInputView = createTypedView<readonly unknown[]>([]);
    const viewportEvalView = createTypedView<readonly unknown[]>([]);
    const viewportRowView = createTypedView<readonly unknown[]>([]);
    const viewportColumnView = createTypedView<readonly unknown[]>([]);

    const views = [
      workbookView,
      sheetListView,
      styleRegistryView,
      numberFormatRegistryView,
      selectionSourceView,
      selectionEvalView,
      viewportInputView,
      viewportEvalView,
      viewportRowView,
      viewportColumnView,
    ];
    let index = 0;

    const zero = {
      materialize(_query: unknown) {
        const view = views[index];
        index += 1;
        return view ?? createTypedView([]);
      },
    } as unknown as Zero;

    const applyViewportPatch = vi.fn((patch) =>
      patch.cells.map((cell: { col: number; row: number }) => ({
        cell: [cell.col, cell.row] as const,
      })),
    );
    const cache = {
      setKnownSheets: vi.fn(),
      peekCell: vi.fn(() => undefined),
      subscribeCells: vi.fn(() => () => {}),
      applyViewportPatch,
    } as unknown as WorkerViewportCache;

    const bridge = new ZeroWorkbookBridge(zero, "bilig-demo", cache, () => {});
    const listener = vi.fn();
    const unsubscribe = bridge.subscribeViewport(
      "Sheet1",
      {
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      listener,
    );

    expect(listener).toHaveBeenCalledTimes(1);

    selectionEvalView.emit({
      workbookId: "bilig-demo",
      sheetId: "Sheet1",
      address: "A1",
      rowNum: 0,
      colNum: 0,
      value: { tag: 0 },
      flags: 0,
      version: 0,
      styleId: "style-fill",
      styleJson: {
        id: "style-fill",
        fill: { backgroundColor: "#c9daf8" },
      },
    });

    expect(listener).toHaveBeenCalledTimes(2);
    expect(listener).toHaveBeenLastCalledWith([{ cell: [0, 0] }]);
    expect(applyViewportPatch).toHaveBeenLastCalledWith(
      expect.objectContaining({
        styles: [
          expect.objectContaining({
            id: "style-fill",
            fill: { backgroundColor: "#c9daf8" },
          }),
        ],
      }),
    );

    unsubscribe();
    bridge.dispose();
  });

  it("hydrates authoritative selected-cell data even when cache already has a stale cell", () => {
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
    const initialSelectionSourceView = createTypedView<unknown>(undefined);
    const initialSelectionEvalView = createTypedView<unknown>(undefined);
    const selectionSourceView = createTypedView<unknown>(undefined);
    const selectionEvalView = createTypedView<unknown>(undefined);

    const views = [
      workbookView,
      sheetListView,
      styleRegistryView,
      numberFormatRegistryView,
      initialSelectionSourceView,
      initialSelectionEvalView,
      selectionSourceView,
      selectionEvalView,
    ];
    let index = 0;

    const zero = {
      materialize(_query: unknown) {
        const view = views[index];
        index += 1;
        return view ?? createTypedView(undefined);
      },
    } as unknown as Zero;

    const cache = {
      setKnownSheets: vi.fn(),
      peekCell: vi.fn(() => ({
        sheetName: "Sheet1",
        address: "G7",
        value: { tag: ValueTag.Number, value: 8 },
        input: "8",
        flags: 0,
        version: 1,
      })),
      subscribeCells: vi.fn(() => () => {}),
      applyViewportPatch: vi.fn(() => []),
    } as unknown as WorkerViewportCache;

    const bridge = new ZeroWorkbookBridge(zero, "bilig-demo", cache, () => {});
    bridge.setSelection("Sheet1", "G7");
    const listener = vi.fn();
    const unsubscribe = bridge.subscribeSelectedCell(listener);

    selectionSourceView.emit({
      workbookId: "bilig-demo",
      sheetId: "Sheet1",
      address: "G7",
      rowNum: 6,
      colNum: 6,
      inputValue: "=F7*2",
      editorText: "=F7*2",
      formula: "F7*2",
    });
    selectionEvalView.emit({
      workbookId: "bilig-demo",
      sheetId: "Sheet1",
      address: "G7",
      rowNum: 6,
      colNum: 6,
      value: { tag: ValueTag.Number, value: 8 },
      flags: 0,
      version: 2,
    });

    expect(listener).toHaveBeenLastCalledWith(
      expect.objectContaining({
        address: "G7",
        formula: "F7*2",
        input: "=F7*2",
        value: { tag: ValueTag.Number, value: 8 },
      }),
    );

    unsubscribe();
    bridge.dispose();
  });
});
