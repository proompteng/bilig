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
  it("repaints subscribed viewports when style rows arrive after a style-range update", () => {
    const workbookView = createTypedView({
      id: "bilig-demo",
      name: "bilig-demo",
      headRevision: 1,
      calculatedRevision: 1,
    });
    const sheetView = createTypedView([
      {
        workbookId: "bilig-demo",
        name: "Sheet1",
        sortOrder: 0,
      },
    ]);
    const stylesView = createTypedView<readonly unknown[]>([]);
    const numberFormatsView = createTypedView<readonly unknown[]>([]);
    const sourceView = createTypedView([
      {
        workbookId: "bilig-demo",
        sheetName: "Sheet1",
        address: "B2",
        rowNum: 1,
        colNum: 1,
        inputValue: "relay",
      },
    ]);
    const evalView = createTypedView([
      {
        workbookId: "bilig-demo",
        sheetName: "Sheet1",
        address: "B2",
        rowNum: 1,
        colNum: 1,
        value: { tag: 0 },
        flags: 0,
        version: 1,
      },
    ]);
    const rowView = createTypedView<readonly unknown[]>([]);
    const columnView = createTypedView<readonly unknown[]>([]);
    const styleRangeView = createTypedView([
      {
        id: "range-B2-C3",
        workbookId: "bilig-demo",
        sheetName: "Sheet1",
        startRow: 1,
        endRow: 2,
        startCol: 1,
        endCol: 2,
        styleId: "style-border",
        updatedAt: 1,
      },
    ]);
    const formatRangeView = createTypedView<readonly unknown[]>([]);

    const views = [
      workbookView,
      sheetView,
      stylesView,
      numberFormatsView,
      sourceView,
      evalView,
      rowView,
      columnView,
      styleRangeView,
      formatRangeView,
      sourceView,
      evalView,
      rowView,
      columnView,
      styleRangeView,
      formatRangeView,
    ] as const;

    let index = 0;
    const zero = {
      materialize() {
        const view = views[index];
        if (!view) {
          throw new Error(`No view available for materialize index ${index}`);
        }
        index += 1;
        return view;
      },
    } as unknown as Zero;

    const appliedPatches = new Array<{ styles: readonly { id: string }[] }>();
    const cache = {
      setKnownSheets: vi.fn(),
      peekCell: vi.fn(() => undefined),
      applyViewportPatch: vi.fn((patch) => {
        appliedPatches.push({ styles: patch.styles });
        return patch.cells.map((cell) => ({ cell: [cell.col, cell.row] as const }));
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
    expect(appliedPatches.at(-1)?.styles.every((style) => style.id === "style-0")).toBe(true);
    const initialPatchCount = appliedPatches.length;

    stylesView.emit([
      {
        workbookId: "bilig-demo",
        id: "style-border",
        recordJSON: {
          id: "style-border",
          fill: { backgroundColor: "#dbeafe" },
          borders: {
            top: { style: "solid", weight: "thin", color: "#111827" },
            right: { style: "solid", weight: "thin", color: "#111827" },
            bottom: { style: "solid", weight: "thin", color: "#111827" },
            left: { style: "solid", weight: "thin", color: "#111827" },
          },
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
