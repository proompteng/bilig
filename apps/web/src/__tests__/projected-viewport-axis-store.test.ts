import { describe, expect, it, vi } from "vitest";
import { ProjectedViewportAxisStore } from "../projected-viewport-axis-store.js";

describe("ProjectedViewportAxisStore", () => {
  it("tracks optimistic axis writes and sheet eviction independently of the cell cache", () => {
    const markSheetKnown = vi.fn();
    const notifyListeners = vi.fn();
    const axisStore = new ProjectedViewportAxisStore({
      markSheetKnown,
      notifyListeners,
    });

    axisStore.setColumnWidth("Sheet1", 0, 93);
    axisStore.setRowHidden("Sheet1", 1, true, 44);

    expect(axisStore.getColumnWidths("Sheet1")[0]).toBe(93);
    expect(axisStore.getColumnSizes("Sheet1")[0]).toBe(93);
    expect(axisStore.getRowHeights("Sheet1")[1]).toBe(0);
    expect(axisStore.getRowSizes("Sheet1")[1]).toBe(44);
    expect(axisStore.getHiddenRows("Sheet1")[1]).toBe(true);
    expect(markSheetKnown).toHaveBeenCalledWith("Sheet1");
    expect(notifyListeners).toHaveBeenCalledTimes(2);

    axisStore.dropSheets(["Sheet1"]);

    expect(axisStore.getColumnWidths("Sheet1")[0]).toBeUndefined();
    expect(axisStore.getHiddenRows("Sheet1")[1]).toBeUndefined();
  });

  it("skips listener notifications for identical axis writes", () => {
    const notifyListeners = vi.fn();
    const axisStore = new ProjectedViewportAxisStore({
      notifyListeners,
    });

    axisStore.setColumnHidden("Sheet1", 0, true, 68);
    axisStore.setColumnHidden("Sheet1", 0, true, 68);
    axisStore.ackColumnWidth("Sheet1", 0, 0);

    expect(notifyListeners).toHaveBeenCalledTimes(1);
    expect(axisStore.getHiddenColumns("Sheet1")[0]).toBe(true);
    expect(axisStore.getColumnWidths("Sheet1")[0]).toBe(0);
    expect(axisStore.getColumnSizes("Sheet1")[0]).toBe(68);
  });
});
