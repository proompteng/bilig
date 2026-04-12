import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../engine.js";

describe("engine chart metadata", () => {
  it("normalizes, round-trips, and clones chart metadata", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "charts-spec" });
    await engine.ready();
    engine.createSheet("Data");
    engine.createSheet("Dashboard");
    engine.setRangeValues({ sheetName: "Data", startAddress: "A1", endAddress: "B4" }, [
      ["Month", "Revenue"],
      ["Jan", 10],
      ["Feb", 15],
      ["Mar", 9],
    ]);

    engine.setChart({
      id: " Revenue Chart ",
      sheetName: "Dashboard",
      address: "c3",
      source: { sheetName: "Data", startAddress: "b4", endAddress: "a1" },
      chartType: "column",
      rows: 12,
      cols: 8,
      title: "Monthly revenue",
      seriesOrientation: "columns",
      firstRowAsHeaders: true,
      firstColumnAsLabels: true,
      legendPosition: "bottom",
    });

    const chart = engine.getChart("revenue chart");
    expect(chart).toEqual({
      id: "Revenue Chart",
      sheetName: "Dashboard",
      address: "C3",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      chartType: "column",
      rows: 12,
      cols: 8,
      title: "Monthly revenue",
      seriesOrientation: "columns",
      firstRowAsHeaders: true,
      firstColumnAsLabels: true,
      legendPosition: "bottom",
    });

    if (!chart) {
      throw new TypeError("Expected chart metadata");
    }
    chart.address = "Z9";

    expect(engine.getChart("Revenue Chart")).toEqual({
      id: "Revenue Chart",
      sheetName: "Dashboard",
      address: "C3",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      chartType: "column",
      rows: 12,
      cols: 8,
      title: "Monthly revenue",
      seriesOrientation: "columns",
      firstRowAsHeaders: true,
      firstColumnAsLabels: true,
      legendPosition: "bottom",
    });

    const snapshot = engine.exportSnapshot();
    expect(snapshot.workbook.metadata?.charts).toEqual([
      {
        id: "Revenue Chart",
        sheetName: "Dashboard",
        address: "C3",
        source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
        chartType: "column",
        rows: 12,
        cols: 8,
        title: "Monthly revenue",
        seriesOrientation: "columns",
        firstRowAsHeaders: true,
        firstColumnAsLabels: true,
        legendPosition: "bottom",
      },
    ]);

    const restored = new SpreadsheetEngine({ workbookName: "charts-restored" });
    await restored.ready();
    restored.importSnapshot(snapshot);
    expect(restored.getCharts()).toEqual(snapshot.workbook.metadata?.charts);
  });

  it("rewrites chart anchors and source ranges across structural edits and removes invalidated charts", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "charts-structure" });
    await engine.ready();
    engine.createSheet("Data");
    engine.createSheet("Dashboard");
    engine.setRangeValues({ sheetName: "Data", startAddress: "A1", endAddress: "B4" }, [
      ["Month", "Revenue"],
      ["Jan", 10],
      ["Feb", 15],
      ["Mar", 9],
    ]);
    engine.setChart({
      id: "Trend",
      sheetName: "Dashboard",
      address: "B2",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      chartType: "line",
      rows: 10,
      cols: 6,
    });

    engine.insertRows("Data", 0, 1);
    engine.insertColumns("Dashboard", 0, 1);

    expect(engine.getChart("Trend")).toEqual({
      id: "Trend",
      sheetName: "Dashboard",
      address: "C2",
      source: { sheetName: "Data", startAddress: "A2", endAddress: "B5" },
      chartType: "line",
      rows: 10,
      cols: 6,
    });

    engine.deleteRows("Data", 0, 5);
    expect(engine.getChart("Trend")).toBeUndefined();
    expect(engine.getCharts()).toEqual([]);
  });

  it("skips duplicate chart writes and returns correct delete booleans", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "charts-idempotent" });
    await engine.ready();
    engine.createSheet("Data");
    engine.createSheet("Dashboard");
    engine.setRangeValues({ sheetName: "Data", startAddress: "A1", endAddress: "B2" }, [
      ["Month", "Revenue"],
      ["Jan", 10],
    ]);

    engine.setChart({
      id: "Trend",
      sheetName: "Dashboard",
      address: "B2",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B2" },
      chartType: "line",
      rows: 3,
      cols: 4,
    });
    engine.setChart({
      id: "Trend",
      sheetName: "Dashboard",
      address: "B2",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B2" },
      chartType: "line",
      rows: 3,
      cols: 4,
    });

    expect(engine.getCharts()).toHaveLength(1);
    expect(engine.deleteChart("Missing")).toBe(false);
    expect(engine.deleteChart("Trend")).toBe(true);
    expect(engine.getCharts()).toEqual([]);
  });
});
