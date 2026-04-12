import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../engine.js";

describe("engine media metadata", () => {
  it("normalizes, round-trips, and clones image and shape metadata", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "media-spec" });
    await engine.ready();
    engine.createSheet("Dashboard");

    engine.setImage({
      id: " Revenue Image ",
      sheetName: "Dashboard",
      address: "c3",
      sourceUrl: "https://example.com/revenue.png",
      rows: 9,
      cols: 6,
      altText: "Revenue overview",
    });
    engine.setShape({
      id: " callout-shape ",
      sheetName: "Dashboard",
      address: "h4",
      shapeType: "textBox",
      rows: 4,
      cols: 5,
      text: "Review",
      fillColor: "#ffeeaa",
      strokeColor: "#222222",
    });

    const image = engine.getImage("revenue image");
    const shape = engine.getShape("CALLOUT-SHAPE");
    expect(image).toEqual({
      id: "Revenue Image",
      sheetName: "Dashboard",
      address: "C3",
      sourceUrl: "https://example.com/revenue.png",
      rows: 9,
      cols: 6,
      altText: "Revenue overview",
    });
    expect(shape).toEqual({
      id: "callout-shape",
      sheetName: "Dashboard",
      address: "H4",
      shapeType: "textBox",
      rows: 4,
      cols: 5,
      text: "Review",
      fillColor: "#ffeeaa",
      strokeColor: "#222222",
    });

    if (!image || !shape) {
      throw new TypeError("Expected image and shape metadata");
    }
    image.address = "Z9";
    shape.text = "Changed";

    expect(engine.getImage("Revenue Image")?.address).toBe("C3");
    expect(engine.getShape("callout-shape")?.text).toBe("Review");

    const snapshot = engine.exportSnapshot();
    expect(snapshot.workbook.metadata?.images).toEqual([
      {
        id: "Revenue Image",
        sheetName: "Dashboard",
        address: "C3",
        sourceUrl: "https://example.com/revenue.png",
        rows: 9,
        cols: 6,
        altText: "Revenue overview",
      },
    ]);
    expect(snapshot.workbook.metadata?.shapes).toEqual([
      {
        id: "callout-shape",
        sheetName: "Dashboard",
        address: "H4",
        shapeType: "textBox",
        rows: 4,
        cols: 5,
        text: "Review",
        fillColor: "#ffeeaa",
        strokeColor: "#222222",
      },
    ]);

    const restored = new SpreadsheetEngine({ workbookName: "media-restored" });
    await restored.ready();
    restored.importSnapshot(snapshot);
    expect(restored.getImages()).toEqual(snapshot.workbook.metadata?.images);
    expect(restored.getShapes()).toEqual(snapshot.workbook.metadata?.shapes);
  });

  it("rewrites image and shape anchors across structural edits and removes invalidated media", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "media-structure" });
    await engine.ready();
    engine.createSheet("Dashboard");

    engine.setImage({
      id: "Revenue Image",
      sheetName: "Dashboard",
      address: "B2",
      sourceUrl: "https://example.com/revenue.png",
      rows: 9,
      cols: 6,
    });
    engine.setShape({
      id: "Review Callout",
      sheetName: "Dashboard",
      address: "E5",
      shapeType: "roundedRectangle",
      rows: 3,
      cols: 4,
      text: "Review",
    });

    engine.insertRows("Dashboard", 0, 1);
    engine.insertColumns("Dashboard", 1, 2);

    expect(engine.getImage("Revenue Image")).toEqual({
      id: "Revenue Image",
      sheetName: "Dashboard",
      address: "D3",
      sourceUrl: "https://example.com/revenue.png",
      rows: 9,
      cols: 6,
    });
    expect(engine.getShape("Review Callout")).toEqual({
      id: "Review Callout",
      sheetName: "Dashboard",
      address: "G6",
      shapeType: "roundedRectangle",
      rows: 3,
      cols: 4,
      text: "Review",
    });

    engine.deleteRows("Dashboard", 2, 2);
    engine.deleteColumns("Dashboard", 6, 1);

    expect(engine.getImage("Revenue Image")).toBeUndefined();
    expect(engine.getShape("Review Callout")).toBeUndefined();
    expect(engine.getImages()).toEqual([]);
    expect(engine.getShapes()).toEqual([]);
  });
});
