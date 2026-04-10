import { describe, expect, it } from "vitest";
import {
  getCellNumberFormat,
  getCellStyle,
  getRangeFormatId,
  getStyleId,
  internCellNumberFormat,
  internCellStyle,
  setFormatRanges,
  setStyleRanges,
  upsertCellNumberFormat,
  upsertCellStyle,
} from "../workbook-style-format-store.js";

function createCatalog() {
  return {
    cellStyles: new Map(),
    styleKeys: new Map(),
    cellNumberFormats: new Map(),
    numberFormatKeys: new Map(),
  };
}

describe("workbook style format store", () => {
  it("interns normalized styles and number formats", () => {
    const catalog = createCatalog();
    upsertCellStyle(catalog, { id: "style-0" }, () => {});
    upsertCellNumberFormat(catalog, { id: "format-0", code: "general" }, () => {});

    const style = internCellStyle(
      catalog,
      { fill: { backgroundColor: "#abc" }, font: { bold: true } },
      "style-0",
    );
    upsertCellStyle(catalog, style, () => {});

    expect(
      internCellStyle(
        catalog,
        { fill: { backgroundColor: "#aabbcc" }, font: { bold: true } },
        "style-0",
      ),
    ).toEqual(style);
    expect(getCellStyle(catalog, undefined, "style-0")?.id).toBe("style-0");

    const format = internCellNumberFormat(catalog, "$0.00", "format-0");
    upsertCellNumberFormat(catalog, format, () => {});

    expect(internCellNumberFormat(catalog, "$0.00", "format-0")).toEqual(format);
    expect(getCellNumberFormat(catalog, undefined, "format-0")?.id).toBe("format-0");
  });

  it("does not mutate style and format ranges when unknown ids are provided", () => {
    const catalog = createCatalog();
    const sheet = {
      styleRanges: [
        {
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
          styleId: "style-a",
        },
      ],
      formatRanges: [
        {
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
          formatId: "format-a",
        },
      ],
    };
    upsertCellStyle(catalog, { id: "style-a", font: { bold: true } }, () => {});
    upsertCellNumberFormat(catalog, { id: "format-a", code: "$0.00" }, () => {});

    expect(() =>
      setStyleRanges(catalog, sheet, [
        {
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "C3" },
          styleId: "style-missing",
        },
      ]),
    ).toThrow("Unknown cell style: style-missing");
    expect(() =>
      setFormatRanges(catalog, sheet, [
        {
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "C3" },
          formatId: "format-missing",
        },
      ]),
    ).toThrow("Unknown cell number format: format-missing");

    expect(getStyleId(sheet, 0, 0, "style-0")).toBe("style-a");
    expect(getRangeFormatId(sheet, 0, 0, "format-0")).toBe("format-a");
  });
});
