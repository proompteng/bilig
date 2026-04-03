import { describe, expect, test } from "vitest";
import { getEditorTextAlign, getGridTheme, getOverlayStyle } from "../gridPresentation.js";

describe("gridPresentation", () => {
  test("matches editor alignment to numeric and text seeds", () => {
    expect(getEditorTextAlign("123")).toBe("right");
    expect(getEditorTextAlign("-45.6")).toBe("right");
    expect(getEditorTextAlign("hello")).toBe("left");
  });

  test("derives overlay style from overlay bounds only when editing", () => {
    expect(getOverlayStyle(false, { x: 10, y: 20, width: 30, height: 40 })).toBeUndefined();
    expect(getOverlayStyle(true, undefined)).toBeUndefined();
    expect(getOverlayStyle(true, { x: 10, y: 20, width: 30, height: 40 })).toEqual({
      height: 42,
      left: 9,
      position: "fixed",
      top: 19,
      width: 32,
    });
  });

  test("returns the dense product theme values", () => {
    const productTheme = getGridTheme();

    expect(productTheme.accentColor).toBe("#1f7a43");
    expect(productTheme.accentLight).toBe("rgba(31, 122, 67, 0.14)");
    expect(productTheme.bgHeaderHasFocus).toBe("#e6f4ea");
    expect(productTheme.cellHorizontalPadding).toBe(8);
    expect(productTheme.cellVerticalPadding).toBe(4);
    expect(productTheme.editorFontSize).toBe("12px");
    expect(productTheme.fontFamily).toContain("JetBrainsMono Nerd Font");
    expect(productTheme.headerFontStyle).toContain("11px");
  });
});
