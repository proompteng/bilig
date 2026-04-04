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

  test("keeps the product palette stable when the GPU surface is active", () => {
    const gpuTheme = getGridTheme({ gpuSurfaceEnabled: true });

    expect(gpuTheme.accentColor).toBe("#1f7a43");
    expect(gpuTheme.bgCell).toBe("#ffffff");
    expect(gpuTheme.bgCellMedium).toBe("#f8f9fa");
    expect(gpuTheme.accentLight).toBe("rgba(31, 122, 67, 0.14)");
    expect(gpuTheme.borderColor).toBe("#dadce0");
    expect(gpuTheme.bgHeader).toBe("#f8f9fa");
    expect(gpuTheme.bgHeaderHasFocus).toBe("#e6f4ea");
    expect(gpuTheme.textHeaderSelected).toBe("#ffffff");
  });

  test("keeps header text tokens stable when the text overlay surface is active", () => {
    const textSurfaceTheme = getGridTheme({ textSurfaceEnabled: true });

    expect(textSurfaceTheme.textHeader).toBe("#5f6368");
    expect(textSurfaceTheme.textHeaderSelected).toBe("#ffffff");
    expect(textSurfaceTheme.textMedium).toBe("#5f6368");
    expect(textSurfaceTheme.textLight).toBe("#80868b");
  });
});
