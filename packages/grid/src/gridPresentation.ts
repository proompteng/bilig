import type { Rectangle } from "@glideapps/glide-data-grid";
import { isNumericEditorSeed } from "./gridKeyboard.js";

export function getOverlayStyle(isEditingCell: boolean, overlayBounds: Rectangle | undefined) {
  if (!isEditingCell || !overlayBounds) {
    return undefined;
  }
  return {
    height: overlayBounds.height + 2,
    left: overlayBounds.x - 1,
    position: "fixed" as const,
    top: overlayBounds.y - 1,
    width: overlayBounds.width + 2
  };
}

export function getEditorTextAlign(editorValue: string): "left" | "right" {
  return isNumericEditorSeed(editorValue) ? "right" : "left";
}

export function getGridTheme(variant: "playground" | "product") {
  return {
    accentColor: "#1f7a43",
    accentFg: "#ffffff",
    bgCell: "#ffffff",
    bgCellMedium: "#f3f5f7",
    bgHeader: "#f6f7f8",
    borderColor: "#d5d9de",
    cellHorizontalPadding: variant === "product" ? 8 : 10,
    cellVerticalPadding: variant === "product" ? 4 : 6,
    drilldownBorder: "#d5d9de",
    editorFontSize: variant === "product" ? "12px" : "13px",
    fontFamily: '"Aptos","Segoe UI","IBM Plex Sans",sans-serif',
    headerFontStyle: variant === "product"
      ? "600 11px Aptos, Segoe UI, IBM Plex Sans, sans-serif"
      : "600 12px Aptos, Segoe UI, IBM Plex Sans, sans-serif",
    horizontalBorderColor: "#e5e7eb",
    lineHeight: variant === "product" ? 1.2 : 1.3,
    textDark: "#101828",
    textHeader: "#344054",
    textLight: "#667085"
  };
}
