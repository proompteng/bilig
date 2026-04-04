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
    width: overlayBounds.width + 2,
  };
}

export function getEditorTextAlign(editorValue: string): "left" | "right" {
  return isNumericEditorSeed(editorValue) ? "right" : "left";
}

export function getGridTheme(options?: {
  gpuSurfaceEnabled?: boolean;
  textSurfaceEnabled?: boolean;
}) {
  const gpuSurfaceEnabled = options?.gpuSurfaceEnabled === true;
  const textSurfaceEnabled = options?.textSurfaceEnabled === true;
  return {
    accentColor: gpuSurfaceEnabled ? "rgba(31, 122, 67, 0)" : "#1f7a43",
    accentFg: "#ffffff",
    accentLight: gpuSurfaceEnabled ? "rgba(31, 122, 67, 0)" : "rgba(31, 122, 67, 0.14)",
    bgCell: gpuSurfaceEnabled ? "rgba(255, 255, 255, 0)" : "#ffffff",
    bgCellMedium: gpuSurfaceEnabled ? "rgba(255, 255, 255, 0)" : "#f8f9fa",
    bgHeader: gpuSurfaceEnabled ? "rgba(248, 249, 250, 0)" : "#f8f9fa",
    bgHeaderHasFocus: gpuSurfaceEnabled ? "rgba(230, 244, 234, 0)" : "#e6f4ea",
    bgHeaderHovered: gpuSurfaceEnabled ? "rgba(241, 243, 244, 0)" : "#f1f3f4",
    borderColor: gpuSurfaceEnabled ? "rgba(218, 220, 224, 0)" : "#dadce0",
    cellHorizontalPadding: 8,
    cellVerticalPadding: 4,
    drilldownBorder: gpuSurfaceEnabled ? "rgba(218, 220, 224, 0)" : "#dadce0",
    editorFontSize: "12px",
    fontFamily: '"JetBrainsMono Nerd Font","JetBrains Mono",monospace',
    headerFontStyle: '500 11px "JetBrainsMono Nerd Font", "JetBrains Mono", monospace',
    horizontalBorderColor: gpuSurfaceEnabled ? "rgba(236, 239, 241, 0)" : "#eceff1",
    lineHeight: 1.2,
    textDark: "#202124",
    textHeader: textSurfaceEnabled ? "rgba(95, 99, 104, 0)" : "#5f6368",
    textHeaderSelected: textSurfaceEnabled
      ? "rgba(31, 122, 67, 0)"
      : gpuSurfaceEnabled
        ? "#1f7a43"
        : "#ffffff",
    textLight: textSurfaceEnabled ? "rgba(128, 134, 139, 0)" : "#80868b",
    textMedium: textSurfaceEnabled ? "rgba(95, 99, 104, 0)" : "#5f6368",
  };
}
