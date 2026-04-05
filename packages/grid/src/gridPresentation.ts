import type { RenderCellSnapshot } from "./gridCells.js";
import { isNumericEditorSeed } from "./gridKeyboard.js";
import type { Rectangle } from "./gridTypes.js";

export interface GridEditorPresentation {
  readonly backgroundColor: string;
  readonly color: string;
  readonly font: string;
  readonly fontSize: number;
  readonly underline: boolean;
}

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

export function getEditorPresentation(options: {
  renderCell: RenderCellSnapshot;
  fillColor?: string | null | undefined;
}): GridEditorPresentation {
  const { renderCell, fillColor } = options;
  return {
    backgroundColor: fillColor?.trim() ? fillColor : "#ffffff",
    color: renderCell.color,
    font: renderCell.font,
    fontSize: renderCell.fontSize,
    underline: renderCell.underline,
  };
}

export function getGridTheme(options?: {
  gpuSurfaceEnabled?: boolean;
  textSurfaceEnabled?: boolean;
}) {
  void options;
  return {
    accentColor: "#1f7a43",
    accentFg: "#ffffff",
    accentLight: "rgba(31, 122, 67, 0.14)",
    bgCell: "#ffffff",
    bgCellMedium: "#f8f9fa",
    bgHeader: "#f8f9fa",
    bgHeaderHasFocus: "#e6f4ea",
    bgHeaderHovered: "#f1f3f4",
    borderColor: "#dadce0",
    cellHorizontalPadding: 8,
    cellVerticalPadding: 4,
    drilldownBorder: "#dadce0",
    editorFontSize: "12px",
    fontFamily: '"JetBrainsMono Nerd Font","JetBrains Mono",monospace',
    headerFontStyle: '500 11px "JetBrainsMono Nerd Font", "JetBrains Mono", monospace',
    horizontalBorderColor: "#eceff1",
    lineHeight: 1.2,
    textDark: "#202124",
    textHeader: "#5f6368",
    textHeaderSelected: "#ffffff",
    textLight: "#80868b",
    textMedium: "#5f6368",
  };
}
