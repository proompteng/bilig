import type { RenderCellSnapshot } from './gridCells.js'
import { isNumericEditorSeed } from './gridKeyboard.js'
import type { Rectangle } from './gridTypes.js'
import { WORKBOOK_FONT_SANS, workbookThemeColors } from './workbookTheme.js'

export interface GridEditorPresentation {
  readonly backgroundColor: string
  readonly color: string
  readonly font: string
  readonly fontSize: number
  readonly underline: boolean
}

export function getOverlayStyle(isEditingCell: boolean, overlayBounds: Rectangle | undefined) {
  if (!isEditingCell || !overlayBounds) {
    return undefined
  }
  return {
    height: overlayBounds.height,
    left: overlayBounds.x,
    position: 'fixed' as const,
    top: overlayBounds.y,
    width: overlayBounds.width,
  }
}

export function getEditorTextAlign(editorValue: string): 'left' | 'right' {
  return isNumericEditorSeed(editorValue) ? 'right' : 'left'
}

export function getEditorPresentation(options: {
  renderCell: RenderCellSnapshot
  fillColor?: string | null | undefined
}): GridEditorPresentation {
  const { renderCell, fillColor } = options
  return {
    backgroundColor: fillColor?.trim() ? fillColor : workbookThemeColors.surface,
    color: renderCell.color,
    font: renderCell.font,
    fontSize: renderCell.fontSize,
    underline: renderCell.underline,
  }
}

export function getGridTheme(options?: { gpuSurfaceEnabled?: boolean; textSurfaceEnabled?: boolean }) {
  void options
  const fontFamily = WORKBOOK_FONT_SANS
  return {
    accentColor: workbookThemeColors.accent,
    accentFg: workbookThemeColors.surface,
    accentLight: workbookThemeColors.accentSoft,
    bgCell: workbookThemeColors.surface,
    bgCellMedium: workbookThemeColors.surfaceSubtle,
    bgHeader: workbookThemeColors.surfaceSubtle,
    bgHeaderHasFocus: workbookThemeColors.accentSoft,
    bgHeaderHovered: workbookThemeColors.muted,
    borderColor: workbookThemeColors.border,
    cellHorizontalPadding: 8,
    cellVerticalPadding: 4,
    drilldownBorder: workbookThemeColors.border,
    editorFontSize: '12px',
    fontFamily,
    headerFontStyle: `600 11px ${fontFamily}`,
    horizontalBorderColor: workbookThemeColors.gridBorder,
    lineHeight: 1.2,
    textDark: workbookThemeColors.text,
    textHeader: workbookThemeColors.textMuted,
    textHeaderSelected: workbookThemeColors.surface,
    textLight: workbookThemeColors.textSubtle,
    textMedium: workbookThemeColors.textMuted,
  }
}
