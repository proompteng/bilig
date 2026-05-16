export interface CellStyleFillSnapshot {
  backgroundColor: string
}

export interface CellStyleFontSnapshot {
  family?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
}

export const CELL_HORIZONTAL_ALIGNMENT_VALUES = [
  'general',
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
] as const
export const CELL_VERTICAL_ALIGNMENT_VALUES = ['top', 'middle', 'bottom', 'justify', 'distributed'] as const
export const CELL_BORDER_STYLE_VALUES = ['solid', 'dashed', 'dotted', 'double'] as const
export const CELL_BORDER_WEIGHT_VALUES = ['thin', 'medium', 'thick'] as const

export type CellHorizontalAlignment = (typeof CELL_HORIZONTAL_ALIGNMENT_VALUES)[number]
export type CellVerticalAlignment = (typeof CELL_VERTICAL_ALIGNMENT_VALUES)[number]
export type CellBorderStyle = (typeof CELL_BORDER_STYLE_VALUES)[number]
export type CellBorderWeight = (typeof CELL_BORDER_WEIGHT_VALUES)[number]

export interface CellStyleAlignmentSnapshot {
  horizontal?: CellHorizontalAlignment
  vertical?: CellVerticalAlignment
  wrap?: boolean
  indent?: number
  shrinkToFit?: boolean
  readingOrder?: number
  textRotation?: number
  justifyLastLine?: boolean
}

export interface CellBorderSideSnapshot {
  style: CellBorderStyle
  weight: CellBorderWeight
  color: string
}

export interface CellStyleBordersSnapshot {
  top?: CellBorderSideSnapshot
  right?: CellBorderSideSnapshot
  bottom?: CellBorderSideSnapshot
  left?: CellBorderSideSnapshot
}

export interface CellStyleProtectionSnapshot {
  locked?: boolean
  hidden?: boolean
}

export interface CellStyleRecord {
  id: string
  fill?: CellStyleFillSnapshot
  font?: CellStyleFontSnapshot
  alignment?: CellStyleAlignmentSnapshot
  borders?: CellStyleBordersSnapshot
  protection?: CellStyleProtectionSnapshot
}

export interface CellStyleFillPatch {
  backgroundColor?: string | null
}

export interface CellStyleFontPatch {
  family?: string | null
  size?: number | null
  bold?: boolean | null
  italic?: boolean | null
  underline?: boolean | null
  color?: string | null
}

export interface CellStyleAlignmentPatch {
  horizontal?: CellHorizontalAlignment | null
  vertical?: CellVerticalAlignment | null
  wrap?: boolean | null
  indent?: number | null
  shrinkToFit?: boolean | null
  readingOrder?: number | null
  textRotation?: number | null
  justifyLastLine?: boolean | null
}

export interface CellBorderSidePatch {
  style?: CellBorderStyle | null
  weight?: CellBorderWeight | null
  color?: string | null
}

export interface CellStyleBordersPatch {
  top?: CellBorderSidePatch | null
  right?: CellBorderSidePatch | null
  bottom?: CellBorderSidePatch | null
  left?: CellBorderSidePatch | null
}

export interface CellStylePatch {
  fill?: CellStyleFillPatch | null
  font?: CellStyleFontPatch | null
  alignment?: CellStyleAlignmentPatch | null
  borders?: CellStyleBordersPatch | null
}

export const CELL_STYLE_FIELD_VALUES = [
  'backgroundColor',
  'fontFamily',
  'fontSize',
  'fontBold',
  'fontItalic',
  'fontUnderline',
  'fontColor',
  'alignmentHorizontal',
  'alignmentVertical',
  'alignmentWrap',
  'alignmentIndent',
  'alignmentShrinkToFit',
  'alignmentReadingOrder',
  'alignmentTextRotation',
  'alignmentJustifyLastLine',
  'borderTop',
  'borderRight',
  'borderBottom',
  'borderLeft',
] as const

export type CellStyleField = (typeof CELL_STYLE_FIELD_VALUES)[number]

export const CELL_NUMBER_FORMAT_KIND_VALUES = [
  'general',
  'number',
  'currency',
  'accounting',
  'percent',
  'date',
  'time',
  'datetime',
  'text',
] as const
export const CELL_NUMBER_NEGATIVE_STYLE_VALUES = ['minus', 'parentheses'] as const
export const CELL_NUMBER_ZERO_STYLE_VALUES = ['zero', 'dash'] as const
export const CELL_DATE_STYLE_VALUES = ['short', 'iso'] as const

export type CellNumberFormatKind = (typeof CELL_NUMBER_FORMAT_KIND_VALUES)[number]
export type CellNumberNegativeStyle = (typeof CELL_NUMBER_NEGATIVE_STYLE_VALUES)[number]
export type CellNumberZeroStyle = (typeof CELL_NUMBER_ZERO_STYLE_VALUES)[number]
export type CellDateStyle = (typeof CELL_DATE_STYLE_VALUES)[number]

export interface CellNumberFormatPreset {
  kind: CellNumberFormatKind
  currency?: string
  decimals?: number
  useGrouping?: boolean
  negativeStyle?: CellNumberNegativeStyle
  zeroStyle?: CellNumberZeroStyle
  dateStyle?: CellDateStyle
}

export type CellNumberFormatInput = string | CellNumberFormatPreset

export interface CellNumberFormatRecord {
  id: string
  code: string
  kind: CellNumberFormatKind
}
