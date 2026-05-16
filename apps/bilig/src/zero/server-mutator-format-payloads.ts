import type {
  CellBorderSidePatch,
  CellDateStyle,
  CellHorizontalAlignment,
  CellNumberFormatInput,
  CellNumberFormatPreset,
  CellStyleAlignmentPatch,
  CellStyleBordersPatch,
  CellStyleFillPatch,
  CellStyleFontPatch,
  CellStylePatch,
  CellVerticalAlignment,
} from '@bilig/protocol'

export interface ServerCellBorderSidePatchInput {
  readonly style?: CellBorderSidePatch['style'] | undefined
  readonly weight?: CellBorderSidePatch['weight'] | undefined
  readonly color?: CellBorderSidePatch['color'] | undefined
}

export interface ServerCellStylePatchInput {
  readonly fill?:
    | {
        readonly backgroundColor?: string | null | undefined
      }
    | null
    | undefined
  readonly font?:
    | {
        readonly family?: string | null | undefined
        readonly size?: number | null | undefined
        readonly bold?: boolean | null | undefined
        readonly italic?: boolean | null | undefined
        readonly underline?: boolean | null | undefined
        readonly color?: string | null | undefined
      }
    | null
    | undefined
  readonly alignment?:
    | {
        readonly horizontal?: CellHorizontalAlignment | null | undefined
        readonly vertical?: CellVerticalAlignment | null | undefined
        readonly wrap?: boolean | null | undefined
        readonly indent?: number | null | undefined
        readonly shrinkToFit?: boolean | null | undefined
        readonly readingOrder?: number | null | undefined
        readonly textRotation?: number | null | undefined
        readonly justifyLastLine?: boolean | null | undefined
      }
    | null
    | undefined
  readonly borders?:
    | {
        readonly top?: ServerCellBorderSidePatchInput | null | undefined
        readonly right?: ServerCellBorderSidePatchInput | null | undefined
        readonly bottom?: ServerCellBorderSidePatchInput | null | undefined
        readonly left?: ServerCellBorderSidePatchInput | null | undefined
      }
    | null
    | undefined
}

export type ServerCellNumberFormatInput = string | ServerCellNumberFormatPresetInput

export interface ServerCellNumberFormatPresetInput {
  readonly kind: CellNumberFormatPreset['kind']
  readonly currency?: string | undefined
  readonly decimals?: number | undefined
  readonly useGrouping?: boolean | undefined
  readonly negativeStyle?: CellNumberFormatPreset['negativeStyle'] | undefined
  readonly zeroStyle?: CellNumberFormatPreset['zeroStyle'] | undefined
  readonly dateStyle?: CellDateStyle | undefined
}

function normalizeBorderSidePatch(patch: ServerCellBorderSidePatchInput | null): CellBorderSidePatch | null {
  if (patch === null) {
    return null
  }
  const normalized: CellBorderSidePatch = {}
  if (patch.style !== undefined) {
    normalized.style = patch.style
  }
  if (patch.weight !== undefined) {
    normalized.weight = patch.weight
  }
  if (patch.color !== undefined) {
    normalized.color = patch.color
  }
  return normalized
}

export function normalizeStylePatch(patch: ServerCellStylePatchInput): CellStylePatch {
  const normalized: CellStylePatch = {}

  if (patch.fill !== undefined) {
    if (patch.fill === null) {
      normalized.fill = null
    } else {
      const fill: CellStyleFillPatch = {}
      if (patch.fill.backgroundColor !== undefined) {
        fill.backgroundColor = patch.fill.backgroundColor
      }
      normalized.fill = fill
    }
  }
  if (patch.font !== undefined) {
    if (patch.font === null) {
      normalized.font = null
    } else {
      const font: CellStyleFontPatch = {}
      if (patch.font.family !== undefined) {
        font.family = patch.font.family
      }
      if (patch.font.size !== undefined) {
        font.size = patch.font.size
      }
      if (patch.font.bold !== undefined) {
        font.bold = patch.font.bold
      }
      if (patch.font.italic !== undefined) {
        font.italic = patch.font.italic
      }
      if (patch.font.underline !== undefined) {
        font.underline = patch.font.underline
      }
      if (patch.font.color !== undefined) {
        font.color = patch.font.color
      }
      normalized.font = font
    }
  }
  if (patch.alignment !== undefined) {
    if (patch.alignment === null) {
      normalized.alignment = null
    } else {
      const alignment: CellStyleAlignmentPatch = {}
      if (patch.alignment.horizontal !== undefined) {
        alignment.horizontal = patch.alignment.horizontal
      }
      if (patch.alignment.vertical !== undefined) {
        alignment.vertical = patch.alignment.vertical
      }
      if (patch.alignment.wrap !== undefined) {
        alignment.wrap = patch.alignment.wrap
      }
      if (patch.alignment.indent !== undefined) {
        alignment.indent = patch.alignment.indent
      }
      if (patch.alignment.shrinkToFit !== undefined) {
        alignment.shrinkToFit = patch.alignment.shrinkToFit
      }
      if (patch.alignment.readingOrder !== undefined) {
        alignment.readingOrder = patch.alignment.readingOrder
      }
      if (patch.alignment.textRotation !== undefined) {
        alignment.textRotation = patch.alignment.textRotation
      }
      if (patch.alignment.justifyLastLine !== undefined) {
        alignment.justifyLastLine = patch.alignment.justifyLastLine
      }
      normalized.alignment = alignment
    }
  }
  if (patch.borders !== undefined) {
    if (patch.borders === null) {
      normalized.borders = null
    } else {
      const borders: CellStyleBordersPatch = {}
      if (patch.borders.top !== undefined) {
        borders.top = normalizeBorderSidePatch(patch.borders.top)
      }
      if (patch.borders.right !== undefined) {
        borders.right = normalizeBorderSidePatch(patch.borders.right)
      }
      if (patch.borders.bottom !== undefined) {
        borders.bottom = normalizeBorderSidePatch(patch.borders.bottom)
      }
      if (patch.borders.left !== undefined) {
        borders.left = normalizeBorderSidePatch(patch.borders.left)
      }
      normalized.borders = borders
    }
  }

  return normalized
}

export function normalizeNumberFormatInput(format: ServerCellNumberFormatInput): CellNumberFormatInput {
  if (typeof format === 'string') {
    return format
  }

  const normalized: CellNumberFormatPreset = {
    kind: format.kind,
  }
  if (format.currency !== undefined) {
    normalized.currency = format.currency
  }
  if (format.decimals !== undefined) {
    normalized.decimals = format.decimals
  }
  if (format.useGrouping !== undefined) {
    normalized.useGrouping = format.useGrouping
  }
  if (format.negativeStyle !== undefined) {
    normalized.negativeStyle = format.negativeStyle
  }
  if (format.zeroStyle !== undefined) {
    normalized.zeroStyle = format.zeroStyle
  }
  if (format.dateStyle !== undefined) {
    normalized.dateStyle = format.dateStyle
  }
  return normalized
}
