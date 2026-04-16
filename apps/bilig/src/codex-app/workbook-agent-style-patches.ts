import type { CellStylePatch } from '@bilig/protocol'
import { setRangeStyleArgsSchema } from '@bilig/zero-sync'
import { z } from 'zod'

const alignmentHorizontalSchema = z.enum(['general', 'left', 'center', 'right'])
const alignmentVerticalSchema = z.enum(['top', 'middle', 'bottom'])
const borderStyleSchema = z.enum(['solid', 'dashed', 'dotted', 'double'])
const borderWeightSchema = z.enum(['thin', 'medium', 'thick'])

const styleBorderSidePatchSchema = z
  .object({
    style: borderStyleSchema.nullable().optional(),
    weight: borderWeightSchema.nullable().optional(),
    color: z.string().nullable().optional(),
  })
  .nullable()
  .optional()

export const workbookAgentStylePatchSchema = setRangeStyleArgsSchema.shape.patch.extend({
  backgroundColor: z.string().nullable().optional(),
  fillColor: z.string().nullable().optional(),
  fontFamily: z.string().nullable().optional(),
  fontSize: z.number().nullable().optional(),
  fontWeight: z.union([z.boolean(), z.number(), z.string()]).nullable().optional(),
  fontStyle: z.string().trim().min(1).nullable().optional(),
  fontUnderline: z.boolean().nullable().optional(),
  fontColor: z.string().nullable().optional(),
  textColor: z.string().nullable().optional(),
  horizontalAlignment: alignmentHorizontalSchema.nullable().optional(),
  verticalAlignment: alignmentVerticalSchema.nullable().optional(),
  wrap: z.boolean().nullable().optional(),
  indent: z.number().nullable().optional(),
  border: styleBorderSidePatchSchema,
  borderTop: styleBorderSidePatchSchema,
  borderRight: styleBorderSidePatchSchema,
  borderBottom: styleBorderSidePatchSchema,
  borderLeft: styleBorderSidePatchSchema,
})

export type WorkbookAgentStylePatchInput = z.infer<typeof workbookAgentStylePatchSchema>

export const workbookAgentStylePatchJsonSchema = {
  type: 'object',
  additionalProperties: false,
  properties: {
    fill: { type: 'object' },
    font: { type: 'object' },
    alignment: { type: 'object' },
    borders: { type: 'object' },
    backgroundColor: { type: 'string' },
    fillColor: { type: 'string' },
    fontFamily: { type: 'string' },
    fontSize: { type: 'number' },
    fontWeight: { type: 'string' },
    fontStyle: { type: 'string' },
    fontUnderline: { type: 'boolean' },
    fontColor: { type: 'string' },
    textColor: { type: 'string' },
    horizontalAlignment: { type: 'string', enum: ['general', 'left', 'center', 'right'] },
    verticalAlignment: { type: 'string', enum: ['top', 'middle', 'bottom'] },
    wrap: { type: 'boolean' },
    indent: { type: 'number' },
    border: { type: 'object' },
    borderTop: { type: 'object' },
    borderRight: { type: 'object' },
    borderBottom: { type: 'object' },
    borderLeft: { type: 'object' },
  },
}

export function normalizeWorkbookAgentStylePatch(patch: WorkbookAgentStylePatchInput): CellStylePatch {
  const normalized: CellStylePatch = {}

  if (patch.fill === null) {
    normalized.fill = null
  } else {
    const fill: NonNullable<CellStylePatch['fill']> = {}
    const backgroundColor = firstDefined(patch.fill?.backgroundColor, patch.fillColor, patch.backgroundColor)
    if (backgroundColor !== undefined) {
      fill.backgroundColor = backgroundColor
    }
    if (Object.keys(fill).length > 0) {
      normalized.fill = fill
    }
  }

  if (patch.font === null) {
    normalized.font = null
  } else {
    const font: NonNullable<CellStylePatch['font']> = {}
    const family = firstDefined(patch.font?.family, patch.fontFamily)
    if (family !== undefined) {
      font.family = family
    }
    const size = firstDefined(patch.font?.size, patch.fontSize)
    if (size !== undefined) {
      font.size = size
    }
    const bold = firstDefined(patch.font?.bold, normalizeFontWeightAlias(patch.fontWeight))
    if (bold !== undefined) {
      font.bold = bold
    }
    const italic = firstDefined(patch.font?.italic, normalizeFontStyleAlias(patch.fontStyle))
    if (italic !== undefined) {
      font.italic = italic
    }
    const underline = firstDefined(patch.font?.underline, patch.fontUnderline)
    if (underline !== undefined) {
      font.underline = underline
    }
    const color = firstDefined(patch.font?.color, patch.textColor, patch.fontColor)
    if (color !== undefined) {
      font.color = color
    }
    if (Object.keys(font).length > 0) {
      normalized.font = font
    }
  }

  if (patch.alignment === null) {
    normalized.alignment = null
  } else {
    const alignment: NonNullable<CellStylePatch['alignment']> = {}
    const horizontal = firstDefined(patch.alignment?.horizontal, patch.horizontalAlignment)
    if (horizontal !== undefined) {
      alignment.horizontal = horizontal
    }
    const vertical = firstDefined(patch.alignment?.vertical, patch.verticalAlignment)
    if (vertical !== undefined) {
      alignment.vertical = vertical
    }
    const wrap = firstDefined(patch.alignment?.wrap, patch.wrap)
    if (wrap !== undefined) {
      alignment.wrap = wrap
    }
    const indent = firstDefined(patch.alignment?.indent, patch.indent)
    if (indent !== undefined) {
      alignment.indent = indent
    }
    if (Object.keys(alignment).length > 0) {
      normalized.alignment = alignment
    }
  }

  if (patch.borders === null) {
    normalized.borders = null
  } else {
    const borders: NonNullable<CellStylePatch['borders']> = {}
    const borderAliases = {
      top: patch.borderTop,
      right: patch.borderRight,
      bottom: patch.borderBottom,
      left: patch.borderLeft,
    } as const
    for (const sideName of ['top', 'right', 'bottom', 'left'] as const) {
      const side = firstDefined(
        normalizeBorderSideAlias(patch.borders?.[sideName]),
        normalizeBorderSideAlias(borderAliases[sideName]),
        normalizeBorderSideAlias(patch.border),
      )
      if (side !== undefined) {
        borders[sideName] = side
      }
    }
    if (Object.values(borders).some((value) => value !== undefined)) {
      normalized.borders = borders
    }
  }

  return normalized
}

export function workbookAgentStylePatchHasChanges(patch: CellStylePatch): boolean {
  return (
    styleSectionHasChanges(patch.fill) ||
    styleSectionHasChanges(patch.font) ||
    styleSectionHasChanges(patch.alignment) ||
    styleBordersHaveChanges(patch.borders)
  )
}

function normalizeFontWeightAlias(value: WorkbookAgentStylePatchInput['fontWeight']): boolean | null | undefined {
  if (value === undefined || value === null || typeof value === 'boolean') {
    return value
  }
  if (typeof value === 'number') {
    return value >= 600
  }
  const normalized = value.trim().toLowerCase()
  if (normalized === 'bold' || normalized === 'bolder') {
    return true
  }
  if (normalized === 'semibold' || normalized === 'semi-bold' || normalized === 'demibold' || normalized === 'demi-bold') {
    return true
  }
  if (normalized === 'normal' || normalized === 'lighter') {
    return false
  }
  const numeric = Number(normalized)
  if (!Number.isNaN(numeric)) {
    return numeric >= 600
  }
  return undefined
}

function normalizeFontStyleAlias(value: WorkbookAgentStylePatchInput['fontStyle']): boolean | null | undefined {
  if (value === undefined || value === null) {
    return value
  }
  const normalized = value.trim().toLowerCase()
  if (normalized === 'italic') {
    return true
  }
  if (normalized === 'normal' || normalized === 'roman') {
    return false
  }
  return undefined
}

function normalizeBorderSideAlias(
  value: WorkbookAgentStylePatchInput['border'],
): NonNullable<NonNullable<CellStylePatch['borders']>['top']> | null | undefined {
  if (value === undefined || value === null) {
    return value
  }
  const side: NonNullable<NonNullable<CellStylePatch['borders']>['top']> = {}
  if (value.style !== undefined) {
    side.style = value.style
  }
  if (value.weight !== undefined) {
    side.weight = value.weight
  }
  if (value.color !== undefined) {
    side.color = value.color
  }
  return Object.keys(side).length > 0 ? side : undefined
}

function styleSectionHasChanges(value: object | null | undefined): boolean {
  return value === null || (value !== undefined && Object.keys(value).length > 0)
}

function styleBordersHaveChanges(value: CellStylePatch['borders']): boolean {
  if (value === null) {
    return true
  }
  if (value === undefined) {
    return false
  }
  return (['top', 'right', 'bottom', 'left'] as const).some((sideName) => styleSectionHasChanges(value[sideName]))
}

function firstDefined<T>(...values: readonly (T | undefined)[]): T | undefined {
  for (const value of values) {
    if (value !== undefined) {
      return value
    }
  }
  return undefined
}
