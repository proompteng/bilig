import type {
  CellRangeRef,
  CellStyleAlignmentPatch,
  CellStyleField,
  CellStyleFillPatch,
  CellStyleFontPatch,
  CellStylePatch,
  CellStyleRecord,
} from '@bilig/protocol'
import type { BorderPreset } from './workbook-toolbar.js'
import { isTextEntryTarget } from './worker-workbook-app-model.js'

export const BORDER_CLEAR_FIELDS: readonly CellStyleField[] = ['borderTop', 'borderRight', 'borderBottom', 'borderLeft'] as const

export const DEFAULT_BORDER_SIDE = {
  style: 'solid',
  weight: 'thin',
  color: '#111827',
} as const

const PENDING_STYLE_ID = '__bilig_pending_toolbar_style__'
const WORKBOOK_KEYBOARD_SCOPE_SELECTOR = '[data-workbook-keyboard-scope="true"]'

export interface OptimisticToolbarStyle {
  readonly rangeKey: string
  readonly patch: CellStylePatch
  readonly style: CellStyleRecord
}

export function mergeToolbarStylePatch(previous: CellStylePatch, next: CellStylePatch): CellStylePatch {
  const merged: CellStylePatch = {
    ...previous,
    ...next,
  }
  if (next.alignment !== undefined) {
    merged.alignment = next.alignment ? { ...previous.alignment, ...next.alignment } : null
  }
  if (next.borders !== undefined) {
    merged.borders = next.borders ? { ...previous.borders, ...next.borders } : null
  }
  if (next.fill !== undefined) {
    merged.fill = next.fill ? { ...previous.fill, ...next.fill } : null
  }
  if (next.font !== undefined) {
    merged.font = next.font ? { ...previous.font, ...next.font } : null
  }
  return merged
}

function patchValueMatches<T extends string | number | boolean>(actual: T | undefined, expected: T | null | undefined): boolean {
  if (expected === undefined) {
    return true
  }
  if (expected === null || expected === false) {
    return actual === undefined || actual === false
  }
  return actual === expected
}

export function shouldKeepWorkbookShortcutInsideTextEntry(target: EventTarget | null): boolean {
  if (!isTextEntryTarget(target)) {
    return false
  }
  return !(
    (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) &&
    target.dataset['testid'] === 'formula-input' &&
    target.dataset['formulaEditing'] === 'false'
  )
}

export function shouldRouteWorkbookShortcutToWorkbookScope(activeElement: Element | null = getActiveElement()): boolean {
  if (typeof document === 'undefined') {
    return false
  }
  if (activeElement === null || activeElement === document.body || activeElement === document.documentElement) {
    return true
  }
  return activeElement.closest(WORKBOOK_KEYBOARD_SCOPE_SELECTOR) !== null
}

export function queueWorkbookHistoryShortcut(callback: () => void): void {
  if (typeof queueMicrotask === 'function') {
    queueMicrotask(callback)
    return
  }
  setTimeout(callback, 0)
}

function getActiveElement(): Element | null {
  if (typeof document === 'undefined') {
    return null
  }
  return document.activeElement
}

function fillPatchMatches(selectedStyle: CellStyleRecord | undefined, patch: CellStyleFillPatch | null | undefined): boolean {
  return !patch || patchValueMatches(selectedStyle?.fill?.backgroundColor, patch.backgroundColor)
}

function fontPatchMatches(selectedStyle: CellStyleRecord | undefined, patch: CellStyleFontPatch | null | undefined): boolean {
  return (
    !patch ||
    (patchValueMatches(selectedStyle?.font?.family, patch.family) &&
      patchValueMatches(selectedStyle?.font?.size, patch.size) &&
      patchValueMatches(selectedStyle?.font?.bold, patch.bold) &&
      patchValueMatches(selectedStyle?.font?.italic, patch.italic) &&
      patchValueMatches(selectedStyle?.font?.underline, patch.underline) &&
      patchValueMatches(selectedStyle?.font?.color, patch.color))
  )
}

function alignmentPatchMatches(selectedStyle: CellStyleRecord | undefined, patch: CellStyleAlignmentPatch | null | undefined): boolean {
  return (
    !patch ||
    (patchValueMatches(selectedStyle?.alignment?.horizontal, patch.horizontal) &&
      patchValueMatches(selectedStyle?.alignment?.vertical, patch.vertical) &&
      patchValueMatches(selectedStyle?.alignment?.wrap, patch.wrap) &&
      patchValueMatches(selectedStyle?.alignment?.indent, patch.indent))
  )
}

function borderSidePatchMatches(
  actual: NonNullable<CellStyleRecord['borders']>['top'] | undefined,
  patch: NonNullable<CellStylePatch['borders']>['top'] | null | undefined,
): boolean {
  if (patch === undefined) {
    return true
  }
  if (patch === null) {
    return actual === undefined
  }
  return (
    patchValueMatches(actual?.style, patch.style) &&
    patchValueMatches(actual?.weight, patch.weight) &&
    patchValueMatches(actual?.color, patch.color)
  )
}

function borderPatchMatches(selectedStyle: CellStyleRecord | undefined, patch: CellStylePatch['borders'] | undefined): boolean {
  if (patch === undefined) {
    return true
  }
  if (patch === null) {
    return !hasAnyBorder(selectedStyle)
  }
  return (
    borderSidePatchMatches(selectedStyle?.borders?.top, patch.top) &&
    borderSidePatchMatches(selectedStyle?.borders?.right, patch.right) &&
    borderSidePatchMatches(selectedStyle?.borders?.bottom, patch.bottom) &&
    borderSidePatchMatches(selectedStyle?.borders?.left, patch.left)
  )
}

export function selectedStyleMatchesPatch(selectedStyle: CellStyleRecord | undefined, patch: CellStylePatch): boolean {
  return (
    alignmentPatchMatches(selectedStyle, patch.alignment) &&
    fillPatchMatches(selectedStyle, patch.fill) &&
    fontPatchMatches(selectedStyle, patch.font) &&
    borderPatchMatches(selectedStyle, patch.borders)
  )
}

export function hasAnyBorder(style: CellStyleRecord | undefined): boolean {
  const borders = style?.borders
  return Boolean(borders?.top || borders?.right || borders?.bottom || borders?.left)
}

export function cellRangeKey(range: CellRangeRef): string {
  return `${range.sheetName}:${range.startAddress}:${range.endAddress ?? range.startAddress}`
}

function cloneStyleForToolbar(style: CellStyleRecord | undefined): CellStyleRecord {
  return {
    id: style?.id ?? PENDING_STYLE_ID,
    ...(style?.fill ? { fill: { ...style.fill } } : {}),
    ...(style?.font ? { font: { ...style.font } } : {}),
    ...(style?.alignment ? { alignment: { ...style.alignment } } : {}),
    ...(style?.borders
      ? {
          borders: {
            ...(style.borders.top ? { top: { ...style.borders.top } } : {}),
            ...(style.borders.right ? { right: { ...style.borders.right } } : {}),
            ...(style.borders.bottom ? { bottom: { ...style.borders.bottom } } : {}),
            ...(style.borders.left ? { left: { ...style.borders.left } } : {}),
          },
        }
      : {}),
  }
}

function applyOptionalStyleField<T extends object, K extends keyof T>(target: T, key: K, value: T[K] | null | undefined): void {
  if (value === undefined) {
    return
  }
  if (value === null) {
    delete target[key]
    return
  }
  target[key] = value
}

export function applyToolbarStylePatch(style: CellStyleRecord | undefined, patch: CellStylePatch): CellStyleRecord {
  const next = cloneStyleForToolbar(style)
  if (patch.fill === null) {
    delete next.fill
  } else if (patch.fill !== undefined) {
    const backgroundColor = patch.fill.backgroundColor
    if (backgroundColor === null) {
      delete next.fill
    } else if (backgroundColor !== undefined) {
      next.fill = { backgroundColor }
    }
  }
  if (patch.font === null) {
    delete next.font
  } else if (patch.font) {
    const font = { ...next.font }
    applyOptionalStyleField(font, 'family', patch.font.family)
    applyOptionalStyleField(font, 'size', patch.font.size)
    applyOptionalStyleField(font, 'bold', patch.font.bold)
    applyOptionalStyleField(font, 'italic', patch.font.italic)
    applyOptionalStyleField(font, 'underline', patch.font.underline)
    applyOptionalStyleField(font, 'color', patch.font.color)
    if (Object.keys(font).length > 0) {
      next.font = font
    } else {
      delete next.font
    }
  }
  if (patch.alignment === null) {
    delete next.alignment
  } else if (patch.alignment) {
    const alignment = { ...next.alignment }
    applyOptionalStyleField(alignment, 'horizontal', patch.alignment.horizontal)
    applyOptionalStyleField(alignment, 'vertical', patch.alignment.vertical)
    applyOptionalStyleField(alignment, 'wrap', patch.alignment.wrap)
    applyOptionalStyleField(alignment, 'indent', patch.alignment.indent)
    if (Object.keys(alignment).length > 0) {
      next.alignment = alignment
    } else {
      delete next.alignment
    }
  }
  if (patch.borders === null) {
    delete next.borders
  } else if (patch.borders) {
    const borders = { ...next.borders }
    applyBorderSideStylePatch(borders, 'top', patch.borders.top)
    applyBorderSideStylePatch(borders, 'right', patch.borders.right)
    applyBorderSideStylePatch(borders, 'bottom', patch.borders.bottom)
    applyBorderSideStylePatch(borders, 'left', patch.borders.left)
    if (Object.keys(borders).length > 0) {
      next.borders = borders
    } else {
      delete next.borders
    }
  }
  return next
}

function applyBorderSideStylePatch(
  borders: NonNullable<CellStyleRecord['borders']>,
  side: keyof NonNullable<CellStyleRecord['borders']>,
  patch: NonNullable<CellStylePatch['borders']>['top'] | null | undefined,
): void {
  if (patch === undefined) {
    return
  }
  if (patch === null) {
    delete borders[side]
    return
  }
  const nextSide = {
    ...(borders[side] ?? DEFAULT_BORDER_SIDE),
  }
  applyOptionalStyleField(nextSide, 'style', patch.style)
  applyOptionalStyleField(nextSide, 'weight', patch.weight)
  applyOptionalStyleField(nextSide, 'color', patch.color)
  if (nextSide.style && nextSide.weight && nextSide.color) {
    borders[side] = nextSide
  } else {
    delete borders[side]
  }
}

export function borderPresetOptimisticPatch(preset: BorderPreset): CellStylePatch {
  switch (preset) {
    case 'clear':
      return { borders: null }
    case 'all':
    case 'outer':
      return {
        borders: {
          top: DEFAULT_BORDER_SIDE,
          right: DEFAULT_BORDER_SIDE,
          bottom: DEFAULT_BORDER_SIDE,
          left: DEFAULT_BORDER_SIDE,
        },
      }
    case 'left':
      return { borders: { left: DEFAULT_BORDER_SIDE } }
    case 'top':
      return { borders: { top: DEFAULT_BORDER_SIDE } }
    case 'right':
      return { borders: { right: DEFAULT_BORDER_SIDE } }
    case 'bottom':
      return { borders: { bottom: DEFAULT_BORDER_SIDE } }
    default: {
      const exhaustive: never = preset
      return exhaustive
    }
  }
}

export function clearStyleFieldsOptimisticPatch(fields?: readonly CellStyleField[]): CellStylePatch | null {
  if (fields === undefined) {
    return {
      alignment: null,
      borders: null,
      fill: null,
      font: null,
    }
  }
  const fieldSet = new Set(fields)
  const borders: NonNullable<CellStylePatch['borders']> = {}
  if (fieldSet.has('borderTop')) {
    borders.top = null
  }
  if (fieldSet.has('borderRight')) {
    borders.right = null
  }
  if (fieldSet.has('borderBottom')) {
    borders.bottom = null
  }
  if (fieldSet.has('borderLeft')) {
    borders.left = null
  }
  return Object.keys(borders).length > 0 ? { borders } : null
}
