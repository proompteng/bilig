import { useCallback, useEffect, useMemo, useRef, useState, type MutableRefObject, type ReactNode } from 'react'
import {
  parseCellNumberFormatCode,
  type CellRangeRef,
  type CellSnapshot,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
} from '@bilig/protocol'
import { WorkbookToolbar, type BorderPreset } from './workbook-toolbar.js'
import { isPresetColor, mergeRecentCustomColors, normalizeHexColor } from './workbook-colors.js'
import { WorkbookHeaderStatusChip } from './workbook-header-controls.js'
import type { WorkbookMutationMethod } from './workbook-sync.js'
import {
  createRangeRef,
  formatConnectionStateLabel,
  getNormalizedRangeBounds,
  isTextEntryTarget,
  type ZeroConnectionState,
} from './worker-workbook-app-model.js'

const BORDER_CLEAR_FIELDS: readonly CellStyleField[] = ['borderTop', 'borderRight', 'borderBottom', 'borderLeft'] as const

const DEFAULT_BORDER_SIDE = {
  style: 'solid',
  weight: 'thin',
  color: '#111827',
} as const

const PENDING_STYLE_ID = '__bilig_pending_toolbar_style__'

type WorkbookHeaderStatusTone = 'positive' | 'progress' | 'warning' | 'danger' | 'neutral'

export interface WorkbookStatusPresentation {
  readonly modeLabel: string
  readonly syncLabel: string
  readonly tone: WorkbookHeaderStatusTone
}

export function deriveWorkbookStatusPresentation(input: {
  connectionStateName: ZeroConnectionState['name']
  runtimeReady: boolean
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower' | undefined
  remoteSyncAvailable: boolean
  zeroConfigured: boolean
  zeroHealthReady: boolean
  writesAllowed: boolean
  pendingMutationSummary?:
    | {
        readonly activeCount: number
        readonly failedCount: number
      }
    | undefined
  failedPendingMutation?: unknown
}): WorkbookStatusPresentation {
  const modeLabel = formatConnectionStateLabel(input.connectionStateName)
  if (!input.runtimeReady) {
    return { modeLabel, syncLabel: 'Loading…', tone: 'neutral' }
  }
  if (!input.writesAllowed) {
    return { modeLabel, syncLabel: 'Read only', tone: 'warning' }
  }
  if (input.failedPendingMutation || (input.pendingMutationSummary?.failedCount ?? 0) > 0) {
    return { modeLabel, syncLabel: 'Sync issue', tone: 'danger' }
  }
  if ((input.pendingMutationSummary?.activeCount ?? 0) > 0) {
    if (input.connectionStateName === 'connected' && input.remoteSyncAvailable && input.zeroHealthReady) {
      return { modeLabel, syncLabel: 'Saving…', tone: 'progress' }
    }
    return { modeLabel, syncLabel: 'Sync pending', tone: 'warning' }
  }
  if (input.localPersistenceMode === 'follower' && !input.remoteSyncAvailable) {
    return { modeLabel, syncLabel: 'Read only', tone: 'warning' }
  }
  if (!input.zeroConfigured) {
    return { modeLabel, syncLabel: 'Local only', tone: 'warning' }
  }
  if (input.connectionStateName === 'needs-auth' || input.connectionStateName === 'error') {
    return { modeLabel, syncLabel: 'Sync issue', tone: 'danger' }
  }
  if (input.connectionStateName === 'disconnected' || input.connectionStateName === 'closed') {
    return { modeLabel, syncLabel: 'Offline', tone: 'warning' }
  }
  if (input.connectionStateName === 'connecting' || !input.remoteSyncAvailable || !input.zeroHealthReady) {
    return { modeLabel, syncLabel: 'Local saved', tone: 'warning' }
  }
  return { modeLabel, syncLabel: 'Saved', tone: 'positive' }
}

function cellRangeKey(range: CellRangeRef): string {
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

function applyToolbarStylePatch(style: CellStyleRecord | undefined, patch: CellStylePatch): CellStyleRecord {
  const next = cloneStyleForToolbar(style)
  const backgroundColor = patch.fill?.backgroundColor
  if (backgroundColor !== undefined) {
    if (backgroundColor === null) {
      delete next.fill
    } else {
      next.fill = { backgroundColor }
    }
  }
  if (patch.font) {
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
  if (patch.alignment) {
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
  return next
}

export function useWorkbookToolbar(input: {
  connectionStateName: ZeroConnectionState['name']
  runtimeReady: boolean
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower' | undefined
  pendingMutationSummary?:
    | {
        readonly activeCount: number
        readonly failedCount: number
      }
    | undefined
  failedPendingMutation?: unknown
  remoteSyncAvailable: boolean
  zeroConfigured: boolean
  zeroHealthReady: boolean
  canUndo: boolean
  canRedo: boolean
  onUndo: () => void
  onRedo: () => void
  canHideCurrentRow: boolean
  canHideCurrentColumn: boolean
  canUnhideCurrentRow: boolean
  canUnhideCurrentColumn: boolean
  onHideCurrentRow: () => void
  onHideCurrentColumn: () => void
  onUnhideCurrentRow: () => void
  onUnhideCurrentColumn: () => void
  invokeMutation: (method: WorkbookMutationMethod, ...args: unknown[]) => Promise<void>
  selectionRangeRef: MutableRefObject<CellRangeRef>
  selectedCell: CellSnapshot
  selectedStyle: CellStyleRecord | undefined
  writesAllowed: boolean
  trailingContent?: ReactNode
}) {
  const {
    connectionStateName,
    runtimeReady,
    localPersistenceMode,
    pendingMutationSummary,
    failedPendingMutation,
    remoteSyncAvailable,
    zeroConfigured,
    zeroHealthReady,
    canUndo,
    canRedo,
    onUndo,
    onRedo,
    canHideCurrentRow,
    canHideCurrentColumn,
    canUnhideCurrentRow,
    canUnhideCurrentColumn,
    onHideCurrentRow,
    onHideCurrentColumn,
    onUnhideCurrentRow,
    onUnhideCurrentColumn,
    invokeMutation,
    selectionRangeRef,
    selectedCell,
    selectedStyle,
    writesAllowed,
    trailingContent,
  } = input
  const [recentFillColors, setRecentFillColors] = useState<readonly string[]>([])
  const [recentTextColors, setRecentTextColors] = useState<readonly string[]>([])
  const selectedRangeKey = cellRangeKey(selectionRangeRef.current)
  const [optimisticStyle, setOptimisticStyle] = useState<{
    readonly rangeKey: string
    readonly style: CellStyleRecord
  } | null>(null)
  useEffect(() => {
    setOptimisticStyle(null)
  }, [selectedStyle])
  const activeSelectedStyle = optimisticStyle?.rangeKey === selectedRangeKey ? optimisticStyle.style : selectedStyle
  const currentNumberFormat = parseCellNumberFormatCode(selectedCell.format)
  const selectedFontSize = String(activeSelectedStyle?.font?.size ?? 11)
  const isBoldActive = activeSelectedStyle?.font?.bold === true
  const isItalicActive = activeSelectedStyle?.font?.italic === true
  const isUnderlineActive = activeSelectedStyle?.font?.underline === true
  const horizontalAlignment = activeSelectedStyle?.alignment?.horizontal ?? null
  const isWrapActive = activeSelectedStyle?.alignment?.wrap === true
  const currentFillColor = normalizeHexColor(activeSelectedStyle?.fill?.backgroundColor ?? '#ffffff')
  const currentTextColor = normalizeHexColor(activeSelectedStyle?.font?.color ?? '#111827')
  const visibleRecentFillColors = useMemo(
    () => (isPresetColor(currentFillColor) ? recentFillColors : mergeRecentCustomColors(recentFillColors, currentFillColor)),
    [currentFillColor, recentFillColors],
  )
  const visibleRecentTextColors = useMemo(
    () => (isPresetColor(currentTextColor) ? recentTextColors : mergeRecentCustomColors(recentTextColors, currentTextColor)),
    [currentTextColor, recentTextColors],
  )
  const statusPresentation = deriveWorkbookStatusPresentation({
    connectionStateName,
    runtimeReady,
    localPersistenceMode,
    pendingMutationSummary,
    failedPendingMutation,
    remoteSyncAvailable,
    zeroConfigured,
    zeroHealthReady,
    writesAllowed,
  })
  const statusModeLabel = statusPresentation.modeLabel
  const applyRangeStyle = useCallback(
    async (patch: CellStylePatch) => {
      const range = selectionRangeRef.current
      const rangeKey = cellRangeKey(range)
      setOptimisticStyle((current) => ({
        rangeKey,
        style: applyToolbarStylePatch(current?.rangeKey === rangeKey ? current.style : selectedStyle, patch),
      }))
      try {
        await invokeMutation('setRangeStyle', range, patch)
      } catch (error) {
        setOptimisticStyle((current) => (current?.rangeKey === rangeKey ? null : current))
        throw error
      }
    },
    [invokeMutation, selectedStyle, selectionRangeRef],
  )

  const clearRangeStyleFields = useCallback(
    async (fields?: CellStyleField[]) => {
      await invokeMutation('clearRangeStyle', selectionRangeRef.current, fields)
    },
    [invokeMutation, selectionRangeRef],
  )

  const applyFillColor = useCallback(
    async (color: string, source: 'preset' | 'custom') => {
      const normalized = normalizeHexColor(color)
      await applyRangeStyle({ fill: { backgroundColor: normalized } })
      if (source === 'custom') {
        setRecentFillColors((current) => mergeRecentCustomColors(current, normalized))
      }
    },
    [applyRangeStyle],
  )

  const resetFillColor = useCallback(async () => {
    await applyRangeStyle({ fill: { backgroundColor: null } })
  }, [applyRangeStyle])

  const applyTextColor = useCallback(
    async (color: string, source: 'preset' | 'custom') => {
      const normalized = normalizeHexColor(color)
      await applyRangeStyle({ font: { color: normalized } })
      if (source === 'custom') {
        setRecentTextColors((current) => mergeRecentCustomColors(current, normalized))
      }
    },
    [applyRangeStyle],
  )

  const resetTextColor = useCallback(async () => {
    await applyRangeStyle({ font: { color: null } })
  }, [applyRangeStyle])

  const applyBorderPreset = useCallback(
    async (preset: BorderPreset) => {
      const selectionRange = selectionRangeRef.current
      const { sheetName, startRow, endRow, startCol, endCol } = getNormalizedRangeBounds(selectionRange)
      const applyBorders = async (range: CellRangeRef, borders: NonNullable<CellStylePatch['borders']>) => {
        await invokeMutation('setRangeStyle', range, { borders })
      }
      const applyRowBorder = async (rowStart: number, rowEnd: number, side: 'top' | 'bottom') => {
        if (rowStart > rowEnd) {
          return
        }
        await applyBorders(createRangeRef(sheetName, rowStart, startCol, rowEnd, endCol), {
          [side]: DEFAULT_BORDER_SIDE,
        })
      }
      const applyColumnBorder = async (colStart: number, colEnd: number, side: 'left' | 'right') => {
        if (colStart > colEnd) {
          return
        }
        await applyBorders(createRangeRef(sheetName, startRow, colStart, endRow, colEnd), {
          [side]: DEFAULT_BORDER_SIDE,
        })
      }

      await invokeMutation('clearRangeStyle', selectionRange, [...BORDER_CLEAR_FIELDS])

      switch (preset) {
        case 'clear':
          return
        case 'all':
          await applyRowBorder(startRow, endRow, 'top')
          await applyColumnBorder(startCol, endCol, 'left')
          await applyRowBorder(endRow, endRow, 'bottom')
          await applyColumnBorder(endCol, endCol, 'right')
          return
        case 'outer':
          await applyRowBorder(startRow, startRow, 'top')
          await applyRowBorder(endRow, endRow, 'bottom')
          await applyColumnBorder(startCol, startCol, 'left')
          await applyColumnBorder(endCol, endCol, 'right')
          return
        case 'left':
          await applyColumnBorder(startCol, startCol, 'left')
          return
        case 'top':
          await applyRowBorder(startRow, startRow, 'top')
          return
        case 'right':
          await applyColumnBorder(endCol, endCol, 'right')
          return
        case 'bottom':
          await applyRowBorder(endRow, endRow, 'bottom')
          return
        default: {
          const exhaustive: never = preset
          return exhaustive
        }
      }
    },
    [invokeMutation, selectionRangeRef],
  )

  const setNumberFormatPreset = useCallback(
    async (preset: string) => {
      const selectionRange = selectionRangeRef.current
      switch (preset) {
        case 'general':
          await invokeMutation('clearRangeNumberFormat', selectionRange)
          return
        case 'number':
          await invokeMutation('setRangeNumberFormat', selectionRange, {
            kind: 'number',
            decimals: 2,
            useGrouping: true,
          })
          return
        case 'currency':
          await invokeMutation('setRangeNumberFormat', selectionRange, {
            kind: 'currency',
            currency: 'USD',
            decimals: 2,
            useGrouping: true,
            negativeStyle: 'minus',
            zeroStyle: 'zero',
          })
          return
        case 'accounting':
          await invokeMutation('setRangeNumberFormat', selectionRange, {
            kind: 'accounting',
            currency: 'USD',
            decimals: 2,
            useGrouping: true,
            negativeStyle: 'parentheses',
            zeroStyle: 'dash',
          })
          return
        case 'percent':
          await invokeMutation('setRangeNumberFormat', selectionRange, {
            kind: 'percent',
            decimals: 2,
          })
          return
        case 'date':
          await invokeMutation('setRangeNumberFormat', selectionRange, {
            kind: 'date',
            dateStyle: 'short',
          })
          return
        case 'text':
          await invokeMutation('setRangeNumberFormat', selectionRange, 'text')
          return
      }
    },
    [invokeMutation, selectionRangeRef],
  )

  const shortcutStateRef = useRef({
    applyBorderPreset,
    applyRangeStyle,
    clearRangeStyleFields,
    isBoldActive,
    isItalicActive,
    isUnderlineActive,
    onRedo,
    onUndo,
    setNumberFormatPreset,
    writesAllowed,
  })
  shortcutStateRef.current = {
    applyBorderPreset,
    applyRangeStyle,
    clearRangeStyleFields,
    isBoldActive,
    isItalicActive,
    isUnderlineActive,
    onRedo,
    onUndo,
    setNumberFormatPreset,
    writesAllowed,
  }

  useEffect(() => {
    const handleWindowShortcut = (event: KeyboardEvent) => {
      if (event.defaultPrevented || isTextEntryTarget(event.target) || event.altKey) {
        return
      }

      const hasPrimaryModifier = event.metaKey || event.ctrlKey
      const shortcutState = shortcutStateRef.current
      if (!hasPrimaryModifier || !shortcutState.writesAllowed) {
        return
      }

      const normalizedKey = event.key.toLowerCase()
      if (!event.shiftKey && normalizedKey === 'z') {
        event.preventDefault()
        shortcutState.onUndo()
        return
      }
      if ((event.shiftKey && normalizedKey === 'z') || (!event.metaKey && event.ctrlKey && !event.shiftKey && normalizedKey === 'y')) {
        event.preventDefault()
        shortcutState.onRedo()
        return
      }
      if (!event.shiftKey && normalizedKey === 'b') {
        event.preventDefault()
        void shortcutState.applyRangeStyle({ font: { bold: !shortcutState.isBoldActive } })
        return
      }
      if (!event.shiftKey && normalizedKey === 'i') {
        event.preventDefault()
        void shortcutState.applyRangeStyle({ font: { italic: !shortcutState.isItalicActive } })
        return
      }
      if (!event.shiftKey && normalizedKey === 'u') {
        event.preventDefault()
        void shortcutState.applyRangeStyle({ font: { underline: !shortcutState.isUnderlineActive } })
        return
      }
      if (event.shiftKey && event.code === 'Digit1') {
        event.preventDefault()
        void shortcutState.setNumberFormatPreset('number')
        return
      }
      if (event.shiftKey && event.code === 'Digit4') {
        event.preventDefault()
        void shortcutState.setNumberFormatPreset('currency')
        return
      }
      if (event.shiftKey && event.code === 'Digit5') {
        event.preventDefault()
        void shortcutState.setNumberFormatPreset('percent')
        return
      }
      if (event.shiftKey && event.code === 'Digit7') {
        event.preventDefault()
        void shortcutState.applyBorderPreset('outer')
        return
      }
      if (event.shiftKey && normalizedKey === 'l') {
        event.preventDefault()
        void shortcutState.applyRangeStyle({ alignment: { horizontal: 'left' } })
        return
      }
      if (event.shiftKey && normalizedKey === 'e') {
        event.preventDefault()
        void shortcutState.applyRangeStyle({ alignment: { horizontal: 'center' } })
        return
      }
      if (event.shiftKey && normalizedKey === 'r') {
        event.preventDefault()
        void shortcutState.applyRangeStyle({ alignment: { horizontal: 'right' } })
        return
      }
      if (!event.shiftKey && event.code === 'Backslash') {
        event.preventDefault()
        void shortcutState.clearRangeStyleFields()
      }
    }

    window.addEventListener('keydown', handleWindowShortcut, true)
    return () => {
      window.removeEventListener('keydown', handleWindowShortcut, true)
    }
  }, [])

  const ribbon = useMemo(
    () => (
      <WorkbookToolbar
        canRedo={canRedo}
        canHideCurrentColumn={canHideCurrentColumn}
        canHideCurrentRow={canHideCurrentRow}
        canUnhideCurrentColumn={canUnhideCurrentColumn}
        canUnhideCurrentRow={canUnhideCurrentRow}
        canUndo={canUndo}
        currentFillColor={currentFillColor}
        currentNumberFormatKind={currentNumberFormat.kind}
        currentTextColor={currentTextColor}
        horizontalAlignment={horizontalAlignment}
        isBoldActive={isBoldActive}
        isItalicActive={isItalicActive}
        isUnderlineActive={isUnderlineActive}
        isWrapActive={isWrapActive}
        onApplyBorderPreset={applyBorderPreset}
        onClearStyle={() => {
          void clearRangeStyleFields()
        }}
        onRedo={onRedo}
        onFillColorReset={() => {
          void resetFillColor()
        }}
        onFillColorSelect={(color, source) => {
          void applyFillColor(color, source)
        }}
        onFontSizeChange={(value) => {
          void applyRangeStyle({ font: { size: value ? Number(value) : null } })
        }}
        onHorizontalAlignmentChange={(alignment) => {
          void applyRangeStyle({
            alignment: {
              horizontal: horizontalAlignment === alignment ? null : alignment,
            },
          })
        }}
        onHideCurrentColumn={onHideCurrentColumn}
        onHideCurrentRow={onHideCurrentRow}
        onNumberFormatChange={(value) => {
          void setNumberFormatPreset(value)
        }}
        onTextColorReset={() => {
          void resetTextColor()
        }}
        onTextColorSelect={(color, source) => {
          void applyTextColor(color, source)
        }}
        onToggleBold={() => {
          void applyRangeStyle({ font: { bold: !isBoldActive } })
        }}
        onToggleItalic={() => {
          void applyRangeStyle({ font: { italic: !isItalicActive } })
        }}
        onToggleUnderline={() => {
          void applyRangeStyle({ font: { underline: !isUnderlineActive } })
        }}
        onToggleWrap={() => {
          void applyRangeStyle({
            alignment: { wrap: !isWrapActive },
          })
        }}
        onUndo={onUndo}
        onUnhideCurrentColumn={onUnhideCurrentColumn}
        onUnhideCurrentRow={onUnhideCurrentRow}
        recentFillColors={visibleRecentFillColors}
        recentTextColors={visibleRecentTextColors}
        selectedFontSize={selectedFontSize}
        trailingContent={
          <>
            <WorkbookHeaderStatusChip
              modeLabel={statusPresentation.modeLabel}
              syncLabel={statusPresentation.syncLabel}
              tone={statusPresentation.tone}
            />
            {trailingContent}
          </>
        }
        writesAllowed={writesAllowed}
      />
    ),
    [
      applyBorderPreset,
      applyFillColor,
      applyRangeStyle,
      applyTextColor,
      clearRangeStyleFields,
      canRedo,
      canHideCurrentColumn,
      canHideCurrentRow,
      canUnhideCurrentColumn,
      canUnhideCurrentRow,
      canUndo,
      currentFillColor,
      currentNumberFormat.kind,
      currentTextColor,
      horizontalAlignment,
      isBoldActive,
      isItalicActive,
      isUnderlineActive,
      isWrapActive,
      onRedo,
      onHideCurrentColumn,
      onHideCurrentRow,
      onUndo,
      onUnhideCurrentColumn,
      onUnhideCurrentRow,
      resetFillColor,
      resetTextColor,
      selectedFontSize,
      setNumberFormatPreset,
      statusPresentation.modeLabel,
      statusPresentation.syncLabel,
      statusPresentation.tone,
      trailingContent,
      visibleRecentFillColors,
      visibleRecentTextColors,
      writesAllowed,
    ],
  )

  return {
    ribbon,
    statusModeLabel,
  }
}
