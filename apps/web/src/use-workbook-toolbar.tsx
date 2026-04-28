import { useCallback, useEffect, useMemo, useState, type MutableRefObject, type ReactNode } from 'react'
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
  const currentNumberFormat = parseCellNumberFormatCode(selectedCell.format)
  const selectedFontSize = String(selectedStyle?.font?.size ?? 11)
  const isBoldActive = selectedStyle?.font?.bold === true
  const isItalicActive = selectedStyle?.font?.italic === true
  const isUnderlineActive = selectedStyle?.font?.underline === true
  const horizontalAlignment = selectedStyle?.alignment?.horizontal ?? null
  const isWrapActive = selectedStyle?.alignment?.wrap === true
  const currentFillColor = normalizeHexColor(selectedStyle?.fill?.backgroundColor ?? '#ffffff')
  const currentTextColor = normalizeHexColor(selectedStyle?.font?.color ?? '#111827')
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
      await invokeMutation('setRangeStyle', selectionRangeRef.current, patch)
    },
    [invokeMutation, selectionRangeRef],
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

  useEffect(() => {
    const handleWindowShortcut = (event: KeyboardEvent) => {
      if (event.defaultPrevented || isTextEntryTarget(event.target) || event.altKey) {
        return
      }

      const hasPrimaryModifier = event.metaKey || event.ctrlKey
      if (!hasPrimaryModifier || !writesAllowed) {
        return
      }

      const normalizedKey = event.key.toLowerCase()
      if (!event.shiftKey && normalizedKey === 'z') {
        event.preventDefault()
        onUndo()
        return
      }
      if ((event.shiftKey && normalizedKey === 'z') || (!event.metaKey && event.ctrlKey && !event.shiftKey && normalizedKey === 'y')) {
        event.preventDefault()
        onRedo()
        return
      }
      if (!event.shiftKey && normalizedKey === 'b') {
        event.preventDefault()
        void applyRangeStyle({ font: { bold: !isBoldActive } })
        return
      }
      if (!event.shiftKey && normalizedKey === 'i') {
        event.preventDefault()
        void applyRangeStyle({ font: { italic: !isItalicActive } })
        return
      }
      if (!event.shiftKey && normalizedKey === 'u') {
        event.preventDefault()
        void applyRangeStyle({ font: { underline: !isUnderlineActive } })
        return
      }
      if (event.shiftKey && event.code === 'Digit1') {
        event.preventDefault()
        void setNumberFormatPreset('number')
        return
      }
      if (event.shiftKey && event.code === 'Digit4') {
        event.preventDefault()
        void setNumberFormatPreset('currency')
        return
      }
      if (event.shiftKey && event.code === 'Digit5') {
        event.preventDefault()
        void setNumberFormatPreset('percent')
        return
      }
      if (event.shiftKey && event.code === 'Digit7') {
        event.preventDefault()
        void applyBorderPreset('outer')
        return
      }
      if (event.shiftKey && normalizedKey === 'l') {
        event.preventDefault()
        void applyRangeStyle({ alignment: { horizontal: 'left' } })
        return
      }
      if (event.shiftKey && normalizedKey === 'e') {
        event.preventDefault()
        void applyRangeStyle({ alignment: { horizontal: 'center' } })
        return
      }
      if (event.shiftKey && normalizedKey === 'r') {
        event.preventDefault()
        void applyRangeStyle({ alignment: { horizontal: 'right' } })
        return
      }
      if (!event.shiftKey && event.code === 'Backslash') {
        event.preventDefault()
        void clearRangeStyleFields()
      }
    }

    window.addEventListener('keydown', handleWindowShortcut, true)
    return () => {
      window.removeEventListener('keydown', handleWindowShortcut, true)
    }
  }, [
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
  ])

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
