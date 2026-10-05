import { useCallback, useEffect, useMemo, useRef, useState, type MutableRefObject, type ReactNode } from 'react'
import { WORKBOOK_DEFAULT_FONT_SIZE } from '@bilig/grid'
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
  applyToolbarStylePatch,
  BORDER_CLEAR_FIELDS,
  borderPresetOptimisticPatch,
  cellRangeKey,
  clearStyleFieldsOptimisticPatch,
  DEFAULT_BORDER_SIDE,
  deriveWorkbookStatusPresentation,
  hasAnyBorder,
  mergeToolbarStylePatch,
  type OptimisticToolbarStyle,
  queueWorkbookHistoryShortcut,
  selectedStyleMatchesPatch,
  shouldKeepWorkbookShortcutInsideTextEntry,
  shouldRouteWorkbookShortcutToWorkbookScope,
} from './workbook-toolbar-state.js'
import { createRangeRef, getNormalizedRangeBounds, type ZeroConnectionState } from './worker-workbook-app-model.js'

export { deriveWorkbookStatusPresentation } from './workbook-toolbar-state.js'

export function useWorkbookToolbar(input: {
  connectionStateName: ZeroConnectionState['name']
  runtimeReady: boolean
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
  hasLocalMutationInFlight?: boolean
  canUndo: boolean
  canRedo: boolean
  onUndo: () => void
  onRedo: () => void
  canHideCurrentRow: boolean
  canHideCurrentColumn: boolean
  canCreateTableFromSelection?: boolean
  canUnhideCurrentRow: boolean
  canUnhideCurrentColumn: boolean
  canUnmergeSelection: boolean
  onHideCurrentRow: () => void
  onHideCurrentColumn: () => void
  onCreateTableFromSelection?: () => void
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
    pendingMutationSummary,
    failedPendingMutation,
    remoteSyncAvailable,
    zeroConfigured,
    zeroHealthReady,
    hasLocalMutationInFlight,
    canUndo,
    canRedo,
    onUndo,
    onRedo,
    canHideCurrentRow,
    canHideCurrentColumn,
    canCreateTableFromSelection,
    canUnhideCurrentRow,
    canUnhideCurrentColumn,
    canUnmergeSelection,
    onHideCurrentRow,
    onHideCurrentColumn,
    onCreateTableFromSelection,
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
  const [optimisticStyle, setOptimisticStyle] = useState<OptimisticToolbarStyle | null>(null)
  const activeSelectedStyle = optimisticStyle?.rangeKey === selectedRangeKey ? optimisticStyle.style : selectedStyle
  const currentNumberFormat = parseCellNumberFormatCode(selectedCell.format)
  const selectedFontSize = String(activeSelectedStyle?.font?.size ?? WORKBOOK_DEFAULT_FONT_SIZE)
  const isBoldActive = activeSelectedStyle?.font?.bold === true
  const isItalicActive = activeSelectedStyle?.font?.italic === true
  const isUnderlineActive = activeSelectedStyle?.font?.underline === true
  const horizontalAlignment = activeSelectedStyle?.alignment?.horizontal ?? null
  const isWrapActive = activeSelectedStyle?.alignment?.wrap === true
  const isBorderActive = hasAnyBorder(activeSelectedStyle)
  const selectedRangeBounds = getNormalizedRangeBounds(selectionRangeRef.current)
  const toolbarMutationQueueRef = useRef<Promise<void> | null>(null)
  const canMergeSelection =
    selectedRangeBounds.startRow !== selectedRangeBounds.endRow || selectedRangeBounds.startCol !== selectedRangeBounds.endCol
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
  const applyOptimisticStylePatch = useCallback(
    (range: CellRangeRef, patch: CellStylePatch) => {
      const rangeKey = cellRangeKey(range)
      setOptimisticStyle((current) => ({
        rangeKey,
        patch: current?.rangeKey === rangeKey ? mergeToolbarStylePatch(current.patch, patch) : patch,
        style: applyToolbarStylePatch(current?.rangeKey === rangeKey ? current.style : selectedStyle, patch),
      }))
      return rangeKey
    },
    [selectedStyle],
  )
  const enqueueToolbarMutation = useCallback((run: () => Promise<void>) => {
    const queuedMutation = (async () => {
      const previousMutation = toolbarMutationQueueRef.current
      if (previousMutation) {
        await previousMutation
      }
      await run()
    })()
    let queueTail: Promise<void> | null = null
    queueTail = (async () => {
      try {
        await queuedMutation
      } catch {
        // The command caller receives the failure; later toolbar commands must keep draining in order.
      }
      if (toolbarMutationQueueRef.current === queueTail) {
        toolbarMutationQueueRef.current = null
      }
    })()
    toolbarMutationQueueRef.current = queueTail
    return queuedMutation
  }, [])
  const statusPresentation = deriveWorkbookStatusPresentation({
    connectionStateName,
    runtimeReady,
    pendingMutationSummary,
    failedPendingMutation,
    remoteSyncAvailable,
    zeroConfigured,
    zeroHealthReady,
    hasLocalMutationInFlight: hasLocalMutationInFlight === true,
    writesAllowed,
  })
  const statusModeLabel = statusPresentation.modeLabel

  useEffect(() => {
    if (optimisticStyle && optimisticStyle.rangeKey !== selectedRangeKey) {
      setOptimisticStyle(null)
    }
  }, [optimisticStyle, selectedRangeKey])

  useEffect(() => {
    if (optimisticStyle && selectedStyleMatchesPatch(selectedStyle, optimisticStyle.patch)) {
      setOptimisticStyle(null)
    }
  }, [optimisticStyle, selectedStyle])

  const handleUndo = useCallback(() => {
    setOptimisticStyle(null)
    onUndo()
  }, [onUndo])

  const handleRedo = useCallback(() => {
    setOptimisticStyle(null)
    onRedo()
  }, [onRedo])

  const applyRangeStyle = useCallback(
    async (patch: CellStylePatch) => {
      const range = selectionRangeRef.current
      const rangeKey = applyOptimisticStylePatch(range, patch)
      try {
        await enqueueToolbarMutation(() => invokeMutation('setRangeStyle', range, patch))
      } catch (error) {
        setOptimisticStyle((current) => (current?.rangeKey === rangeKey ? null : current))
        throw error
      }
    },
    [applyOptimisticStylePatch, enqueueToolbarMutation, invokeMutation, selectionRangeRef],
  )

  const clearRangeStyleFields = useCallback(
    async (fields?: CellStyleField[]) => {
      const range = selectionRangeRef.current
      const optimisticPatch = clearStyleFieldsOptimisticPatch(fields)
      const rangeKey = optimisticPatch ? applyOptimisticStylePatch(range, optimisticPatch) : null
      try {
        await enqueueToolbarMutation(async () => {
          await invokeMutation('clearRangeStyle', range, fields)
          if (fields !== undefined) {
            return
          }
          await invokeMutation('clearRangeNumberFormat', range)
        })
      } catch (error) {
        setOptimisticStyle((current) => (current?.rangeKey === rangeKey ? null : current))
        throw error
      }
    },
    [applyOptimisticStylePatch, enqueueToolbarMutation, invokeMutation, selectionRangeRef],
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
      const rangeKey = applyOptimisticStylePatch(selectionRange, borderPresetOptimisticPatch(preset))
      const { sheetName, startRow, endRow, startCol, endCol } = getNormalizedRangeBounds(selectionRange)
      const borderMutations: Array<readonly [CellRangeRef, CellStylePatch]> = []
      const queueBorders = (range: CellRangeRef, borders: NonNullable<CellStylePatch['borders']>) => {
        borderMutations.push([range, { borders }])
      }
      const queueRowBorder = (rowStart: number, rowEnd: number, side: 'top' | 'bottom') => {
        if (rowStart > rowEnd) {
          return
        }
        queueBorders(createRangeRef(sheetName, rowStart, startCol, rowEnd, endCol), {
          [side]: DEFAULT_BORDER_SIDE,
        })
      }
      const queueColumnBorder = (colStart: number, colEnd: number, side: 'left' | 'right') => {
        if (colStart > colEnd) {
          return
        }
        queueBorders(createRangeRef(sheetName, startRow, colStart, endRow, colEnd), {
          [side]: DEFAULT_BORDER_SIDE,
        })
      }

      switch (preset) {
        case 'clear':
          break
        case 'all':
          queueRowBorder(startRow, endRow, 'top')
          queueColumnBorder(startCol, endCol, 'left')
          queueRowBorder(endRow, endRow, 'bottom')
          queueColumnBorder(endCol, endCol, 'right')
          break
        case 'outer':
          queueRowBorder(startRow, startRow, 'top')
          queueRowBorder(endRow, endRow, 'bottom')
          queueColumnBorder(startCol, startCol, 'left')
          queueColumnBorder(endCol, endCol, 'right')
          break
        case 'left':
          queueColumnBorder(startCol, startCol, 'left')
          break
        case 'top':
          queueRowBorder(startRow, startRow, 'top')
          break
        case 'right':
          queueColumnBorder(endCol, endCol, 'right')
          break
        case 'bottom':
          queueRowBorder(endRow, endRow, 'bottom')
          break
        default: {
          const exhaustive: never = preset
          return exhaustive
        }
      }

      try {
        await enqueueToolbarMutation(async () => {
          await invokeMutation('clearRangeStyle', selectionRange, [...BORDER_CLEAR_FIELDS])
          await Promise.all(borderMutations.map(([range, patch]) => invokeMutation('setRangeStyle', range, patch)))
        })
      } catch (error) {
        setOptimisticStyle((current) => (current?.rangeKey === rangeKey ? null : current))
        throw error
      }
    },
    [applyOptimisticStylePatch, enqueueToolbarMutation, invokeMutation, selectionRangeRef],
  )

  const setNumberFormatPreset = useCallback(
    async (preset: string) => {
      const selectionRange = selectionRangeRef.current
      switch (preset) {
        case 'general':
          await enqueueToolbarMutation(() => invokeMutation('clearRangeNumberFormat', selectionRange))
          return
        case 'number':
          await enqueueToolbarMutation(() =>
            invokeMutation('setRangeNumberFormat', selectionRange, {
              kind: 'number',
              decimals: 2,
              useGrouping: true,
            }),
          )
          return
        case 'currency':
          await enqueueToolbarMutation(() =>
            invokeMutation('setRangeNumberFormat', selectionRange, {
              kind: 'currency',
              currency: 'USD',
              decimals: 2,
              useGrouping: true,
              negativeStyle: 'minus',
              zeroStyle: 'zero',
            }),
          )
          return
        case 'accounting':
          await enqueueToolbarMutation(() =>
            invokeMutation('setRangeNumberFormat', selectionRange, {
              kind: 'accounting',
              currency: 'USD',
              decimals: 2,
              useGrouping: true,
              negativeStyle: 'parentheses',
              zeroStyle: 'dash',
            }),
          )
          return
        case 'percent':
          await enqueueToolbarMutation(() =>
            invokeMutation('setRangeNumberFormat', selectionRange, {
              kind: 'percent',
              decimals: 2,
            }),
          )
          return
        case 'date':
          await enqueueToolbarMutation(() =>
            invokeMutation('setRangeNumberFormat', selectionRange, {
              kind: 'date',
              dateStyle: 'short',
            }),
          )
          return
        case 'text':
          await enqueueToolbarMutation(() => invokeMutation('setRangeNumberFormat', selectionRange, 'text'))
          return
      }
    },
    [enqueueToolbarMutation, invokeMutation, selectionRangeRef],
  )

  const mergeSelectedCells = useCallback(async () => {
    if (!canMergeSelection) {
      return
    }
    await invokeMutation('mergeCells', selectionRangeRef.current)
  }, [canMergeSelection, invokeMutation, selectionRangeRef])

  const unmergeSelectedCells = useCallback(async () => {
    if (!canUnmergeSelection) {
      return
    }
    await invokeMutation('unmergeCells', selectionRangeRef.current)
  }, [canUnmergeSelection, invokeMutation, selectionRangeRef])

  const shortcutStateRef = useRef({
    applyBorderPreset,
    applyRangeStyle,
    clearRangeStyleFields,
    isBoldActive,
    isItalicActive,
    isUnderlineActive,
    onRedo: handleRedo,
    onUndo: handleUndo,
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
    onRedo: handleRedo,
    onUndo: handleUndo,
    setNumberFormatPreset,
    writesAllowed,
  }

  useEffect(() => {
    const handleWindowHistoryInput = (event: InputEvent) => {
      if (event.defaultPrevented) {
        return
      }
      if (shouldKeepWorkbookShortcutInsideTextEntry(event.target)) {
        return
      }
      if (!shouldRouteWorkbookShortcutToWorkbookScope()) {
        return
      }
      if (event.inputType !== 'historyUndo' && event.inputType !== 'historyRedo') {
        return
      }
      const shortcutState = shortcutStateRef.current
      if (!shortcutState.writesAllowed) {
        return
      }
      if (event.inputType === 'historyUndo') {
        event.preventDefault()
        queueWorkbookHistoryShortcut(shortcutState.onUndo)
        return
      }
      if (event.inputType === 'historyRedo') {
        event.preventDefault()
        queueWorkbookHistoryShortcut(shortcutState.onRedo)
      }
    }

    const handleWindowShortcut = (event: KeyboardEvent) => {
      if (
        event.defaultPrevented ||
        shouldKeepWorkbookShortcutInsideTextEntry(event.target) ||
        !shouldRouteWorkbookShortcutToWorkbookScope() ||
        event.altKey
      ) {
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
        queueWorkbookHistoryShortcut(shortcutState.onUndo)
        return
      }
      if ((event.shiftKey && normalizedKey === 'z') || (!event.metaKey && event.ctrlKey && !event.shiftKey && normalizedKey === 'y')) {
        event.preventDefault()
        queueWorkbookHistoryShortcut(shortcutState.onRedo)
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

    window.addEventListener('beforeinput', handleWindowHistoryInput, true)
    window.addEventListener('keydown', handleWindowShortcut, true)
    document.addEventListener('keydown', handleWindowShortcut)
    return () => {
      window.removeEventListener('beforeinput', handleWindowHistoryInput, true)
      window.removeEventListener('keydown', handleWindowShortcut, true)
      document.removeEventListener('keydown', handleWindowShortcut)
    }
  }, [])

  const ribbon = useMemo(
    () => (
      <WorkbookToolbar
        canRedo={canRedo}
        canCreateTableFromSelection={canCreateTableFromSelection ?? canMergeSelection}
        canHideCurrentColumn={canHideCurrentColumn}
        canHideCurrentRow={canHideCurrentRow}
        canMergeSelection={canMergeSelection}
        canUnmergeSelection={canUnmergeSelection}
        canUnhideCurrentColumn={canUnhideCurrentColumn}
        canUnhideCurrentRow={canUnhideCurrentRow}
        canUndo={canUndo}
        currentFillColor={currentFillColor}
        currentNumberFormatKind={currentNumberFormat.kind}
        currentTextColor={currentTextColor}
        horizontalAlignment={horizontalAlignment}
        isBoldActive={isBoldActive}
        isBorderActive={isBorderActive}
        isItalicActive={isItalicActive}
        isUnderlineActive={isUnderlineActive}
        isWrapActive={isWrapActive}
        onApplyBorderPreset={applyBorderPreset}
        onClearStyle={() => {
          void clearRangeStyleFields()
        }}
        onRedo={handleRedo}
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
        {...(onCreateTableFromSelection ? { onCreateTableFromSelection } : {})}
        onMergeSelectedCells={() => {
          void mergeSelectedCells()
        }}
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
        onUndo={handleUndo}
        onUnhideCurrentColumn={onUnhideCurrentColumn}
        onUnhideCurrentRow={onUnhideCurrentRow}
        onUnmergeSelectedCells={() => {
          void unmergeSelectedCells()
        }}
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
      canCreateTableFromSelection,
      canHideCurrentColumn,
      canHideCurrentRow,
      canMergeSelection,
      canUnmergeSelection,
      canUnhideCurrentColumn,
      canUnhideCurrentRow,
      canUndo,
      currentFillColor,
      currentNumberFormat.kind,
      currentTextColor,
      horizontalAlignment,
      isBoldActive,
      isBorderActive,
      isItalicActive,
      isUnderlineActive,
      isWrapActive,
      mergeSelectedCells,
      handleRedo,
      onHideCurrentColumn,
      onHideCurrentRow,
      onCreateTableFromSelection,
      handleUndo,
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
      unmergeSelectedCells,
    ],
  )

  return {
    ribbon,
    statusModeLabel,
  }
}
