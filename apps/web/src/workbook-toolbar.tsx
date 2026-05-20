import { memo, useCallback, type ReactNode } from 'react'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Baseline,
  Bold,
  Italic,
  PaintBucket,
  Redo2,
  RemoveFormatting,
  Undo2,
  Underline,
  WrapText,
} from 'lucide-react'
import { Toolbar } from '@base-ui/react/toolbar'
import { useWorkbookGridFocusReturn } from '@bilig/grid'
import type { CellHorizontalAlignment } from '@bilig/protocol'
import { cn } from './cn.js'
import { GOOGLE_SHEETS_SWATCH_ROWS } from './workbook-colors.js'
import { ColorPaletteButton } from './workbook-color-picker.js'
import { getWorkbookShortcutLabel } from './shortcut-registry.js'
import {
  ToolbarOverflowCue,
  WorkbookToolbarBorderMenu,
  WorkbookToolbarSelect,
  WorkbookToolbarStructureMenu,
} from './workbook-toolbar-controls.js'
import { useToolbarScrollCue } from './workbook-toolbar-scroll.js'
import type { BorderPreset, StructureActionAvailability, StructureActionTemplate } from './workbook-toolbar-options.js'
import {
  FONT_SIZE_OPTIONS,
  NUMBER_FORMAT_OPTIONS,
  toolbarButtonClass,
  toolbarFormattingRegionClass,
  toolbarFormattingScrollClass,
  toolbarGroupClass,
  toolbarIconClass,
  toolbarRootClass,
  toolbarRowClass,
  toolbarSegmentedClass,
  toolbarSelectTriggerClass,
  toolbarSeparatorClass,
  toolbarTrailingRegionClass,
} from './workbook-toolbar-theme.js'

export type { BorderPreset } from './workbook-toolbar-options.js'

export interface WorkbookToolbarProps {
  writesAllowed: boolean
  canUndo: boolean
  canRedo: boolean
  currentNumberFormatKind: string
  selectedFontSize: string
  isBoldActive: boolean
  isBorderActive?: boolean
  isItalicActive: boolean
  isUnderlineActive: boolean
  currentFillColor: string
  currentTextColor: string
  recentFillColors: readonly string[]
  recentTextColors: readonly string[]
  horizontalAlignment: CellHorizontalAlignment | null
  isWrapActive: boolean
  onNumberFormatChange(this: void, value: string): void
  onFontSizeChange(this: void, value: string): void
  onToggleBold(this: void): void
  onToggleItalic(this: void): void
  onToggleUnderline(this: void): void
  onFillColorSelect(this: void, color: string, source: 'preset' | 'custom'): void
  onFillColorReset(this: void): void
  onTextColorSelect(this: void, color: string, source: 'preset' | 'custom'): void
  onTextColorReset(this: void): void
  onHorizontalAlignmentChange(this: void, alignment: 'left' | 'center' | 'right'): void
  onApplyBorderPreset(this: void, preset: BorderPreset): void
  canHideCurrentRow: boolean
  canHideCurrentColumn: boolean
  canMergeSelection: boolean
  canUnmergeSelection: boolean
  canUnhideCurrentRow: boolean
  canUnhideCurrentColumn: boolean
  onHideCurrentRow(this: void): void
  onHideCurrentColumn(this: void): void
  onMergeSelectedCells(this: void): void
  onUnmergeSelectedCells(this: void): void
  onUnhideCurrentRow(this: void): void
  onUnhideCurrentColumn(this: void): void
  onToggleWrap(this: void): void
  onClearStyle(this: void): void
  onUndo(this: void): void
  onRedo(this: void): void
  trailingContent?: ReactNode
}

export const WorkbookToolbar = memo(function WorkbookToolbar({
  writesAllowed,
  canUndo,
  canRedo,
  currentNumberFormatKind,
  selectedFontSize,
  isBoldActive,
  isBorderActive = false,
  isItalicActive,
  isUnderlineActive,
  currentFillColor,
  currentTextColor,
  recentFillColors,
  recentTextColors,
  horizontalAlignment,
  isWrapActive,
  onNumberFormatChange,
  onFontSizeChange,
  onToggleBold,
  onToggleItalic,
  onToggleUnderline,
  onFillColorSelect,
  onFillColorReset,
  onTextColorSelect,
  onTextColorReset,
  onHorizontalAlignmentChange,
  onApplyBorderPreset,
  canHideCurrentRow,
  canHideCurrentColumn,
  canMergeSelection,
  canUnmergeSelection,
  canUnhideCurrentRow,
  canUnhideCurrentColumn,
  onHideCurrentRow,
  onHideCurrentColumn,
  onMergeSelectedCells,
  onUnmergeSelectedCells,
  onUnhideCurrentRow,
  onUnhideCurrentColumn,
  onToggleWrap,
  onClearStyle,
  onUndo,
  onRedo,
  trailingContent,
}: WorkbookToolbarProps) {
  const { scrollContainerRef, showBackwardCue, showForwardCue } = useToolbarScrollCue()
  const requestGridFocus = useWorkbookGridFocusReturn()
  const returnFocusToGridAfterCommand = useCallback(() => {
    if (!requestGridFocus) {
      return
    }
    const run = () => requestGridFocus()
    if (typeof window === 'undefined') {
      run()
      return
    }
    window.setTimeout(run, 0)
  }, [requestGridFocus])
  const runToolbarButtonCommand = useCallback(
    (command: () => void) => {
      command()
      returnFocusToGridAfterCommand()
    },
    [returnFocusToGridAfterCommand],
  )
  const getToolbarScrollStep = () => {
    const scrollContainer = scrollContainerRef.current
    return scrollContainer ? Math.max(96, Math.floor(scrollContainer.clientWidth)) : 96
  }
  const getToolbarBackwardScrollStep = () => {
    const scrollContainer = scrollContainerRef.current
    if (!scrollContainer) {
      return 96
    }

    const comfortableBackStep = Math.max(getToolbarScrollStep(), Math.floor(scrollContainer.clientWidth * 1.25))
    return Math.min(scrollContainer.scrollLeft, comfortableBackStep)
  }
  const scrollFormattingToolbarBackward = () => {
    const scrollContainer = scrollContainerRef.current
    if (!scrollContainer) {
      return
    }

    scrollContainer.scrollBy({
      behavior: 'smooth',
      left: -getToolbarBackwardScrollStep(),
    })
  }
  const scrollFormattingToolbarForward = () => {
    const scrollContainer = scrollContainerRef.current
    if (!scrollContainer) {
      return
    }

    scrollContainer.scrollBy({
      behavior: 'smooth',
      left: getToolbarScrollStep(),
    })
  }
  const structureActionAvailability: StructureActionAvailability = {
    mergeSelectedCells: canMergeSelection,
    unmergeSelectedCells: canUnmergeSelection,
    hideCurrentRow: canHideCurrentRow,
    hideCurrentColumn: canHideCurrentColumn,
    unhideCurrentRow: canUnhideCurrentRow,
    unhideCurrentColumn: canUnhideCurrentColumn,
  }
  const runStructureAction = (template: StructureActionTemplate) => {
    switch (template) {
      case 'mergeSelectedCells':
        onMergeSelectedCells()
        break
      case 'unmergeSelectedCells':
        onUnmergeSelectedCells()
        break
      case 'hideCurrentRow':
        onHideCurrentRow()
        break
      case 'hideCurrentColumn':
        onHideCurrentColumn()
        break
      case 'unhideCurrentRow':
        onUnhideCurrentRow()
        break
      case 'unhideCurrentColumn':
        onUnhideCurrentColumn()
        break
    }
    returnFocusToGridAfterCommand()
  }

  return (
    <div className={toolbarRootClass()}>
      <Toolbar.Root aria-label="Formatting toolbar" className={toolbarRowClass()}>
        <div className={toolbarFormattingRegionClass()}>
          {showBackwardCue ? <ToolbarOverflowCue direction="backward" onClick={scrollFormattingToolbarBackward} /> : null}
          <div ref={scrollContainerRef} className={toolbarFormattingScrollClass()} data-testid="toolbar-formatting-scroll">
            <Toolbar.Group className={toolbarGroupClass()}>
              <div className={toolbarSegmentedClass()} role="group" aria-label="History">
                <Toolbar.Button
                  aria-label="Undo"
                  className={toolbarButtonClass({ embedded: true })}
                  disabled={!writesAllowed || !canUndo}
                  title={`Undo (${getWorkbookShortcutLabel('undo')})`}
                  onClick={() => runToolbarButtonCommand(onUndo)}
                >
                  <Undo2 className={toolbarIconClass()} />
                </Toolbar.Button>
                <Toolbar.Button
                  aria-label="Redo"
                  className={toolbarButtonClass({ embedded: true })}
                  disabled={!writesAllowed || !canRedo}
                  title={`Redo (${getWorkbookShortcutLabel('redo')})`}
                  onClick={() => runToolbarButtonCommand(onRedo)}
                >
                  <Redo2 className={toolbarIconClass()} />
                </Toolbar.Button>
              </div>
            </Toolbar.Group>

            <Toolbar.Separator className={toolbarSeparatorClass()} />

            <Toolbar.Group className={toolbarGroupClass()}>
              <WorkbookToolbarSelect
                ariaLabel="Number format"
                options={NUMBER_FORMAT_OPTIONS}
                triggerClassName={cn(toolbarSelectTriggerClass(), 'w-32 max-[420px]:w-28 max-[360px]:w-24')}
                value={currentNumberFormatKind}
                valueClassName="min-w-0 flex-1 truncate whitespace-nowrap text-left"
                onChange={(nextValue) => {
                  onNumberFormatChange(nextValue)
                  returnFocusToGridAfterCommand()
                }}
              />
            </Toolbar.Group>

            <Toolbar.Separator className={toolbarSeparatorClass()} />

            <Toolbar.Group className={toolbarGroupClass()}>
              <WorkbookToolbarSelect
                ariaLabel="Font size"
                options={FONT_SIZE_OPTIONS}
                triggerClassName={cn(toolbarSelectTriggerClass(), 'w-[5rem] max-[420px]:w-14')}
                value={selectedFontSize}
                valueClassName="flex-none w-[2ch] overflow-visible text-center font-medium tabular-nums"
                onChange={(nextValue) => {
                  onFontSizeChange(nextValue)
                  returnFocusToGridAfterCommand()
                }}
              />
              <div className={toolbarSegmentedClass()} role="group" aria-label="Font emphasis">
                <Toolbar.Button
                  aria-label="Bold"
                  aria-pressed={isBoldActive}
                  className={toolbarButtonClass({ active: isBoldActive, embedded: true })}
                  title={`Bold (${getWorkbookShortcutLabel('bold')})`}
                  onClick={() => runToolbarButtonCommand(onToggleBold)}
                >
                  <Bold className={toolbarIconClass()} />
                </Toolbar.Button>
                <Toolbar.Button
                  aria-label="Italic"
                  aria-pressed={isItalicActive}
                  className={toolbarButtonClass({ active: isItalicActive, embedded: true })}
                  title={`Italic (${getWorkbookShortcutLabel('italic')})`}
                  onClick={() => runToolbarButtonCommand(onToggleItalic)}
                >
                  <Italic className={toolbarIconClass()} />
                </Toolbar.Button>
                <Toolbar.Button
                  aria-label="Underline"
                  aria-pressed={isUnderlineActive}
                  className={toolbarButtonClass({ active: isUnderlineActive, embedded: true })}
                  title={`Underline (${getWorkbookShortcutLabel('underline')})`}
                  onClick={() => runToolbarButtonCommand(onToggleUnderline)}
                >
                  <Underline className={toolbarIconClass()} />
                </Toolbar.Button>
              </div>
              <ColorPaletteButton
                ariaLabel="Fill color"
                currentColor={currentFillColor}
                customInputLabel="Custom fill color"
                icon={<PaintBucket className={toolbarIconClass()} />}
                onActionComplete={returnFocusToGridAfterCommand}
                onReset={onFillColorReset}
                onSelectColor={onFillColorSelect}
                recentColors={recentFillColors}
                swatches={GOOGLE_SHEETS_SWATCH_ROWS}
              />
              <ColorPaletteButton
                ariaLabel="Text color"
                currentColor={currentTextColor}
                customInputLabel="Custom text color"
                icon={<Baseline className={toolbarIconClass()} />}
                onActionComplete={returnFocusToGridAfterCommand}
                onReset={onTextColorReset}
                onSelectColor={onTextColorSelect}
                recentColors={recentTextColors}
                swatches={GOOGLE_SHEETS_SWATCH_ROWS}
              />
            </Toolbar.Group>

            <Toolbar.Separator className={toolbarSeparatorClass()} />

            <Toolbar.Group className={toolbarGroupClass()}>
              <div className={toolbarSegmentedClass()} role="group" aria-label="Horizontal alignment">
                <Toolbar.Button
                  aria-label="Align left"
                  aria-pressed={horizontalAlignment === 'left'}
                  className={toolbarButtonClass({
                    active: horizontalAlignment === 'left',
                    embedded: true,
                  })}
                  title={`Align left (${getWorkbookShortcutLabel('align-left')})`}
                  onClick={() => runToolbarButtonCommand(() => onHorizontalAlignmentChange('left'))}
                >
                  <AlignLeft className={toolbarIconClass()} />
                </Toolbar.Button>
                <Toolbar.Button
                  aria-label="Align center"
                  aria-pressed={horizontalAlignment === 'center'}
                  className={toolbarButtonClass({
                    active: horizontalAlignment === 'center',
                    embedded: true,
                  })}
                  title={`Align center (${getWorkbookShortcutLabel('align-center')})`}
                  onClick={() => runToolbarButtonCommand(() => onHorizontalAlignmentChange('center'))}
                >
                  <AlignCenter className={toolbarIconClass()} />
                </Toolbar.Button>
                <Toolbar.Button
                  aria-label="Align right"
                  aria-pressed={horizontalAlignment === 'right'}
                  className={toolbarButtonClass({
                    active: horizontalAlignment === 'right',
                    embedded: true,
                  })}
                  title={`Align right (${getWorkbookShortcutLabel('align-right')})`}
                  onClick={() => runToolbarButtonCommand(() => onHorizontalAlignmentChange('right'))}
                >
                  <AlignRight className={toolbarIconClass()} />
                </Toolbar.Button>
              </div>
            </Toolbar.Group>

            <Toolbar.Separator className={toolbarSeparatorClass()} />

            <Toolbar.Group className={toolbarGroupClass()}>
              <WorkbookToolbarBorderMenu
                disabled={!writesAllowed}
                isActive={isBorderActive}
                onApplyPreset={(preset) => runToolbarButtonCommand(() => onApplyBorderPreset(preset))}
              />
            </Toolbar.Group>

            <Toolbar.Separator className={toolbarSeparatorClass()} />

            <Toolbar.Group className={toolbarGroupClass()}>
              <WorkbookToolbarStructureMenu
                availability={structureActionAvailability}
                disabled={!writesAllowed}
                onRunAction={runStructureAction}
              />
            </Toolbar.Group>

            <Toolbar.Separator className={toolbarSeparatorClass()} />

            <Toolbar.Group className={toolbarGroupClass()}>
              <Toolbar.Button
                aria-label="Wrap"
                aria-pressed={isWrapActive}
                className={toolbarButtonClass({ active: isWrapActive })}
                type="button"
                onClick={() => runToolbarButtonCommand(onToggleWrap)}
              >
                <WrapText className={toolbarIconClass()} />
                <span className="sr-only">Wrap</span>
              </Toolbar.Button>
              <Toolbar.Button
                aria-label="Clear style"
                className={toolbarButtonClass()}
                type="button"
                onClick={() => runToolbarButtonCommand(onClearStyle)}
              >
                <RemoveFormatting className={toolbarIconClass()} />
                <span className="sr-only">Clear style</span>
              </Toolbar.Button>
            </Toolbar.Group>
          </div>
          {showForwardCue ? <ToolbarOverflowCue direction="forward" onClick={scrollFormattingToolbarForward} /> : null}
        </div>
        {trailingContent ? (
          <>
            <Toolbar.Separator className={cn(toolbarSeparatorClass(), 'max-[420px]:mx-0.5')} />
            <div className={toolbarTrailingRegionClass()} data-testid="toolbar-trailing-content">
              {trailingContent}
            </div>
          </>
        ) : null}
      </Toolbar.Root>
    </div>
  )
})
