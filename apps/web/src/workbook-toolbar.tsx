import { memo, type ReactNode } from 'react'
import { Check, ChevronDown } from 'lucide-react'
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
import { Popover } from '@base-ui/react/popover'
import { Select } from '@base-ui/react/select'
import { Toolbar } from '@base-ui/react/toolbar'
import { HugeiconsIcon, type IconSvgElement } from '@hugeicons/react'
import {
  BorderAllIcon,
  BorderBottomIcon,
  BorderFullIcon,
  BorderLeftIcon,
  BorderNoneIcon,
  BorderRightIcon,
  BorderTopIcon,
} from '@hugeicons/core-free-icons'
import type { CellHorizontalAlignment } from '@bilig/protocol'
import { cn } from './cn.js'
import { GOOGLE_SHEETS_SWATCH_ROWS } from './workbook-colors.js'
import { ColorPaletteButton } from './workbook-color-picker.js'
import { getWorkbookShortcutLabel } from './shortcut-registry.js'
import {
  FONT_SIZE_OPTIONS,
  NUMBER_FORMAT_OPTIONS,
  toolbarBorderIconClass,
  toolbarBorderPopupClass,
  toolbarButtonClass,
  toolbarGroupClass,
  toolbarIconClass,
  toolbarPopupClass,
  toolbarRootClass,
  toolbarRowClass,
  toolbarSegmentedClass,
  toolbarSelectTriggerClass,
  toolbarSeparatorClass,
} from './workbook-toolbar-theme.js'

export type BorderPreset = 'all' | 'outer' | 'left' | 'top' | 'right' | 'bottom' | 'clear'

interface BorderPresetOption {
  key: BorderPreset
  icon: IconSvgElement
  label: string
  shortLabel: string
}

const BORDER_PRESET_OPTIONS: readonly BorderPresetOption[] = [
  { key: 'all', label: 'All borders', shortLabel: 'All', icon: BorderAllIcon },
  { key: 'outer', label: 'Outer borders', shortLabel: 'Outer', icon: BorderFullIcon },
  { key: 'left', label: 'Left border', shortLabel: 'Left', icon: BorderLeftIcon },
  { key: 'top', label: 'Top border', shortLabel: 'Top', icon: BorderTopIcon },
  { key: 'right', label: 'Right border', shortLabel: 'Right', icon: BorderRightIcon },
  { key: 'bottom', label: 'Bottom border', shortLabel: 'Bottom', icon: BorderBottomIcon },
  { key: 'clear', label: 'Clear borders', shortLabel: 'Clear', icon: BorderNoneIcon },
] as const

const STRUCTURE_ACTIONS = [
  { key: 'hide-row', label: 'Hide row', template: 'hideCurrentRow' },
  { key: 'unhide-row', label: 'Unhide row', template: 'unhideCurrentRow' },
  { key: 'hide-column', label: 'Hide column', template: 'hideCurrentColumn' },
  { key: 'unhide-column', label: 'Unhide column', template: 'unhideCurrentColumn' },
] as const

export interface WorkbookToolbarProps {
  writesAllowed: boolean
  canUndo: boolean
  canRedo: boolean
  currentNumberFormatKind: string
  selectedFontSize: string
  isBoldActive: boolean
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
  canUnhideCurrentRow: boolean
  canUnhideCurrentColumn: boolean
  onHideCurrentRow(this: void): void
  onHideCurrentColumn(this: void): void
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
  canUnhideCurrentRow,
  canUnhideCurrentColumn,
  onHideCurrentRow,
  onHideCurrentColumn,
  onUnhideCurrentRow,
  onUnhideCurrentColumn,
  onToggleWrap,
  onClearStyle,
  onUndo,
  onRedo,
  trailingContent,
}: WorkbookToolbarProps) {
  const structureActionAvailability = {
    hideCurrentRow: canHideCurrentRow,
    hideCurrentColumn: canHideCurrentColumn,
    unhideCurrentRow: canUnhideCurrentRow,
    unhideCurrentColumn: canUnhideCurrentColumn,
  } as const

  return (
    <div className={toolbarRootClass()}>
      <Toolbar.Root aria-label="Formatting toolbar" className={toolbarRowClass()}>
        <Toolbar.Group className={toolbarGroupClass()}>
          <div className={toolbarSegmentedClass()} role="group" aria-label="History">
            <Toolbar.Button
              aria-label="Undo"
              className={toolbarButtonClass({ embedded: true })}
              disabled={!writesAllowed || !canUndo}
              title={`Undo (${getWorkbookShortcutLabel('undo')})`}
              onClick={onUndo}
            >
              <Undo2 className={toolbarIconClass()} />
            </Toolbar.Button>
            <Toolbar.Button
              aria-label="Redo"
              className={toolbarButtonClass({ embedded: true })}
              disabled={!writesAllowed || !canRedo}
              title={`Redo (${getWorkbookShortcutLabel('redo')})`}
              onClick={onRedo}
            >
              <Redo2 className={toolbarIconClass()} />
            </Toolbar.Button>
          </div>
        </Toolbar.Group>

        <Toolbar.Separator className={toolbarSeparatorClass()} />

        <Toolbar.Group className={toolbarGroupClass()}>
          <Select.Root
            items={NUMBER_FORMAT_OPTIONS}
            value={currentNumberFormatKind}
            onValueChange={(nextValue: string | null) => {
              if (typeof nextValue === 'string') {
                onNumberFormatChange(nextValue)
              }
            }}
          >
            <Select.Trigger
              aria-label="Number format"
              className={cn(toolbarSelectTriggerClass(), 'w-32')}
              data-current-value={currentNumberFormatKind}
            >
              <span className="min-w-0 flex-1 truncate whitespace-nowrap text-left">
                {(NUMBER_FORMAT_OPTIONS.find((option) => option.value === currentNumberFormatKind) ?? NUMBER_FORMAT_OPTIONS[0])?.label ??
                  ''}
              </span>
              <Select.Icon className="ml-2 text-[var(--color-mauve-700)]">
                <ChevronDown className="h-3.5 w-3.5 stroke-[1.75]" />
              </Select.Icon>
            </Select.Trigger>
            <Select.Portal>
              <Select.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={6}>
                <Select.Popup className={toolbarPopupClass()}>
                  <Select.List className="max-h-72 min-w-[var(--anchor-width)] overflow-auto py-1">
                    {NUMBER_FORMAT_OPTIONS.map((option) => (
                      <Select.Item
                        className="flex cursor-default items-center justify-between gap-3 rounded-md px-2 py-1.5 text-[12px] text-[var(--color-mauve-900)] outline-none data-[highlighted]:bg-[var(--color-mauve-100)] data-[selected]:font-semibold"
                        key={`number-format-${option.value || 'default'}`}
                        label={option.label}
                        value={option.value}
                      >
                        <Select.ItemText>{option.label}</Select.ItemText>
                        <Select.ItemIndicator className="text-[var(--color-mauve-700)]">
                          <Check className="h-3.5 w-3.5 stroke-[2]" />
                        </Select.ItemIndicator>
                      </Select.Item>
                    ))}
                  </Select.List>
                </Select.Popup>
              </Select.Positioner>
            </Select.Portal>
          </Select.Root>
        </Toolbar.Group>

        <Toolbar.Separator className={toolbarSeparatorClass()} />

        <Toolbar.Group className={toolbarGroupClass()}>
          <Select.Root
            items={FONT_SIZE_OPTIONS}
            value={selectedFontSize}
            onValueChange={(nextValue: string | null) => {
              if (typeof nextValue === 'string') {
                onFontSizeChange(nextValue)
              }
            }}
          >
            <Select.Trigger
              aria-label="Font size"
              className={cn(toolbarSelectTriggerClass(), 'w-[5rem]')}
              data-current-value={selectedFontSize}
            >
              <span className="flex-none w-[2ch] overflow-visible text-center font-medium tabular-nums">
                {(FONT_SIZE_OPTIONS.find((option) => option.value === selectedFontSize) ?? FONT_SIZE_OPTIONS[0])?.label ?? ''}
              </span>
              <Select.Icon className="ml-2 text-[var(--color-mauve-700)]">
                <ChevronDown className="h-3.5 w-3.5 stroke-[1.75]" />
              </Select.Icon>
            </Select.Trigger>
            <Select.Portal>
              <Select.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={6}>
                <Select.Popup className={toolbarPopupClass()}>
                  <Select.List className="max-h-72 min-w-[var(--anchor-width)] overflow-auto py-1">
                    {FONT_SIZE_OPTIONS.map((option) => (
                      <Select.Item
                        className="flex cursor-default items-center justify-between gap-3 rounded-md px-2 py-1.5 text-[12px] text-[var(--color-mauve-900)] outline-none data-[highlighted]:bg-[var(--color-mauve-100)] data-[selected]:font-semibold"
                        key={`font-size-${option.value || 'default'}`}
                        label={option.label}
                        value={option.value}
                      >
                        <Select.ItemText>{option.label}</Select.ItemText>
                        <Select.ItemIndicator className="text-[var(--color-mauve-700)]">
                          <Check className="h-3.5 w-3.5 stroke-[2]" />
                        </Select.ItemIndicator>
                      </Select.Item>
                    ))}
                  </Select.List>
                </Select.Popup>
              </Select.Positioner>
            </Select.Portal>
          </Select.Root>
          <div className={toolbarSegmentedClass()} role="group" aria-label="Font emphasis">
            <Toolbar.Button
              aria-label="Bold"
              aria-pressed={isBoldActive}
              className={toolbarButtonClass({ active: isBoldActive, embedded: true })}
              title={`Bold (${getWorkbookShortcutLabel('bold')})`}
              onClick={onToggleBold}
            >
              <Bold className={toolbarIconClass()} />
            </Toolbar.Button>
            <Toolbar.Button
              aria-label="Italic"
              aria-pressed={isItalicActive}
              className={toolbarButtonClass({ active: isItalicActive, embedded: true })}
              title={`Italic (${getWorkbookShortcutLabel('italic')})`}
              onClick={onToggleItalic}
            >
              <Italic className={toolbarIconClass()} />
            </Toolbar.Button>
            <Toolbar.Button
              aria-label="Underline"
              aria-pressed={isUnderlineActive}
              className={toolbarButtonClass({ active: isUnderlineActive, embedded: true })}
              title={`Underline (${getWorkbookShortcutLabel('underline')})`}
              onClick={onToggleUnderline}
            >
              <Underline className={toolbarIconClass()} />
            </Toolbar.Button>
          </div>
          <ColorPaletteButton
            ariaLabel="Fill color"
            currentColor={currentFillColor}
            customInputLabel="Custom fill color"
            icon={<PaintBucket className={toolbarIconClass()} />}
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
              onClick={() => onHorizontalAlignmentChange('left')}
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
              onClick={() => onHorizontalAlignmentChange('center')}
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
              onClick={() => onHorizontalAlignmentChange('right')}
            >
              <AlignRight className={toolbarIconClass()} />
            </Toolbar.Button>
          </div>
        </Toolbar.Group>

        <Toolbar.Separator className={toolbarSeparatorClass()} />

        <Toolbar.Group className={toolbarGroupClass()}>
          <Popover.Root modal={false}>
            <Popover.Trigger
              aria-label="Borders"
              aria-haspopup="menu"
              className={cn(toolbarButtonClass(), 'gap-1 px-1.5')}
              disabled={!writesAllowed}
              title="Borders"
              type="button"
            >
              <HugeiconsIcon className={toolbarBorderIconClass()} color="currentColor" icon={BorderAllIcon} size={20} strokeWidth={1.6} />
              <ChevronDown className="h-3 w-3 shrink-0 stroke-[1.75] text-[var(--wb-text-muted)]" />
            </Popover.Trigger>
            <Popover.Portal>
              <Popover.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={8}>
                <Popover.Popup aria-label="Border presets" className={cn(toolbarBorderPopupClass(), 'w-[188px]')}>
                  <div className="grid grid-cols-2 gap-1">
                    {BORDER_PRESET_OPTIONS.map(({ key, label, shortLabel, icon }) => (
                      <Toolbar.Button
                        key={key}
                        aria-label={label}
                        className="inline-flex h-8 items-center gap-2 rounded-md border border-transparent px-2 text-left text-[11px] font-medium text-[var(--color-mauve-900)] outline-none transition-colors hover:bg-[var(--color-mauve-100)] focus-visible:border-[var(--color-mauve-400)] focus-visible:bg-[var(--color-mauve-100)]"
                        title={label}
                        type="button"
                        onClick={() => onApplyBorderPreset(key)}
                      >
                        <HugeiconsIcon
                          className="shrink-0 text-[var(--color-mauve-700)]"
                          color="currentColor"
                          icon={icon}
                          size={18}
                          strokeWidth={1.6}
                        />
                        <span className="truncate">{shortLabel}</span>
                      </Toolbar.Button>
                    ))}
                  </div>
                </Popover.Popup>
              </Popover.Positioner>
            </Popover.Portal>
          </Popover.Root>
        </Toolbar.Group>

        <Toolbar.Separator className={toolbarSeparatorClass()} />

        <Toolbar.Group className={toolbarGroupClass()}>
          <Popover.Root modal={false}>
            <Popover.Trigger
              aria-label="Structure"
              aria-haspopup="menu"
              className={cn(toolbarButtonClass(), 'gap-1 px-2')}
              disabled={!writesAllowed || STRUCTURE_ACTIONS.every((action) => !structureActionAvailability[action.template])}
              title="Structure"
              type="button"
            >
              <span className="text-[11px] font-semibold">Structure</span>
              <ChevronDown className="h-3 w-3 shrink-0 stroke-[1.75] text-[var(--wb-text-muted)]" />
            </Popover.Trigger>
            <Popover.Portal>
              <Popover.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={8}>
                <Popover.Popup aria-label="Structure actions" className={cn(toolbarPopupClass(), 'w-[176px] p-1')}>
                  <div className="grid gap-1">
                    {STRUCTURE_ACTIONS.map((action) => (
                      <Toolbar.Button
                        aria-label={action.label}
                        className="inline-flex h-8 items-center rounded-md border border-transparent px-2 text-left text-[11px] font-medium text-[var(--color-mauve-900)] outline-none transition-colors hover:bg-[var(--color-mauve-100)] focus-visible:border-[var(--color-mauve-400)] focus-visible:bg-[var(--color-mauve-100)] disabled:cursor-not-allowed disabled:opacity-45"
                        disabled={!structureActionAvailability[action.template]}
                        key={action.key}
                        type="button"
                        onClick={() => {
                          switch (action.template) {
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
                        }}
                      >
                        {action.label}
                      </Toolbar.Button>
                    ))}
                  </div>
                </Popover.Popup>
              </Popover.Positioner>
            </Popover.Portal>
          </Popover.Root>
        </Toolbar.Group>

        <Toolbar.Separator className={toolbarSeparatorClass()} />

        <Toolbar.Group className={toolbarGroupClass()}>
          <Toolbar.Button
            aria-label="Wrap"
            aria-pressed={isWrapActive}
            className={toolbarButtonClass({ active: isWrapActive })}
            type="button"
            onClick={onToggleWrap}
          >
            <WrapText className={toolbarIconClass()} />
            <span className="sr-only">Wrap</span>
          </Toolbar.Button>
          <Toolbar.Button aria-label="Clear style" className={toolbarButtonClass()} type="button" onClick={onClearStyle}>
            <RemoveFormatting className={toolbarIconClass()} />
            <span className="sr-only">Clear style</span>
          </Toolbar.Button>
        </Toolbar.Group>
        {trailingContent ? (
          <>
            <Toolbar.Separator className={toolbarSeparatorClass()} />
            <div className="ml-auto flex flex-none items-center gap-1.5 pl-2" data-testid="toolbar-trailing-content">
              {trailingContent}
            </div>
          </>
        ) : null}
      </Toolbar.Root>
    </div>
  )
})
