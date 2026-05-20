import { Check, ChevronDown, ChevronLeft, ChevronRight, TableProperties } from 'lucide-react'
import { Popover } from '@base-ui/react/popover'
import { Select } from '@base-ui/react/select'
import { Toolbar } from '@base-ui/react/toolbar'
import { HugeiconsIcon } from '@hugeicons/react'
import { BorderAllIcon } from '@hugeicons/core-free-icons'
import { cn } from './cn.js'
import {
  BORDER_PRESET_OPTIONS,
  STRUCTURE_ACTIONS,
  hasAvailableStructureAction,
  type BorderPreset,
  type StructureActionAvailability,
  type StructureActionTemplate,
} from './workbook-toolbar-options.js'
import {
  toolbarBorderIconClass,
  toolbarBorderPopupClass,
  toolbarButtonClass,
  toolbarIconClass,
  toolbarOverflowCueClass,
  toolbarPopupClass,
} from './workbook-toolbar-theme.js'
import type { ToolbarSelectOption } from './workbook-toolbar-theme.js'

export function ToolbarOverflowCue(props: { readonly direction: 'backward' | 'forward'; readonly onClick: () => void }) {
  const isBackward = props.direction === 'backward'
  const Icon = isBackward ? ChevronLeft : ChevronRight

  return (
    <button
      aria-label={isBackward ? 'Show previous toolbar actions' : 'Show more toolbar actions'}
      className={cn(toolbarOverflowCueClass(), isBackward ? 'border-l-0 border-r' : null)}
      data-testid={isBackward ? 'toolbar-overflow-back-cue' : 'toolbar-overflow-cue'}
      title={isBackward ? 'Show previous toolbar actions' : 'Show more toolbar actions'}
      type="button"
      onClick={props.onClick}
    >
      <Icon className="h-3.5 w-3.5 stroke-[2]" />
    </button>
  )
}

export function WorkbookToolbarSelect(props: {
  readonly ariaLabel: string
  readonly options: readonly ToolbarSelectOption[]
  readonly triggerClassName: string
  readonly value: string
  readonly valueClassName: string
  readonly onChange: (value: string) => void
}) {
  const selectedOption = props.options.find((option) => option.value === props.value) ?? props.options[0]

  return (
    <Select.Root
      items={props.options}
      value={props.value}
      onValueChange={(nextValue: string | null) => {
        if (typeof nextValue === 'string') {
          props.onChange(nextValue)
        }
      }}
    >
      <Select.Trigger aria-label={props.ariaLabel} className={props.triggerClassName} data-current-value={props.value}>
        <span className={props.valueClassName}>{selectedOption?.label ?? ''}</span>
        <Select.Icon className="ml-2 text-[var(--color-mauve-700)]">
          <ChevronDown className="h-3.5 w-3.5 stroke-[1.75]" />
        </Select.Icon>
      </Select.Trigger>
      <Select.Portal>
        <Select.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={6}>
          <Select.Popup className={toolbarPopupClass()}>
            <Select.List className="max-h-72 min-w-[var(--anchor-width)] overflow-auto py-1">
              {props.options.map((option) => (
                <Select.Item
                  className="flex cursor-default items-center justify-between gap-3 rounded-md px-2 py-1.5 text-[12px] text-[var(--color-mauve-900)] outline-none data-[highlighted]:bg-[var(--color-mauve-100)] data-[selected]:font-semibold"
                  key={`${props.ariaLabel}-${option.value || 'default'}`}
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
  )
}

export function WorkbookToolbarBorderMenu(props: {
  readonly disabled: boolean
  readonly isActive: boolean
  readonly onApplyPreset: (preset: BorderPreset) => void
}) {
  return (
    <Popover.Root modal={false}>
      <Popover.Trigger
        aria-label="Borders"
        aria-haspopup="menu"
        aria-pressed={props.isActive}
        className={cn(toolbarButtonClass({ active: props.isActive }), 'gap-1 px-1.5')}
        disabled={props.disabled}
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
                  onClick={() => props.onApplyPreset(key)}
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
  )
}

export function WorkbookToolbarStructureMenu(props: {
  readonly availability: StructureActionAvailability
  readonly disabled: boolean
  readonly onRunAction: (template: StructureActionTemplate) => void
}) {
  return (
    <Popover.Root modal={false}>
      <Popover.Trigger
        aria-label="Structure"
        aria-haspopup="menu"
        className={cn(toolbarButtonClass(), 'gap-1 px-1.5')}
        disabled={props.disabled || !hasAvailableStructureAction(props.availability)}
        title="Structure"
        type="button"
      >
        <TableProperties className={toolbarIconClass()} />
        <ChevronDown className="h-3 w-3 shrink-0 stroke-[1.75] text-[var(--wb-text-muted)]" />
      </Popover.Trigger>
      <Popover.Portal>
        <Popover.Positioner align="start" className="z-[1000]" side="bottom" sideOffset={8}>
          <Popover.Popup aria-label="Structure actions" className={cn(toolbarPopupClass(), 'w-[176px] p-1')}>
            <div className="grid gap-1">
              {STRUCTURE_ACTIONS.map((action) => {
                const isAvailable = props.availability[action.template]
                const ActionIcon = action.icon

                return (
                  <Popover.Close
                    aria-label={action.label}
                    className={cn(
                      'inline-flex h-8 items-center rounded-md border border-transparent px-2 text-left text-[11px] font-medium outline-none transition-colors focus-visible:border-[var(--color-mauve-400)] focus-visible:bg-[var(--color-mauve-100)]',
                      isAvailable
                        ? 'text-[var(--color-mauve-900)] hover:bg-[var(--color-mauve-100)]'
                        : 'cursor-not-allowed text-[var(--wb-text-muted)] opacity-45',
                    )}
                    disabled={!isAvailable}
                    key={action.key}
                    type="button"
                    onClick={() => props.onRunAction(action.template)}
                  >
                    {ActionIcon ? <ActionIcon className={cn(toolbarIconClass(), 'mr-2')} /> : null}
                    {action.label}
                  </Popover.Close>
                )
              })}
            </div>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  )
}
