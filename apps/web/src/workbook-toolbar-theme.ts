import { cva } from 'class-variance-authority'

export interface ToolbarSelectOption {
  label: string
  value: string
}

export const NUMBER_FORMAT_OPTIONS: readonly ToolbarSelectOption[] = [
  { label: 'General', value: 'general' },
  { label: 'Number', value: 'number' },
  { label: 'Currency', value: 'currency' },
  { label: 'Accounting', value: 'accounting' },
  { label: 'Percent', value: 'percent' },
  { label: 'Date', value: 'date' },
  { label: 'Text', value: 'text' },
] as const

export const FONT_SIZE_OPTIONS: readonly ToolbarSelectOption[] = [10, 11, 12, 13, 14, 16, 18, 20].map((size) => ({
  label: String(size),
  value: String(size),
}))

export const toolbarRootClass = cva('border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] font-sans')

export const toolbarRowClass = cva('mx-0 flex h-10 items-center gap-1 overflow-hidden px-2.5 py-0 text-[12px] text-[var(--wb-text)]')

export const toolbarFormattingScrollClass = cva(
  'wb-scrollbar-none flex min-w-0 flex-1 items-center gap-1 overflow-x-auto overflow-y-hidden',
)

export const toolbarFormattingRegionClass = cva('relative flex min-w-0 flex-1 items-center')

export const toolbarOverflowCueClass = cva(
  'absolute inset-y-0 right-0 z-10 inline-flex w-7 items-center justify-center border-l border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-accent)] outline-none transition-colors hover:bg-[var(--wb-muted)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface-subtle)]',
)

export const toolbarTrailingRegionClass = cva('ml-auto flex flex-none items-center gap-1.5 pl-2')

export const toolbarGroupClass = cva('flex flex-none items-center gap-1')

export const toolbarSeparatorClass = cva('mx-1.5 h-5 w-px shrink-0 bg-[var(--wb-border)]')

export const toolbarSegmentedClass = cva('inline-flex h-8 items-center gap-0.5')

export const toolbarIconClass = cva('h-3.5 w-3.5 shrink-0 stroke-[1.75]')

export const toolbarButtonClass = cva(
  'inline-flex items-center justify-center rounded-[var(--wb-radius-control)] border text-[var(--wb-text-muted)] transition-[background-color,border-color,color,box-shadow] outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface-subtle)] disabled:cursor-default disabled:opacity-60',
  {
    variants: {
      active: {
        true: 'border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] text-[var(--wb-accent)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]',
        false: 'border-transparent bg-transparent hover:border-[var(--wb-border)] hover:bg-[var(--wb-muted)] hover:text-[var(--wb-text)]',
      },
      embedded: {
        true: 'h-8 min-w-8 px-1.5',
        false: 'h-8 min-w-8 px-1.5',
      },
    },
    defaultVariants: {
      active: false,
      embedded: false,
    },
  },
)

export const toolbarSelectTriggerClass = cva(
  'inline-flex h-8 items-center justify-between gap-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-2 text-[12px] font-medium text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] outline-none transition-[background-color,border-color,color,box-shadow] hover:bg-[var(--wb-muted)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface-subtle)] disabled:cursor-default disabled:opacity-60',
)

export const toolbarBorderIconClass = cva('text-[20px] leading-none text-[var(--wb-text-muted)]')

export const toolbarPopupClass = cva(
  'overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-1.5 shadow-[var(--wb-shadow-md)]',
)

export const toolbarBorderPopupClass = cva(
  'overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-1.5 py-1.5 shadow-[var(--wb-shadow-md)]',
)

export const toolbarPopupActionClass = cva(
  'inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-2 text-[11px] font-semibold text-[var(--wb-text)] transition-[background-color,border-color,color] hover:bg-[var(--wb-muted)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface)] disabled:opacity-50',
)

export const colorPickerPopupClass = cva(
  'overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-2 shadow-[var(--wb-shadow-md)]',
)

export const colorPickerSwatchClass = cva(
  'relative border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] outline-none transition-colors hover:border-[var(--wb-accent)] focus-visible:border-[var(--wb-accent)] focus-visible:ring-1 focus-visible:ring-[var(--wb-accent)]',
)

export function classNames(...values: Array<string | false | null | undefined>): string {
  return values.filter(Boolean).join(' ')
}
