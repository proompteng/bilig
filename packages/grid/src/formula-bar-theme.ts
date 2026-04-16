import { cva } from 'class-variance-authority'

export const formulaBarRootClass = cva(
  'formula-bar flex items-start gap-2 border-b border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] px-2.5 py-1.5 font-sans',
)

export const formulaFieldShellClass = cva(
  'box-border flex h-8 items-center rounded-md border border-[var(--color-mauve-200)] bg-white shadow-[0_1px_2px_rgba(15,23,42,0.04)] transition-[border-color,box-shadow]',
  {
    variants: {
      focused: {
        true: 'border-[var(--color-mauve-400)] ring-2 ring-[var(--color-mauve-400)]/40',
        false: '',
      },
    },
    defaultVariants: {
      focused: false,
    },
  },
)

export const formulaFieldAddonClass = cva(
  'inline-flex h-full shrink-0 items-center justify-center border-r border-[var(--color-mauve-200)] bg-[var(--color-mauve-100)] text-[11px] font-semibold uppercase tracking-[0.08em] leading-none text-[var(--color-mauve-600)]',
)

export const formulaInputClass = cva(
  'h-full min-w-0 flex-1 border-0 bg-transparent px-3 text-[12px] leading-none text-[var(--color-mauve-950)] outline-none placeholder:text-[var(--color-mauve-500)]',
)

export const formulaStandaloneInputClass = cva(
  'box-border h-8 w-full rounded-md border border-[var(--color-mauve-200)] bg-white px-2.5 text-[12px] font-medium leading-none text-[var(--color-mauve-950)] shadow-[0_1px_2px_rgba(15,23,42,0.04)] outline-none transition-[border-color,box-shadow] placeholder:text-[var(--color-mauve-500)] focus-visible:border-[var(--color-mauve-400)] focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)]/40',
)

export const formulaPopupClass = cva(
  'overflow-hidden rounded-xl border border-[var(--color-mauve-200)] bg-white shadow-[0_14px_32px_rgba(15,23,42,0.14)]',
)

export const formulaPopupOptionClass = cva('cursor-pointer px-3 py-2 transition-colors', {
  variants: {
    active: {
      true: 'bg-[var(--color-mauve-100)]',
      false: 'hover:bg-[var(--color-mauve-50)]',
    },
  },
  defaultVariants: {
    active: false,
  },
})

export const formulaHintClass = cva(
  'wb-scrollbar-none flex min-h-7 items-center gap-2 overflow-x-auto rounded-md border border-[var(--color-mauve-200)] bg-white px-2.5 text-[11px] text-[var(--color-mauve-600)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]',
)
