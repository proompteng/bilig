import { cva } from 'class-variance-authority'

export const formulaBarRootClass = cva(
  'formula-bar flex items-start gap-2 border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2.5 py-1.5 font-sans',
)

export const formulaFieldShellClass = cva(
  'box-border flex h-8 min-h-8 items-stretch rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)] transition-[border-color,box-shadow]',
  {
    variants: {
      focused: {
        true: 'border-[var(--wb-accent)] ring-2 ring-[var(--wb-accent-ring)]',
        false: '',
      },
    },
    defaultVariants: {
      focused: false,
    },
  },
)

export const formulaFieldAddonClass = cva(
  'inline-flex h-full shrink-0 items-center justify-center border-r border-[var(--wb-border)] bg-[var(--wb-muted)] text-[11px] font-semibold uppercase tracking-[0.08em] leading-none text-[var(--wb-text-subtle)]',
)

export const formulaInputClass = cva(
  'min-h-8 min-w-0 flex-1 border-0 bg-transparent px-3 py-1 text-[12px] leading-4 text-[var(--wb-text)] outline-none placeholder:text-[var(--wb-text-subtle)]',
)

export const formulaStandaloneInputClass = cva(
  'box-border h-8 w-full rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-2.5 text-[12px] font-medium leading-none text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] outline-none transition-[border-color,box-shadow] placeholder:text-[var(--wb-text-subtle)] focus-visible:border-[var(--wb-accent)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]',
)

export const formulaPopupClass = cva(
  'overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-md)]',
)

export const formulaPopupOptionClass = cva('cursor-pointer px-3 py-2 transition-colors', {
  variants: {
    active: {
      true: 'bg-[var(--wb-muted)]',
      false: 'hover:bg-[var(--wb-surface-subtle)]',
    },
  },
  defaultVariants: {
    active: false,
  },
})

export const formulaHintClass = cva(
  'wb-scrollbar-none flex min-h-7 items-center gap-2 overflow-x-auto rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-2.5 text-[11px] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)]',
)
