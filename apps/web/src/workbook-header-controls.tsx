import { cva } from 'class-variance-authority'

export const workbookHeaderActionButtonClass = cva(
  'inline-flex h-8 items-center justify-center rounded-[var(--wb-radius-control)] border text-[12px] font-medium transition-[background-color,border-color,color,box-shadow] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface-subtle)] disabled:cursor-not-allowed disabled:opacity-50',
  {
    variants: {
      active: {
        true: '',
        false: '',
      },
      grouped: {
        true: '',
        false: '',
      },
      iconOnly: {
        true: 'w-8 gap-2 px-0',
        false: 'gap-2 px-2.5',
      },
    },
    compoundVariants: [
      {
        active: true,
        grouped: true,
        className: 'border-[var(--wb-border-strong)] bg-[var(--wb-surface)] text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)]',
      },
      {
        active: false,
        grouped: true,
        className:
          'border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-muted)] shadow-[inset_0_1px_0_rgba(255,255,255,0.7)] hover:border-[var(--wb-border-strong)] hover:bg-[var(--wb-surface)] hover:text-[var(--wb-text)]',
      },
      {
        active: true,
        grouped: false,
        className: 'border-[var(--wb-border-strong)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)]',
      },
      {
        active: false,
        grouped: false,
        className:
          'border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] hover:bg-[var(--wb-muted)] hover:text-[var(--wb-text)]',
      },
    ],
    defaultVariants: {
      active: false,
      grouped: false,
      iconOnly: false,
    },
  },
)

export const workbookHeaderSurfaceClass =
  'inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)]'
