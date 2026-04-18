import { cva } from 'class-variance-authority'
import { cn } from './cn.js'

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
        true: 'w-8 gap-0 px-0',
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

export const workbookHeaderCountClass =
  'inline-flex min-w-4 items-center justify-center rounded-full bg-[var(--wb-muted)] px-1.5 text-[10px] font-semibold leading-none text-[var(--wb-text-muted)]'

interface WorkbookHeaderStatusChipProps {
  modeLabel: string
  syncLabel: string
  tone?: 'positive' | 'progress' | 'warning' | 'danger' | 'neutral'
}

export function WorkbookHeaderStatusChip({ modeLabel, syncLabel, tone = 'neutral' }: WorkbookHeaderStatusChipProps) {
  const isDotOnlySavedState = syncLabel === 'Saved'
  const toneClass =
    tone === 'positive'
      ? 'bg-[var(--wb-success)]'
      : tone === 'progress'
        ? 'bg-[var(--wb-accent)]'
        : tone === 'warning'
          ? 'bg-[var(--wb-warning)]'
          : tone === 'danger'
            ? 'bg-[var(--wb-danger)]'
            : 'bg-[var(--wb-text-subtle)]'

  const surfaceClass =
    tone === 'positive'
      ? 'text-[var(--wb-success)]'
      : tone === 'progress'
        ? 'text-[var(--wb-accent)]'
        : tone === 'warning'
          ? 'text-[var(--wb-warning)]'
          : tone === 'danger'
            ? 'text-[var(--wb-danger-text)]'
            : 'text-[var(--wb-text-muted)]'

  return (
    <>
      <span
        aria-label={`Workbook status: ${modeLabel}, ${syncLabel}`}
        className={`inline-flex h-8 items-center px-0 text-[12px] font-medium ${isDotOnlySavedState ? 'gap-0' : 'gap-2'} ${surfaceClass}`}
        data-testid="status-mode"
        role="status"
        title={`${modeLabel} • ${syncLabel}`}
      >
        <span aria-hidden="true" className={cn('size-2 rounded-full', toneClass)} />
        {isDotOnlySavedState ? null : <span data-testid="status-label">{syncLabel}</span>}
      </span>
      <span className="sr-only" data-testid="status-sync">
        {syncLabel}
      </span>
    </>
  )
}
