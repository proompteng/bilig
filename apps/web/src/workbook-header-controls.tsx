import { cva } from 'class-variance-authority'
import { cn } from './cn.js'

export const workbookHeaderActionButtonClass = cva(
  'inline-flex h-8 items-center justify-center rounded-md border text-[12px] font-medium transition-[background-color,border-color,color,box-shadow] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)] disabled:cursor-not-allowed disabled:opacity-50',
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
        className: 'border-[var(--color-mauve-300)] bg-white text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]',
      },
      {
        active: false,
        grouped: true,
        className:
          'border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] text-[var(--color-mauve-700)] shadow-[inset_0_1px_0_rgba(255,255,255,0.7)] hover:border-[var(--color-mauve-300)] hover:bg-white hover:text-[var(--color-mauve-900)]',
      },
      {
        active: true,
        grouped: false,
        className:
          'border-[var(--color-mauve-300)] bg-[var(--color-mauve-50)] text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]',
      },
      {
        active: false,
        grouped: false,
        className:
          'border-[var(--color-mauve-200)] bg-white text-[var(--color-mauve-700)] shadow-[0_1px_2px_rgba(15,23,42,0.04)] hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]',
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
  'inline-flex h-8 items-center rounded-md border border-[var(--color-mauve-200)] bg-white shadow-[0_1px_2px_rgba(15,23,42,0.04)]'

export const workbookHeaderCountClass =
  'inline-flex min-w-4 items-center justify-center rounded-full bg-[var(--color-mauve-100)] px-1.5 text-[10px] font-semibold leading-none text-[var(--color-mauve-700)]'

export function WorkbookHeaderStatusChip(props: { modeLabel: string; syncLabel: string }) {
  const toneClass =
    props.syncLabel === 'Ready'
      ? 'bg-[var(--color-mauve-900)]'
      : props.syncLabel === 'Syncing' || props.syncLabel === 'Loading'
        ? 'bg-[var(--color-mauve-500)]'
        : 'bg-[var(--color-mauve-400)]'

  return (
    <>
      <span
        aria-label={`Workbook status: ${props.modeLabel}, ${props.syncLabel}`}
        className="inline-flex h-8 items-center justify-center rounded-md px-1.5 text-[var(--color-mauve-600)]"
        data-testid="status-mode"
        role="status"
        title={`${props.modeLabel} • ${props.syncLabel}`}
      >
        <span aria-hidden="true" className={cn('size-2 rounded-full', toneClass)} />
      </span>
      <span className="sr-only" data-testid="status-sync">
        {props.syncLabel}
      </span>
    </>
  )
}
