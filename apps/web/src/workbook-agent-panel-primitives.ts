import { cva } from 'class-variance-authority'

export const agentPanelFooterClass = cva('border-t border-[var(--wb-border)] bg-[var(--wb-app-bg)] px-3 py-2.5')

export const agentPanelEyebrowTextClass = cva('text-[10px] leading-4 font-medium uppercase tracking-[0.06em] text-[var(--wb-text-subtle)]')

export const agentPanelLabelTextClass = cva('text-[13px] leading-5 font-medium text-[var(--wb-text)]')

export const agentPanelMetaTextClass = cva('text-[12px] leading-[1.45] text-[var(--wb-text-subtle)]')

export const agentPanelBodyTextClass = cva('text-[13px] leading-[1.65] text-[var(--wb-text)]')

export const agentPanelBodyMutedTextClass = cva('text-[13px] leading-[1.65] text-[var(--wb-text-muted)]')

export const agentPanelComposerFrameClass = cva(
  'relative rounded-[calc(var(--wb-radius-control)+2px)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[inset_0_1px_0_rgba(255,255,255,0.55)] transition-colors focus-within:border-[var(--wb-border-strong)] focus-within:bg-[var(--wb-surface)]',
)

export const agentPanelComposerScrollRootClass = cva('relative overflow-hidden rounded-[calc(var(--wb-radius-control)+1px)]')

export const agentPanelComposerScrollViewportClass = cva('w-full overflow-x-hidden pr-11')

export const agentPanelComposerScrollContentClass = cva('min-w-0 w-full')

export const agentPanelComposerTextareaClass = cva(
  'block w-full resize-none overflow-hidden border-0 bg-transparent px-3 py-3 text-[13px] leading-[1.65] text-[var(--wb-text)] placeholder:text-[var(--wb-text-subtle)] outline-none',
)

export const agentPanelComposerSendButtonClass = cva(
  'absolute right-3 bottom-3 inline-flex h-8 w-8 items-center justify-center rounded-full border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:bg-[var(--wb-surface-muted)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-surface-muted)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:border-[var(--wb-border)] disabled:bg-[var(--wb-surface-subtle)] disabled:text-[var(--wb-text-subtle)] disabled:shadow-none disabled:opacity-70',
)

export const agentPanelThreadListClass = cva('mt-2 flex flex-col gap-1.5')

export const agentPanelThreadButtonClass = cva(
  'flex w-full min-w-0 items-start justify-between gap-3 rounded-[var(--wb-radius-control)] border px-3 py-2.5 text-left transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1',
  {
    variants: {
      active: {
        true: 'border-[var(--wb-border-strong)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)]',
        false: 'border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] hover:bg-[var(--wb-surface)]',
      },
    },
    defaultVariants: {
      active: false,
    },
  },
)

export const agentPanelScrollAreaRootClass = cva('relative min-h-0 flex-1 overflow-hidden')

export const agentPanelScrollAreaViewportClass = cva('h-full w-full overflow-x-hidden')

export const agentPanelScrollAreaContentClass = cva('min-h-full min-w-0 w-full')

export const agentPanelScrollAreaScrollbarClass = cva(
  'flex touch-none select-none p-0.5 transition-opacity data-[orientation=vertical]:w-2.5 data-[orientation=horizontal]:h-2.5',
)

export const agentPanelScrollAreaThumbClass = cva('relative flex-1 rounded-full bg-[var(--wb-border-strong)]')

export const agentPanelTimelineListClass = cva('flex min-w-0 flex-col gap-2')

export const agentPanelDisclosureFrameClass = cva('min-w-0 w-full max-w-full space-y-2 overflow-hidden transition-colors', {
  variants: {
    open: {
      true: '',
      false: '',
    },
  },
  defaultVariants: {
    open: false,
  },
})

export const agentPanelDisclosureTriggerClass = cva(
  'grid min-h-11 min-w-0 w-full max-w-full grid-cols-[auto_minmax(0,1fr)_auto] items-start gap-x-2.5 overflow-hidden rounded-[var(--wb-radius-control)] px-2.5 py-1 text-left outline-none ring-0 transition-colors hover:bg-[var(--wb-surface-subtle)] active:bg-transparent focus:outline-none focus-visible:outline-none focus-visible:ring-0 focus-visible:shadow-none [webkit-tap-highlight-color:transparent]',
)

export const agentPanelDisclosureChevronClass = cva('mt-0.5 size-3.5 shrink-0 text-[var(--wb-text-subtle)] transition-transform', {
  variants: {
    open: {
      true: 'rotate-90',
      false: '',
    },
  },
  defaultVariants: {
    open: false,
  },
})

export const agentPanelDisclosureContentClass = cva('min-w-0 w-full max-w-full overflow-hidden', {
  variants: {
    open: {
      true: 'grid grid-cols-1 gap-y-0.5',
      false: 'flex flex-wrap items-start gap-x-1.5 gap-y-0.5',
    },
  },
  defaultVariants: {
    open: false,
  },
})

export const agentPanelDisclosureLabelClass = cva('text-[13px] leading-5 font-medium text-[var(--wb-text)]')

export const agentPanelDisclosureSummaryClass = cva(
  'block min-w-0 w-full max-w-full overflow-hidden text-[12px] leading-[1.45] text-[var(--wb-text-muted)]',
  {
    variants: {
      open: {
        true: 'whitespace-normal break-all',
        false: 'truncate whitespace-nowrap',
      },
    },
    defaultVariants: {
      open: false,
    },
  },
)

export const agentPanelDisclosureBadgeClass = cva('min-w-0 shrink-0 self-start justify-self-end leading-none')

export const agentPanelDisclosurePanelClass = cva('min-w-0 w-full max-w-full overflow-hidden')

export const agentPanelDisclosureViewportClass = cva('h-44 min-w-0 w-full max-w-full')

export const agentPanelDisclosureBodyClass = cva('min-w-0 w-full max-w-full px-2 pb-2')

export const agentPanelDisclosureBodyCardClass = cva(
  'min-w-0 w-full max-w-full overflow-x-hidden rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-muted)] px-3.5 py-3',
)
