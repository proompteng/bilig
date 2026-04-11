import { cva } from "class-variance-authority";

export const agentPanelHeaderClass = cva(
  "border-b border-[var(--wb-border)] bg-[var(--wb-app-bg)] px-3 py-2.5",
);

export const agentPanelFooterClass = cva(
  "border-t border-[var(--wb-border)] bg-[var(--wb-app-bg)] px-3 py-2.5",
);

export const agentPanelToolbarRowClass = cva("flex flex-wrap items-center justify-between gap-2");

export const agentPanelSectionClass = cva(
  "rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2",
);

export const agentPanelSectionHeaderClass = cva("flex items-center justify-between gap-2");

export const agentPanelSectionTitleClass = cva(
  "text-[11px] font-semibold tracking-[0.01em] text-[var(--wb-text)]",
);

export const agentPanelSectionHintClass = cva("text-[11px] text-[var(--wb-text-muted)]");

export const agentPanelFieldClass = cva(
  "min-w-0 w-full rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-2.5 py-1.5 text-[12px] text-[var(--wb-text)] outline-none transition-colors placeholder:text-[var(--wb-text-muted)] focus:border-[var(--wb-border-strong)] focus:bg-[var(--wb-surface)] focus:ring-2 focus:ring-[var(--wb-surface-muted)]",
  {
    variants: {
      multiline: {
        false: "",
        true: "px-3 py-3 leading-5",
      },
    },
    defaultVariants: {
      multiline: false,
    },
  },
);

export const agentPanelActionGridClass = cva("grid grid-cols-2 gap-2");

export const agentPanelActionButtonClass = cva(
  "h-auto w-full justify-center px-3 py-2 text-center leading-4 whitespace-normal",
  {
    variants: {
      emphasis: {
        subtle: "",
        strong: "",
      },
    },
    defaultVariants: {
      emphasis: "subtle",
    },
  },
);

export const agentPanelToggleButtonClass = cva("min-w-[4.5rem]");

export const agentPanelComposerFrameClass = cva(
  "relative rounded-[calc(var(--wb-radius-control)+2px)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-2 shadow-[inset_0_1px_0_rgba(255,255,255,0.55)] transition-colors focus-within:border-[var(--wb-border-strong)] focus-within:bg-[var(--wb-surface)]",
);

export const agentPanelComposerTextareaClass = cva(
  "min-h-24 w-full resize-none border-0 bg-transparent px-1 py-1 pr-14 text-[13px] leading-5 text-[var(--wb-text)] placeholder:text-[var(--wb-text-subtle)] outline-none",
);

export const agentPanelComposerSendButtonClass = cva(
  "absolute right-2 bottom-2 inline-flex h-9 w-9 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:bg-[var(--wb-surface-muted)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-surface-muted)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:border-[var(--wb-border)] disabled:bg-[var(--wb-surface-subtle)] disabled:text-[var(--wb-text-subtle)] disabled:shadow-none disabled:opacity-70",
);

export const agentPanelSegmentedGroupClass = cva(
  "inline-flex items-center gap-1 rounded-[calc(var(--wb-radius-control)+2px)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] p-1",
);

export const agentPanelSegmentedButtonClass = cva(
  "inline-flex h-8 items-center justify-center rounded-[var(--wb-radius-control)] px-2.5 text-[12px] font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-surface-muted)] focus-visible:ring-offset-1",
  {
    variants: {
      active: {
        true: "border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)]",
        false:
          "border border-transparent bg-transparent text-[var(--wb-text-subtle)] hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)]",
      },
    },
    defaultVariants: {
      active: false,
    },
  },
);

export const agentPanelInlineButtonClass = cva(
  "inline-flex h-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-surface-muted)] focus-visible:ring-offset-1",
);

export const agentPanelThreadListClass = cva("mt-2 flex flex-col gap-1.5");

export const agentPanelThreadButtonClass = cva(
  "flex w-full min-w-0 items-start justify-between gap-3 rounded-[var(--wb-radius-control)] border px-3 py-2 text-left transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1",
  {
    variants: {
      active: {
        true: "border-[var(--wb-border-strong)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)]",
        false:
          "border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] hover:bg-[var(--wb-surface)]",
      },
    },
    defaultVariants: {
      active: false,
    },
  },
);

export const agentPanelToolsPanelClass = cva(
  "mt-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] p-2",
);
