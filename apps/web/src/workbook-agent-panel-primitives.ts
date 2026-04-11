import { cva } from "class-variance-authority";

export const agentPanelSectionClass = cva(
  "rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2",
);

export const agentPanelSectionHeaderClass = cva("flex items-center justify-between gap-2");

export const agentPanelSectionTitleClass = cva(
  "text-[11px] font-semibold tracking-[0.01em] text-[var(--wb-text)]",
);

export const agentPanelSectionHintClass = cva("text-[11px] text-[var(--wb-text-muted)]");

export const agentPanelFieldClass = cva(
  "min-w-0 w-full rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-app-bg)] px-2.5 py-1.5 text-[12px] text-[var(--wb-text)] outline-none transition-colors placeholder:text-[var(--wb-text-muted)] focus:border-[var(--wb-accent)] focus:ring-2 focus:ring-[var(--wb-accent-ring)]",
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
