import { cva } from "class-variance-authority";

export const workbookPillClass = cva(
  "inline-flex h-5 items-center rounded-full border px-2 text-[10px] uppercase tracking-[0.04em]",
  {
    variants: {
      tone: {
        neutral:
          "border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
        accent: "border-[var(--wb-border)] bg-[var(--wb-surface-muted)] text-[var(--wb-text)]",
        danger: "border-[#efc7c7] bg-[#fff6f6] text-[#8f2d2d]",
      },
      weight: {
        regular: "font-medium",
        strong: "font-semibold",
      },
    },
    defaultVariants: {
      tone: "neutral",
      weight: "regular",
    },
  },
);

export const workbookButtonClass = cva(
  "inline-flex items-center justify-center rounded-[var(--wb-radius-control)] border transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50",
  {
    variants: {
      tone: {
        neutral:
          "border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-muted)] hover:bg-[var(--wb-surface)] hover:text-[var(--wb-text)]",
        accent:
          "border-[var(--wb-border-strong)] bg-[var(--wb-surface)] text-[var(--wb-text)] hover:bg-[var(--wb-surface-muted)]",
        danger: "border-[#efc7c7] bg-[#fffafa] text-[#8f2d2d] hover:bg-[#fff6f6]",
      },
      size: {
        sm: "h-8 px-3 text-[12px]",
        md: "h-10 px-3 text-[12px]",
      },
      weight: {
        regular: "font-medium",
        strong: "font-semibold",
      },
    },
    defaultVariants: {
      tone: "neutral",
      size: "sm",
      weight: "regular",
    },
  },
);
