import { cva } from "class-variance-authority";
import { cn } from "./cn.js";

export const workbookHeaderActionButtonClass = cva(
  "inline-flex h-8 items-center gap-2 rounded-md border px-2.5 text-[12px] font-medium transition-[background-color,border-color,color,box-shadow] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)] disabled:cursor-not-allowed disabled:opacity-50",
  {
    variants: {
      active: {
        true: "",
        false: "",
      },
      grouped: {
        true: "",
        false: "",
      },
    },
    compoundVariants: [
      {
        active: true,
        grouped: true,
        className:
          "border-[var(--color-mauve-300)] bg-white text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]",
      },
      {
        active: false,
        grouped: true,
        className:
          "border-transparent bg-transparent text-[var(--color-mauve-600)] hover:bg-[var(--color-mauve-200)] hover:text-[var(--color-mauve-900)]",
      },
      {
        active: true,
        grouped: false,
        className:
          "border-[var(--color-mauve-300)] bg-[var(--color-mauve-50)] text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]",
      },
      {
        active: false,
        grouped: false,
        className:
          "border-[var(--color-mauve-200)] bg-white text-[var(--color-mauve-700)] shadow-[0_1px_2px_rgba(15,23,42,0.04)] hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]",
      },
    ],
    defaultVariants: {
      active: false,
      grouped: false,
    },
  },
);

export const workbookHeaderSurfaceClass =
  "inline-flex h-8 items-center rounded-md border border-[var(--color-mauve-200)] bg-white shadow-[0_1px_2px_rgba(15,23,42,0.04)]";

export const workbookHeaderCountClass =
  "inline-flex min-w-4 items-center justify-center rounded-full bg-[var(--color-mauve-100)] px-1.5 text-[10px] font-semibold leading-none text-[var(--color-mauve-700)]";

export function WorkbookHeaderStatusChip(props: { modeLabel: string; syncLabel: string }) {
  const toneClass =
    props.syncLabel === "Ready"
      ? "bg-[var(--color-mauve-900)]"
      : props.syncLabel === "Syncing" || props.syncLabel === "Loading"
        ? "bg-[var(--color-mauve-500)]"
        : "bg-[var(--color-mauve-400)]";

  return (
    <>
      <span
        aria-label={`Workbook status: ${props.modeLabel}, ${props.syncLabel}`}
        className={cn(
          workbookHeaderSurfaceClass,
          "w-8 justify-center text-[var(--color-mauve-700)]",
        )}
        data-testid="status-mode"
        role="status"
        title={`${props.modeLabel} • ${props.syncLabel}`}
      >
        <span aria-hidden="true" className={cn("size-2 rounded-full", toneClass)} />
      </span>
      <span className="sr-only" data-testid="status-sync">
        {props.syncLabel}
      </span>
    </>
  );
}
