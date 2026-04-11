import type { ReactNode } from "react";
import { Button } from "@base-ui/react/button";
import { cn } from "./cn.js";

const HEADER_ACTION_CLASS =
  "inline-flex h-8 items-center gap-2 rounded-md border px-2.5 text-[12px] font-medium transition-[background-color,border-color,color,box-shadow] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)]";

export const workbookHeaderSurfaceClass =
  "inline-flex h-8 items-center rounded-md border border-[var(--color-mauve-200)] bg-white shadow-[0_1px_2px_rgba(15,23,42,0.04)]";

export const workbookHeaderCountClass =
  "inline-flex min-w-4 items-center justify-center rounded-full bg-[var(--color-mauve-100)] px-1.5 text-[10px] font-semibold leading-none text-[var(--color-mauve-700)]";

export function WorkbookHeaderActionButton(props: {
  children: ReactNode;
  "aria-controls"?: string;
  "aria-expanded"?: boolean;
  "aria-label"?: string;
  "aria-pressed"?: boolean;
  "data-testid"?: string;
  disabled?: boolean;
  isActive?: boolean;
  isGrouped?: boolean;
  onClick: () => void;
  type?: "button" | "submit" | "reset";
}) {
  const grouped = props.isGrouped ?? false;
  return (
    <Button
      aria-controls={props["aria-controls"]}
      aria-expanded={props["aria-expanded"]}
      aria-label={props["aria-label"]}
      aria-pressed={props["aria-pressed"]}
      className={cn(
        HEADER_ACTION_CLASS,
        grouped
          ? props.isActive
            ? "border-[var(--color-mauve-300)] bg-white text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]"
            : "border-transparent bg-transparent text-[var(--color-mauve-600)] hover:bg-[var(--color-mauve-200)] hover:text-[var(--color-mauve-900)]"
          : props.isActive
            ? "border-[var(--color-mauve-300)] bg-[var(--color-mauve-50)] text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]"
            : "border-[var(--color-mauve-200)] bg-white text-[var(--color-mauve-700)] shadow-[0_1px_2px_rgba(15,23,42,0.04)] hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]",
        props.disabled ? "cursor-not-allowed opacity-50" : null,
      )}
      data-testid={props["data-testid"]}
      disabled={props.disabled}
      type={props.type ?? "button"}
      onClick={props.onClick}
    >
      {props.children}
    </Button>
  );
}

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
          "gap-2 px-2.5 text-[12px] font-medium text-[var(--color-mauve-700)]",
        )}
        data-testid="status-mode"
        role="status"
        title={`${props.modeLabel} • ${props.syncLabel}`}
      >
        <span aria-hidden="true" className={cn("size-2 rounded-full", toneClass)} />
        <span className="text-[var(--color-mauve-900)]">{props.syncLabel}</span>
      </span>
      <span className="sr-only" data-testid="status-sync">
        {props.syncLabel}
      </span>
    </>
  );
}
