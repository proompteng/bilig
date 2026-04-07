import { cva } from "class-variance-authority";
import { cn } from "./cn.js";

export interface WorkbookToast {
  readonly id: string;
  readonly tone?: "error" | "neutral";
  readonly message: string;
  readonly onDismiss?: () => void;
}

const toastViewportClass = cva(
  "pointer-events-none absolute top-3 right-3 z-40 flex w-[min(26rem,calc(100%-1.5rem))] flex-col gap-2",
);

const toastClass = cva(
  "pointer-events-auto flex items-start gap-3 rounded-[var(--wb-radius-panel)] border bg-[var(--wb-surface)] px-3 py-3 shadow-[var(--wb-shadow-md)]",
  {
    variants: {
      tone: {
        error: "border-[#f1b5b5] bg-[#fff7f7] text-[#991b1b]",
        neutral: "border-[var(--wb-border)] text-[var(--wb-text)]",
      },
    },
    defaultVariants: {
      tone: "neutral",
    },
  },
);

const toastDismissClass = cva(
  "inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-full border transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-1",
  {
    variants: {
      tone: {
        error:
          "border-[#f1b5b5] text-[#991b1b] hover:border-[#e58e8e] focus-visible:ring-[#f1b5b5]",
        neutral:
          "border-[var(--wb-border)] text-[var(--wb-text-muted)] hover:text-[var(--wb-text)] focus-visible:ring-[var(--wb-accent-ring)]",
      },
    },
    defaultVariants: {
      tone: "neutral",
    },
  },
);

export function WorkbookToastRegion(props: { readonly toasts: readonly WorkbookToast[] }) {
  if (props.toasts.length === 0) {
    return null;
  }

  return (
    <div aria-live="polite" className={toastViewportClass()} data-testid="workbook-toast-region">
      {props.toasts.map((toast) => (
        <div
          className={cn(toastClass({ tone: toast.tone ?? "neutral" }))}
          data-testid={`workbook-toast-${toast.id}`}
          key={toast.id}
          role="status"
        >
          <div className="min-w-0 flex-1 break-words text-[12px] leading-5">{toast.message}</div>
          {toast.onDismiss ? (
            <button
              aria-label="Dismiss"
              className={cn(toastDismissClass({ tone: toast.tone ?? "neutral" }))}
              type="button"
              onClick={toast.onDismiss}
            >
              ×
            </button>
          ) : null}
        </div>
      ))}
    </div>
  );
}
