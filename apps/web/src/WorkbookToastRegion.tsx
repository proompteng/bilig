import { useEffect, useEffectEvent, useLayoutEffect, useRef } from "react";
import type { RefObject } from "react";
import { Toaster, toast } from "sonner";
import type { ExternalToast, ToastClassnames } from "sonner";
import {
  workbookButtonClass,
  workbookSurfaceClass,
  workbookAlertClass,
} from "./workbook-shell-chrome.js";
import "sonner/dist/styles.css";

export interface WorkbookToastAction {
  readonly label: string;
  readonly onAction: () => void;
}

export interface WorkbookToast {
  readonly id: string;
  readonly tone?: "error" | "neutral";
  readonly message: string;
  readonly action?: WorkbookToastAction;
  readonly onDismiss?: () => void;
}

const WORKBOOK_TOASTER_ID = "workbook";

function toastClassNames(tone: WorkbookToast["tone"]): ToastClassnames {
  const isError = tone === "error";
  return {
    toast: [
      "pointer-events-auto flex items-start gap-3 rounded-[var(--wb-radius-panel)] px-3 py-3 shadow-[var(--wb-shadow-md)]",
      isError
        ? workbookAlertClass({ tone: "danger" })
        : `${workbookSurfaceClass({ emphasis: "raised" })} text-[var(--wb-text)]`,
    ].join(" "),
    content: "min-w-0 flex-1",
    title: "break-words text-[12px] leading-5",
    actionButton: workbookButtonClass({
      size: "sm",
      tone: isError ? "danger" : "neutral",
    }),
    closeButton: isError
      ? "inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-full border border-[#ead0d0] text-[#8f2d2d] transition-colors hover:bg-[#fff3f3] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#ead0d0] focus-visible:ring-offset-1"
      : "inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-full border border-[var(--wb-border)] text-[var(--wb-text-muted)] transition-colors hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1",
  };
}

function buildToastOptions(
  workbookToast: WorkbookToast,
  ignoredDismissIdsRef: RefObject<Set<string>>,
): ExternalToast {
  return {
    id: workbookToast.id,
    toasterId: WORKBOOK_TOASTER_ID,
    duration: Number.POSITIVE_INFINITY,
    unstyled: true,
    icon: null,
    testId: `workbook-toast-${workbookToast.id}`,
    classNames: toastClassNames(workbookToast.tone),
    closeButton: workbookToast.onDismiss !== undefined,
    dismissible: workbookToast.onDismiss !== undefined,
    ...(workbookToast.action
      ? {
          action: {
            label: workbookToast.action.label,
            onClick: () => {
              workbookToast.action?.onAction();
            },
          },
        }
      : {}),
    ...(workbookToast.onDismiss
      ? {
          onDismiss: () => {
            if (ignoredDismissIdsRef.current?.delete(workbookToast.id)) {
              return;
            }
            workbookToast.onDismiss?.();
          },
        }
      : {}),
  };
}

function dismissToastProgrammatically(
  id: string,
  ignoredDismissIdsRef: RefObject<Set<string>>,
): void {
  ignoredDismissIdsRef.current?.add(id);
  toast.dismiss(id);
}

export function WorkbookToastRegion(props: { readonly toasts: readonly WorkbookToast[] }) {
  const activeIdsRef = useRef(new Set<string>());
  const ignoredDismissIdsRef = useRef(new Set<string>());

  const dismissActiveToasts = useEffectEvent(() => {
    for (const activeId of activeIdsRef.current) {
      dismissToastProgrammatically(activeId, ignoredDismissIdsRef);
    }
    activeIdsRef.current.clear();
  });

  useLayoutEffect(() => {
    const nextIds = new Set(props.toasts.map((workbookToast) => workbookToast.id));
    for (const activeId of activeIdsRef.current) {
      if (nextIds.has(activeId)) {
        continue;
      }
      dismissToastProgrammatically(activeId, ignoredDismissIdsRef);
      activeIdsRef.current.delete(activeId);
    }

    for (const workbookToast of props.toasts) {
      const options = buildToastOptions(workbookToast, ignoredDismissIdsRef);
      if (workbookToast.tone === "error") {
        toast.error(workbookToast.message, options);
      } else {
        toast(workbookToast.message, options);
      }
      activeIdsRef.current.add(workbookToast.id);
    }
  }, [props.toasts]);

  useEffect(() => {
    return () => {
      dismissActiveToasts();
    };
  }, []);

  return (
    <Toaster
      closeButton={false}
      containerAriaLabel="Workbook notifications"
      id={WORKBOOK_TOASTER_ID}
      offset={12}
      position="top-right"
      toastOptions={{
        closeButtonAriaLabel: "Dismiss",
      }}
      visibleToasts={4}
    />
  );
}
