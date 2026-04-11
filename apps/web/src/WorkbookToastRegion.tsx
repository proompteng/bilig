import { useEffect, useEffectEvent, useLayoutEffect, useRef } from "react";
import type { RefObject } from "react";
import { Toaster, toast } from "sonner";
import type { ExternalToast } from "sonner";
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

function buildToastOptions(
  workbookToast: WorkbookToast,
  ignoredDismissIdsRef: RefObject<Set<string>>,
): ExternalToast {
  return {
    id: workbookToast.id,
    toasterId: WORKBOOK_TOASTER_ID,
    duration: Number.POSITIVE_INFINITY,
    testId: `workbook-toast-${workbookToast.id}`,
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
