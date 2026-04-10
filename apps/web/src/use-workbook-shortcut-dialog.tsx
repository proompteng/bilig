import { useCallback, useEffect, useMemo, useState } from "react";
import { WorkbookHeaderActionButton } from "./workbook-header-controls.js";
import { WorkbookShortcutDialog } from "./WorkbookShortcutDialog.js";
import { isTextEntryTarget } from "./worker-workbook-app-model.js";

export function useWorkbookShortcutDialog() {
  const [isOpen, setIsOpen] = useState(false);
  const [query, setQuery] = useState("");

  const closeShortcutDialog = useCallback(() => {
    setIsOpen(false);
    setQuery("");
  }, []);

  const openShortcutDialog = useCallback(() => {
    setIsOpen(true);
  }, []);

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      if (
        event.defaultPrevented ||
        event.altKey ||
        event.ctrlKey ||
        event.metaKey ||
        isTextEntryTarget(event.target)
      ) {
        return;
      }
      if (event.key !== "?") {
        return;
      }
      event.preventDefault();
      setIsOpen(true);
    };

    window.addEventListener("keydown", handleWindowKeyDown, true);
    return () => {
      window.removeEventListener("keydown", handleWindowKeyDown, true);
    };
  }, []);

  const shortcutHelpButton = useMemo(
    () => (
      <WorkbookHeaderActionButton
        aria-label="Show keyboard shortcuts"
        data-testid="workbook-shortcut-button"
        onClick={openShortcutDialog}
      >
        <span className="text-[11px] font-semibold">Shortcuts</span>
      </WorkbookHeaderActionButton>
    ),
    [openShortcutDialog],
  );

  const shortcutDialog = useMemo(
    () => (
      <WorkbookShortcutDialog
        open={isOpen}
        query={query}
        onOpenChange={(open) => {
          if (open) {
            setIsOpen(true);
            return;
          }
          closeShortcutDialog();
        }}
        onQueryChange={setQuery}
      />
    ),
    [closeShortcutDialog, isOpen, query],
  );

  return {
    closeShortcutDialog,
    openShortcutDialog,
    shortcutDialog,
    shortcutHelpButton,
  };
}
