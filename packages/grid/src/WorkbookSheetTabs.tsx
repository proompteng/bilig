import React from "react";
import { Button } from "@base-ui/react/button";
import { Tabs } from "@base-ui/react/tabs";
import { cn } from "./cn.js";

interface WorkbookSheetTabsProps {
  sheetName: string;
  sheetNames: string[];
  selectionStatus?: React.ReactNode;
  onSelectSheet(this: void, sheetName: string): void;
  onCreateSheet?: (() => void) | undefined;
  onRenameSheet?: ((currentName: string, nextName: string) => void) | undefined;
}

const SHEET_STRIP_CLASS =
  "flex min-h-11 items-center justify-between gap-3 border-t border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2.5 py-1.5";
const SHEET_TABS_ROOT_CLASS = "min-w-0";
const SHEET_LIST_CLASS =
  "relative flex max-w-full items-center gap-1 overflow-x-auto rounded-[calc(var(--wb-radius-control)+4px)] border border-[var(--wb-border)] bg-[var(--wb-surface-muted)] p-1";
const SHEET_INDICATOR_CLASS =
  "pointer-events-none absolute inset-y-1 rounded-[var(--wb-radius-control)] border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)] transition-[left,top,width,height] duration-150 ease-out";
const SHEET_RENAME_SHELL_CLASS =
  "relative z-[1] inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent)] bg-[var(--wb-surface)] px-3 shadow-[var(--wb-shadow-sm)]";
const SHEET_ACTION_BUTTON_CLASS =
  "inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-subtle)] outline-none transition-[background-color,border-color,color,box-shadow] hover:border-[var(--wb-border-strong)] hover:bg-[var(--wb-surface)] hover:text-[var(--wb-text)] hover:shadow-[var(--wb-shadow-sm)] focus-visible:border-[var(--wb-accent)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] disabled:cursor-not-allowed disabled:opacity-50";

function SheetAddIcon() {
  return (
    <svg aria-hidden="true" className="h-4 w-4" fill="none" viewBox="0 0 16 16">
      <path
        d="M8 3.25v9.5M3.25 8h9.5"
        stroke="currentColor"
        strokeLinecap="round"
        strokeWidth="1.5"
      />
    </svg>
  );
}

function getSheetTabClassName(state: Tabs.Tab.State): string {
  return cn(
    "relative z-[1] inline-flex h-8 items-center rounded-[var(--wb-radius-control)] px-3 text-[12px] outline-none transition-[color,font-weight] duration-150",
    state.active
      ? "font-semibold text-[var(--wb-text)]"
      : "font-medium text-[var(--wb-text-subtle)] hover:text-[var(--wb-text)]",
    !state.disabled &&
      "focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:text-[var(--wb-text)]",
    state.disabled && "cursor-not-allowed opacity-50",
  );
}

export const WorkbookSheetTabs = React.memo(function WorkbookSheetTabs({
  sheetName,
  sheetNames,
  selectionStatus,
  onSelectSheet,
  onCreateSheet,
  onRenameSheet,
}: WorkbookSheetTabsProps) {
  const [renamingSheetName, setRenamingSheetName] = React.useState<string | null>(null);
  const [renameDraft, setRenameDraft] = React.useState("");
  const renameInputRef = React.useRef<HTMLInputElement | null>(null);

  const startSheetRename = React.useCallback(
    (targetSheetName: string) => {
      if (!onRenameSheet) {
        return;
      }
      onSelectSheet(targetSheetName);
      setRenamingSheetName(targetSheetName);
      setRenameDraft(targetSheetName);
    },
    [onRenameSheet, onSelectSheet],
  );

  const cancelSheetRename = React.useCallback(() => {
    setRenamingSheetName(null);
    setRenameDraft("");
  }, []);

  const commitSheetRename = React.useCallback(
    (targetSheetName: string) => {
      const nextName = renameDraft.trim();
      setRenamingSheetName(null);
      setRenameDraft("");
      if (!onRenameSheet || nextName.length === 0 || nextName === targetSheetName) {
        return;
      }
      onRenameSheet(targetSheetName, nextName);
    },
    [onRenameSheet, renameDraft],
  );

  React.useEffect(() => {
    if (!renamingSheetName) {
      return;
    }
    renameInputRef.current?.focus();
    renameInputRef.current?.select();
  }, [renamingSheetName]);

  React.useEffect(() => {
    if (renamingSheetName && !sheetNames.includes(renamingSheetName)) {
      cancelSheetRename();
    }
  }, [cancelSheetRename, renamingSheetName, sheetNames]);

  return (
    <div className={SHEET_STRIP_CLASS}>
      <div className="flex min-w-0 items-center gap-1.5">
        <Tabs.Root
          className={SHEET_TABS_ROOT_CLASS}
          value={sheetName}
          onValueChange={(value) => onSelectSheet(String(value))}
        >
          <Tabs.List aria-label="Sheets" className={SHEET_LIST_CLASS}>
            <Tabs.Indicator className={SHEET_INDICATOR_CLASS} renderBeforeHydration />
            {sheetNames.map((name) =>
              renamingSheetName === name ? (
                <div className={SHEET_RENAME_SHELL_CLASS} key={name}>
                  <input
                    aria-label={`Rename ${name}`}
                    className="w-[120px] min-w-0 border-none bg-transparent p-0 text-[12px] font-semibold text-[var(--wb-text)] outline-none"
                    onBlur={() => commitSheetRename(name)}
                    onChange={(event) => setRenameDraft(event.target.value)}
                    onClick={(event) => event.stopPropagation()}
                    onKeyDown={(event) => {
                      if (event.key === "Enter") {
                        event.preventDefault();
                        commitSheetRename(name);
                        return;
                      }
                      if (event.key === "Escape") {
                        event.preventDefault();
                        cancelSheetRename();
                      }
                    }}
                    ref={renameInputRef}
                    value={renameDraft}
                  />
                </div>
              ) : (
                <Tabs.Tab
                  className={getSheetTabClassName}
                  key={name}
                  onDoubleClick={() => startSheetRename(name)}
                  onKeyDown={(event) => {
                    if (event.key === "F2") {
                      event.preventDefault();
                      startSheetRename(name);
                    }
                  }}
                  title={name}
                  value={name}
                >
                  {name}
                </Tabs.Tab>
              ),
            )}
          </Tabs.List>
        </Tabs.Root>
        {onCreateSheet ? (
          <Button
            aria-label="Create sheet"
            className={SHEET_ACTION_BUTTON_CLASS}
            onClick={onCreateSheet}
            title="Add sheet"
            type="button"
          >
            <SheetAddIcon />
          </Button>
        ) : null}
      </div>
      {selectionStatus ? (
        <div className="inline-flex flex-wrap items-center gap-1.5">{selectionStatus}</div>
      ) : null}
    </div>
  );
});
