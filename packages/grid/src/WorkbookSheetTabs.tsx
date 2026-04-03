import React from "react";
import { Tabs } from "@base-ui/react/tabs";

interface WorkbookSheetTabsProps {
  sheetName: string;
  sheetNames: string[];
  statusBar?: React.ReactNode;
  onSelectSheet(this: void, sheetName: string): void;
  onCreateSheet?: (() => void) | undefined;
  onRenameSheet?: ((currentName: string, nextName: string) => void) | undefined;
}

const TAB_CLASS_NAME =
  "inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-transparent bg-transparent px-3 text-[12px] font-medium text-[var(--wb-text-muted)] outline-none transition-[background-color,border-color,color,box-shadow] hover:border-[var(--wb-border)] hover:bg-[var(--wb-surface)] hover:text-[var(--wb-text)] focus-visible:border-[var(--wb-accent)] focus-visible:bg-[var(--wb-surface)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] data-[active]:border-[var(--wb-border)] data-[active]:bg-[var(--wb-surface)] data-[active]:text-[var(--wb-text)] data-[active]:shadow-[var(--wb-shadow-sm)]";

export const WorkbookSheetTabs = React.memo(function WorkbookSheetTabs({
  sheetName,
  sheetNames,
  statusBar,
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
    <div className="flex min-h-11 items-center justify-between gap-3 border-t border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2.5 py-1.5">
      <div className="flex min-w-0 items-center gap-1.5">
        <Tabs.Root value={sheetName} onValueChange={(value) => onSelectSheet(String(value))}>
          <Tabs.List
            aria-label="Sheets"
            className="flex max-w-full items-center gap-1 overflow-x-auto pr-1"
          >
            {sheetNames.map((name) =>
              renamingSheetName === name ? (
                <div
                  className={`${TAB_CLASS_NAME} ${sheetName === name ? "border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)]" : ""}`}
                  key={name}
                >
                  <input
                    aria-label={`Rename ${name}`}
                    className="w-[120px] min-w-0 border-none bg-transparent p-0 text-[12px] font-medium text-[var(--wb-text)] outline-none"
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
                  className={TAB_CLASS_NAME}
                  key={name}
                  onDoubleClick={() => startSheetRename(name)}
                  onKeyDown={(event) => {
                    if (event.key === "F2") {
                      event.preventDefault();
                      startSheetRename(name);
                    }
                  }}
                  value={name}
                >
                  {name}
                </Tabs.Tab>
              ),
            )}
          </Tabs.List>
        </Tabs.Root>
        {onCreateSheet ? (
          <button
            aria-label="Create sheet"
            className="inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] text-[18px] leading-none text-[var(--wb-text-muted)] outline-none transition-[background-color,border-color,color,box-shadow] hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)] hover:shadow-[var(--wb-shadow-sm)] focus-visible:border-[var(--wb-accent)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
            onClick={onCreateSheet}
            title="Add sheet"
            type="button"
          >
            +
          </button>
        ) : null}
      </div>
      {statusBar ? (
        <div className="inline-flex flex-wrap items-center gap-1.5">{statusBar}</div>
      ) : null}
    </div>
  );
});
