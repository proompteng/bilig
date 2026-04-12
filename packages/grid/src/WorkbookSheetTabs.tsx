import React from "react";
import { ContextMenu } from "@base-ui/react/context-menu";
import { Tabs } from "@base-ui/react/tabs";
import { cva } from "class-variance-authority";

interface WorkbookSheetTabsProps {
  sheetName: string;
  sheetNames: string[];
  onSelectSheet(this: void, sheetName: string): void;
  onCreateSheet?: (() => void) | undefined;
  onRenameSheet?: ((currentName: string, nextName: string) => void) | undefined;
  onDeleteSheet?: ((sheetName: string) => void) | undefined;
}

const sheetStripClass = cva(
  "flex min-h-11 items-center justify-between gap-2 border-t border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] px-2.5 py-1.5",
);

const sheetTabsShellClass = cva("flex min-w-0 flex-1 items-center gap-2 overflow-hidden");

const sheetListClass = cva(
  "relative flex min-w-0 items-center gap-0.5 overflow-x-auto overflow-y-hidden border-b border-[var(--color-mauve-200)]",
);

const sheetIndicatorClass = cva(
  "absolute bottom-[-1px] left-0 z-10 h-0.5 w-[var(--active-tab-width)] translate-x-[var(--active-tab-left)] rounded-full bg-[var(--color-mauve-800)] transition-[translate,width] duration-200 ease-out",
);

const sheetTabClass = cva(
  "inline-flex h-8 shrink-0 items-center justify-center border-b-2 border-transparent px-3 text-[12px] font-medium whitespace-nowrap text-[var(--color-mauve-600)] outline-none transition-[color] hover:text-[var(--color-mauve-900)] focus-visible:rounded-md focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)]",
  {
    variants: {
      active: {
        true: "text-[var(--color-mauve-950)]",
        false: "border-transparent",
      },
    },
    defaultVariants: {
      active: false,
    },
  },
);

const sheetRenameShellClass = cva(
  "inline-flex h-8 shrink-0 items-center border-b-2 border-[var(--color-mauve-700)] px-2",
);

const sheetRenameInputClass = cva(
  "w-[120px] min-w-0 border-none bg-transparent p-0 text-[12px] font-medium text-[var(--color-mauve-950)] outline-none",
);

const sheetActionButtonClass = cva(
  "inline-flex size-8 shrink-0 items-center justify-center rounded-md border-0 bg-transparent text-[var(--color-mauve-700)] outline-none transition-colors hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)] focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)] disabled:cursor-not-allowed disabled:opacity-50",
);

const sheetContextMenuPopupClass = cva(
  "min-w-40 overflow-hidden rounded-lg border border-[var(--color-mauve-200)] bg-white p-1 shadow-[0_12px_28px_rgba(15,23,42,0.12)] outline-none",
);

const sheetContextMenuItemClass = cva(
  "flex h-8 items-center rounded-md px-2.5 text-[12px] font-medium text-[var(--color-mauve-900)] outline-none transition-colors data-[highlighted]:bg-[var(--color-mauve-100)] data-[disabled]:pointer-events-none data-[disabled]:opacity-45",
);

const sheetContextMenuSeparatorClass = cva("my-1 h-px bg-[var(--color-mauve-200)]");

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

export const WorkbookSheetTabs = React.memo(function WorkbookSheetTabs({
  sheetName,
  sheetNames,
  onSelectSheet,
  onCreateSheet,
  onRenameSheet,
  onDeleteSheet,
}: WorkbookSheetTabsProps) {
  const [renamingSheetName, setRenamingSheetName] = React.useState<string | null>(null);
  const [renameDraft, setRenameDraft] = React.useState("");
  const renameInputRef = React.useRef<HTMLInputElement | null>(null);
  const tabRefs = React.useRef<Record<string, HTMLElement | null>>({});

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

  const focusSheetTab = React.useCallback((targetSheetName: string) => {
    tabRefs.current[targetSheetName]?.focus();
  }, []);

  const moveSheetSelection = React.useCallback(
    (targetIndex: number) => {
      const resolvedIndex = Math.max(0, Math.min(sheetNames.length - 1, targetIndex));
      const targetSheetName = sheetNames[resolvedIndex];
      if (!targetSheetName) {
        return;
      }
      onSelectSheet(targetSheetName);
      queueMicrotask(() => {
        focusSheetTab(targetSheetName);
      });
    },
    [focusSheetTab, onSelectSheet, sheetNames],
  );

  const deleteDisabled = sheetNames.length <= 1;

  return (
    <Tabs.Root
      className={sheetStripClass()}
      value={sheetName}
      onValueChange={(nextValue) => {
        onSelectSheet(String(nextValue));
      }}
    >
      <div className={sheetTabsShellClass()}>
        <Tabs.List aria-label="Sheets" className={sheetListClass()}>
          {sheetNames.map((name, index) =>
            renamingSheetName === name ? (
              <div className={sheetRenameShellClass()} key={name}>
                <input
                  aria-label={`Rename ${name}`}
                  className={sheetRenameInputClass()}
                  data-testid="workbook-sheet-rename-input"
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
              (() => {
                const tab = (
                  <Tabs.Tab
                    className={(state) => sheetTabClass({ active: state.active })}
                    data-testid={`workbook-sheet-tab-${name}`}
                    key={name}
                    ref={(node) => {
                      tabRefs.current[name] = node;
                    }}
                    title={name}
                    value={name}
                    onDoubleClick={() => startSheetRename(name)}
                    onKeyDown={(event) => {
                      if (event.key === "F2") {
                        event.preventDefault();
                        startSheetRename(name);
                        return;
                      }
                      if (event.key === "ArrowRight") {
                        event.preventDefault();
                        moveSheetSelection(index + 1);
                        return;
                      }
                      if (event.key === "ArrowLeft") {
                        event.preventDefault();
                        moveSheetSelection(index - 1);
                        return;
                      }
                      if (event.key === "Home") {
                        event.preventDefault();
                        moveSheetSelection(0);
                        return;
                      }
                      if (event.key === "End") {
                        event.preventDefault();
                        moveSheetSelection(sheetNames.length - 1);
                      }
                    }}
                  >
                    {name}
                  </Tabs.Tab>
                );

                if (!onRenameSheet && !onDeleteSheet) {
                  return tab;
                }

                return (
                  <ContextMenu.Root key={name}>
                    <ContextMenu.Trigger className="shrink-0">{tab}</ContextMenu.Trigger>
                    <ContextMenu.Portal>
                      <ContextMenu.Positioner className="z-[1200]" sideOffset={6}>
                        <ContextMenu.Popup
                          aria-label={`${name} sheet actions`}
                          className={sheetContextMenuPopupClass()}
                          data-testid={`workbook-sheet-menu-${name}`}
                        >
                          {onRenameSheet ? (
                            <ContextMenu.Item
                              className={sheetContextMenuItemClass()}
                              data-testid="workbook-sheet-menu-rename"
                              onClick={() => {
                                startSheetRename(name);
                              }}
                            >
                              Rename sheet
                            </ContextMenu.Item>
                          ) : null}
                          {onRenameSheet && onDeleteSheet ? (
                            <ContextMenu.Separator className={sheetContextMenuSeparatorClass()} />
                          ) : null}
                          {onDeleteSheet ? (
                            <ContextMenu.Item
                              className={sheetContextMenuItemClass()}
                              data-testid="workbook-sheet-menu-delete"
                              disabled={deleteDisabled}
                              onClick={() => {
                                onDeleteSheet(name);
                              }}
                            >
                              Delete sheet
                            </ContextMenu.Item>
                          ) : null}
                        </ContextMenu.Popup>
                      </ContextMenu.Positioner>
                    </ContextMenu.Portal>
                  </ContextMenu.Root>
                );
              })()
            ),
          )}
          <Tabs.Indicator className={sheetIndicatorClass()} renderBeforeHydration />
        </Tabs.List>
        {onCreateSheet ? (
          <button
            aria-label="Create sheet"
            className={sheetActionButtonClass()}
            data-testid="workbook-sheet-add"
            onClick={onCreateSheet}
            title="Add sheet"
            type="button"
          >
            <SheetAddIcon />
          </button>
        ) : null}
      </div>
    </Tabs.Root>
  );
});
