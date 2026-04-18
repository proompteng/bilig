import React from 'react'
import { ContextMenu } from '@base-ui/react/context-menu'
import { Tabs } from '@base-ui/react/tabs'
import { cva } from 'class-variance-authority'

interface WorkbookSheetTabsProps {
  sheetName: string
  sheetNames: string[]
  trailingContent?: React.ReactNode
  onSelectSheet(this: void, sheetName: string): void
  onCreateSheet?: (() => void) | undefined
  onRenameSheet?: ((currentName: string, nextName: string) => void) | undefined
  onDeleteSheet?: ((sheetName: string) => void) | undefined
}

const sheetStripClass = cva(
  'flex min-h-12 items-center justify-between gap-3 border-t border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2.5 pt-1.5 pb-2',
)

const sheetTabsShellClass = cva('flex min-w-0 flex-1 items-end gap-2 overflow-hidden')

const sheetListClass = cva('wb-scrollbar-none relative flex min-w-0 items-center gap-1 overflow-x-auto overflow-y-hidden')

const sheetTabClass = cva(
  'inline-flex h-8 shrink-0 items-center justify-center rounded-[var(--wb-radius-control)] border px-3 text-[12px] font-medium whitespace-nowrap outline-none transition-[color,background-color,border-color,box-shadow] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface-subtle)]',
  {
    variants: {
      active: {
        true: 'border-[var(--wb-border-strong)] bg-[var(--wb-surface)] font-semibold text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)]',
        false:
          'border-transparent bg-transparent text-[var(--wb-text-muted)] hover:border-[var(--wb-border)] hover:bg-[var(--wb-muted)] hover:text-[var(--wb-text)]',
      },
    },
    defaultVariants: {
      active: false,
    },
  },
)

const sheetIndicatorClass = cva(
  'absolute bottom-0 left-0 h-0.5 w-[var(--active-tab-width)] translate-x-[var(--active-tab-left)] rounded-full bg-[var(--wb-accent)] transition-[translate,width] duration-200 ease-out',
)

const sheetRenameShellClass = cva(
  'inline-flex h-8 shrink-0 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent)] bg-[var(--wb-surface)] px-2',
)

const sheetRenameInputClass = cva(
  'w-[120px] min-w-0 border-none bg-transparent p-0 text-[12px] font-medium text-[var(--wb-text)] outline-none',
)

const sheetActionButtonClass = cva(
  'inline-flex size-8 shrink-0 items-center justify-center rounded-[var(--wb-radius-control)] border border-transparent bg-transparent text-[var(--wb-text-muted)] outline-none transition-colors hover:border-[var(--wb-border)] hover:bg-[var(--wb-muted)] hover:text-[var(--wb-text)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface-subtle)] disabled:cursor-not-allowed disabled:opacity-50',
)

const sheetContextMenuPopupClass = cva(
  'min-w-40 overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-1 shadow-[var(--wb-shadow-md)] outline-none',
)

const sheetContextMenuItemClass = cva(
  'flex h-8 items-center rounded-[var(--wb-radius-control)] px-2.5 text-[12px] font-medium text-[var(--wb-text)] outline-none transition-colors data-[highlighted]:bg-[var(--wb-muted)] data-[disabled]:pointer-events-none data-[disabled]:opacity-45',
)

const sheetContextMenuSeparatorClass = cva('my-1 h-px bg-[var(--wb-border)]')

const sheetTrailingContentClass = cva('shrink-0 text-[11px] font-medium text-[var(--wb-text-muted)]')

function SheetAddIcon() {
  return (
    <svg aria-hidden="true" className="h-4 w-4" fill="none" viewBox="0 0 16 16">
      <path d="M8 3.25v9.5M3.25 8h9.5" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    </svg>
  )
}

export const WorkbookSheetTabs = React.memo(function WorkbookSheetTabs({
  sheetName,
  sheetNames,
  trailingContent,
  onSelectSheet,
  onCreateSheet,
  onRenameSheet,
  onDeleteSheet,
}: WorkbookSheetTabsProps) {
  const [renamingSheetName, setRenamingSheetName] = React.useState<string | null>(null)
  const [renameDraft, setRenameDraft] = React.useState('')
  const renameInputRef = React.useRef<HTMLInputElement | null>(null)
  const tabRefs = React.useRef<Record<string, HTMLElement | null>>({})

  const startSheetRename = React.useCallback(
    (targetSheetName: string) => {
      if (!onRenameSheet) {
        return
      }
      onSelectSheet(targetSheetName)
      setRenamingSheetName(targetSheetName)
      setRenameDraft(targetSheetName)
    },
    [onRenameSheet, onSelectSheet],
  )

  const cancelSheetRename = React.useCallback(() => {
    setRenamingSheetName(null)
    setRenameDraft('')
  }, [])

  const commitSheetRename = React.useCallback(
    (targetSheetName: string) => {
      const nextName = renameDraft.trim()
      setRenamingSheetName(null)
      setRenameDraft('')
      if (!onRenameSheet || nextName.length === 0 || nextName === targetSheetName) {
        return
      }
      onRenameSheet(targetSheetName, nextName)
    },
    [onRenameSheet, renameDraft],
  )

  React.useEffect(() => {
    if (!renamingSheetName) {
      return
    }
    renameInputRef.current?.focus()
    renameInputRef.current?.select()
  }, [renamingSheetName])

  React.useEffect(() => {
    if (renamingSheetName && !sheetNames.includes(renamingSheetName)) {
      cancelSheetRename()
    }
  }, [cancelSheetRename, renamingSheetName, sheetNames])

  const focusSheetTab = React.useCallback((targetSheetName: string) => {
    tabRefs.current[targetSheetName]?.focus()
  }, [])

  const moveSheetSelection = React.useCallback(
    (targetIndex: number) => {
      const resolvedIndex = Math.max(0, Math.min(sheetNames.length - 1, targetIndex))
      const targetSheetName = sheetNames[resolvedIndex]
      if (!targetSheetName) {
        return
      }
      onSelectSheet(targetSheetName)
      queueMicrotask(() => {
        focusSheetTab(targetSheetName)
      })
    },
    [focusSheetTab, onSelectSheet, sheetNames],
  )

  const deleteDisabled = sheetNames.length <= 1

  return (
    <Tabs.Root
      className={sheetStripClass()}
      value={sheetName}
      onValueChange={(nextValue) => {
        onSelectSheet(String(nextValue))
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
                    if (event.key === 'Enter') {
                      event.preventDefault()
                      commitSheetRename(name)
                      return
                    }
                    if (event.key === 'Escape') {
                      event.preventDefault()
                      cancelSheetRename()
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
                      tabRefs.current[name] = node
                    }}
                    title={name}
                    value={name}
                    onDoubleClick={() => startSheetRename(name)}
                    onKeyDown={(event) => {
                      if (event.key === 'F2') {
                        event.preventDefault()
                        startSheetRename(name)
                        return
                      }
                      if (event.key === 'ArrowRight') {
                        event.preventDefault()
                        moveSheetSelection(index + 1)
                        return
                      }
                      if (event.key === 'ArrowLeft') {
                        event.preventDefault()
                        moveSheetSelection(index - 1)
                        return
                      }
                      if (event.key === 'Home') {
                        event.preventDefault()
                        moveSheetSelection(0)
                        return
                      }
                      if (event.key === 'End') {
                        event.preventDefault()
                        moveSheetSelection(sheetNames.length - 1)
                      }
                    }}
                  >
                    {name}
                  </Tabs.Tab>
                )

                if (!onRenameSheet && !onDeleteSheet) {
                  return tab
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
                                startSheetRename(name)
                              }}
                            >
                              Rename sheet
                            </ContextMenu.Item>
                          ) : null}
                          {onRenameSheet && onDeleteSheet ? <ContextMenu.Separator className={sheetContextMenuSeparatorClass()} /> : null}
                          {onDeleteSheet ? (
                            <ContextMenu.Item
                              className={sheetContextMenuItemClass()}
                              data-testid="workbook-sheet-menu-delete"
                              disabled={deleteDisabled}
                              onClick={() => {
                                onDeleteSheet(name)
                              }}
                            >
                              Delete sheet
                            </ContextMenu.Item>
                          ) : null}
                        </ContextMenu.Popup>
                      </ContextMenu.Positioner>
                    </ContextMenu.Portal>
                  </ContextMenu.Root>
                )
              })()
            ),
          )}
          <Tabs.Indicator className={sheetIndicatorClass()} data-testid="workbook-sheet-tab-indicator" renderBeforeHydration />
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
      {trailingContent ? <div className={sheetTrailingContentClass()}>{trailingContent}</div> : null}
    </Tabs.Root>
  )
})
