import { Button } from '@base-ui/react/button'
import { Dialog } from '@base-ui/react/dialog'
import { Menu } from '@base-ui/react/menu'
import { ChevronDown, FileSpreadsheet } from 'lucide-react'
import { useEffect, useState } from 'react'
import { BILIG_CONTENT_TYPE } from '@bilig/agent-api'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  createBlankWorkbook,
  duplicateWorkbook,
  loadRecentWorkbooks,
  rememberWorkbook,
  workbookBackupFile,
  type RecentWorkbook,
} from './workbook-documents.js'
import { finalizeWorkbookImport, resolveImportedWorkbookNavigationUrl } from './workbook-import-client.js'
import { resolveWorkbookNavigationUrl } from './workbook-navigation.js'

const actionClass =
  'flex h-8 items-center gap-2 rounded-[var(--wb-radius-control)] px-2 text-[12px] text-[var(--wb-text)] hover:bg-[var(--wb-hover)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] disabled:opacity-50'
const menuItemClass = `${actionClass} w-full cursor-default data-[highlighted]:bg-[var(--wb-hover)] data-[disabled]:opacity-50`

export function WorkbookDocumentHeader(props: {
  readonly documentId: string
  readonly workbookName: string
  readonly userId: string
  readonly serverUrl?: string
  readonly isSynced: boolean
  readonly isReady: boolean
  readonly exportSnapshot: () => Promise<WorkbookSnapshot>
  readonly flushPendingEdit: () => Promise<void>
  readonly onRename: (name: string) => Promise<void>
  readonly onImport: () => void
  readonly onError: (error: unknown) => void
}) {
  const [isBusy, setIsBusy] = useState(false)
  const [isRenaming, setIsRenaming] = useState(false)
  const [nameDraft, setNameDraft] = useState(props.workbookName)
  const [recent, setRecent] = useState<RecentWorkbook[]>([])
  const disabled = !props.isReady || isBusy

  useEffect(() => {
    if (!props.isReady) return
    document.title = `${props.workbookName} · Bilig`
    rememberWorkbook(props.userId, {
      documentId: props.documentId,
      name: props.workbookName,
      ...(props.serverUrl ? { serverUrl: props.serverUrl } : {}),
    })
    setRecent(loadRecentWorkbooks(props.userId))
  }, [props.documentId, props.isReady, props.serverUrl, props.userId, props.workbookName])

  async function run(action: () => Promise<void>): Promise<void> {
    if (disabled) return
    setIsBusy(true)
    try {
      await action()
    } catch (error) {
      props.onError(error)
    } finally {
      setIsBusy(false)
    }
  }

  async function create(snapshot: WorkbookSnapshot): Promise<void> {
    const result = await finalizeWorkbookImport({ file: workbookBackupFile(snapshot), contentType: BILIG_CONTENT_TYPE, openMode: 'create' })
    window.location.assign(resolveImportedWorkbookNavigationUrl(result))
  }

  return (
    <header
      className="flex h-10 shrink-0 items-center gap-2 border-b border-[var(--wb-border)] bg-[var(--wb-surface)] px-3"
      aria-label="Workbook document"
    >
      <FileSpreadsheet aria-hidden="true" className="size-4 shrink-0 text-[var(--wb-text-muted)]" />
      <Button
        className={`${actionClass} min-w-0 max-w-[min(40vw,24rem)] font-semibold`}
        disabled={disabled}
        title="Rename workbook"
        aria-label={`Rename workbook: ${props.workbookName}`}
        onClick={() => {
          setNameDraft(props.workbookName)
          setIsRenaming(true)
        }}
      >
        <span className="truncate" data-testid="workbook-title">
          {props.workbookName}
        </span>
      </Button>
      <Menu.Root>
        <Menu.Trigger className={actionClass} disabled={disabled} aria-label="Workbook menu">
          File <ChevronDown aria-hidden="true" className="size-3" />
        </Menu.Trigger>
        <Menu.Portal>
          <Menu.Positioner sideOffset={4} className="z-[1200]">
            <Menu.Popup className="w-64 rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-1 shadow-[var(--wb-shadow-md)]">
              <Menu.Item
                className={menuItemClass}
                onClick={() => {
                  void run(async () => {
                    await props.flushPendingEdit()
                    await create(createBlankWorkbook())
                  })
                }}
              >
                New workbook
              </Menu.Item>
              <Menu.Item
                className={menuItemClass}
                onClick={() => {
                  setNameDraft(props.workbookName)
                  setIsRenaming(true)
                }}
              >
                Rename workbook
              </Menu.Item>
              <Menu.Item
                className={menuItemClass}
                onClick={() => {
                  void run(async () => {
                    await create(duplicateWorkbook(await props.exportSnapshot()))
                  })
                }}
              >
                Duplicate workbook
              </Menu.Item>
              <Menu.Item className={menuItemClass} onClick={props.onImport}>
                Import workbook…
              </Menu.Item>
              <Menu.Item
                className={menuItemClass}
                onClick={() => {
                  void run(async () => {
                    const file = workbookBackupFile(await props.exportSnapshot())
                    const url = URL.createObjectURL(file)
                    const anchor = document.createElement('a')
                    anchor.href = url
                    anchor.download = file.name
                    anchor.click()
                    window.setTimeout(() => URL.revokeObjectURL(url), 1000)
                  })
                }}
              >
                Export Bilig backup
              </Menu.Item>
              <Menu.Separator className="my-1 h-px bg-[var(--wb-border)]" />
              <Menu.Group>
                <Menu.GroupLabel className="px-2 py-1 text-[11px] text-[var(--wb-text-muted)]">
                  Recent workbooks on this browser
                </Menu.GroupLabel>
                {recent
                  .filter((entry) => entry.documentId !== props.documentId)
                  .map((entry) => (
                    <Menu.Item
                      key={entry.documentId}
                      className={`${menuItemClass} truncate`}
                      onClick={() => {
                        void run(async () => {
                          await props.flushPendingEdit()
                          window.location.assign(resolveWorkbookNavigationUrl(entry))
                        })
                      }}
                    >
                      {entry.name}
                    </Menu.Item>
                  ))}
                {recent.every((entry) => entry.documentId === props.documentId) ? (
                  <div className="px-2 py-1 text-[12px] text-[var(--wb-text-subtle)]">No other recent workbooks</div>
                ) : null}
              </Menu.Group>
            </Menu.Popup>
          </Menu.Positioner>
        </Menu.Portal>
      </Menu.Root>
      <span className="ml-auto truncate text-[11px] text-[var(--wb-text-muted)]" title={`Document: ${props.documentId}`}>
        {isBusy ? 'Working…' : props.isSynced ? 'Synced workbook' : 'Local workbook'}
      </span>
      <Dialog.Root open={isRenaming} onOpenChange={setIsRenaming}>
        <Dialog.Portal>
          <Dialog.Backdrop className="fixed inset-0 z-[1200] bg-black/35" />
          <Dialog.Popup className="fixed left-1/2 top-1/2 z-[1201] w-[min(24rem,calc(100vw-2rem))] -translate-x-1/2 -translate-y-1/2 rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-5 shadow-[var(--wb-shadow-md)]">
            <Dialog.Title className="text-[14px] font-semibold text-[var(--wb-text)]">Rename workbook</Dialog.Title>
            <Dialog.Description className="mt-1 text-[12px] text-[var(--wb-text-muted)]">
              Choose a name for this workbook.
            </Dialog.Description>
            <form
              onSubmit={(event) => {
                event.preventDefault()
                void run(async () => {
                  await props.onRename(nameDraft.trim())
                  setIsRenaming(false)
                })
              }}
            >
              <label className="mt-4 block text-[12px] text-[var(--wb-text)]" htmlFor="workbook-name">
                Workbook name
              </label>
              <input
                id="workbook-name"
                className="mt-1 h-8 w-full rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-2 text-[13px] text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
                value={nameDraft}
                maxLength={200}
                onChange={(event) => setNameDraft(event.target.value)}
              />
              <div className="mt-4 flex justify-end gap-2">
                <Dialog.Close className={actionClass}>Cancel</Dialog.Close>
                <Button
                  className={`${actionClass} border border-[var(--wb-border)]`}
                  type="submit"
                  disabled={disabled || !nameDraft.trim()}
                >
                  Save name
                </Button>
              </div>
            </form>
          </Dialog.Popup>
        </Dialog.Portal>
      </Dialog.Root>
    </header>
  )
}
