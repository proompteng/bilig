import { Button } from '@base-ui/react/button'
import type { WorkbookLoadedResponse } from '@bilig/agent-api'
import { Upload } from 'lucide-react'
import { useMemo } from 'react'
import { WorkbookImportPanel } from './WorkbookImportPanel.js'
import { cn } from './cn.js'
import { workbookHeaderActionButtonClass } from './workbook-header-controls.js'
import { finalizeWorkbookImport, previewWorkbookImport, resolveImportedWorkbookNavigationUrl } from './workbook-import-client.js'
import { useWorkbookImportController } from './use-workbook-import-controller.js'

export function useWorkbookImportPane(input: {
  readonly currentDocumentId: string
  readonly enabled: boolean
  readonly previewFile?: typeof previewWorkbookImport
  readonly finalizeImport?: typeof finalizeWorkbookImport
  readonly navigateToWorkbook?: (result: WorkbookLoadedResponse) => void
}) {
  const {
    currentDocumentId,
    enabled,
    previewFile = previewWorkbookImport,
    finalizeImport = finalizeWorkbookImport,
    navigateToWorkbook = (result: WorkbookLoadedResponse) => {
      window.location.assign(resolveImportedWorkbookNavigationUrl(result))
    },
  } = input
  const controller = useWorkbookImportController({ currentDocumentId, enabled, previewFile, finalizeImport, navigateToWorkbook })

  const importToggle = useMemo(
    () => (
      <Button
        aria-controls="workbook-import-panel"
        aria-expanded={controller.isOpen}
        aria-label="Import workbook"
        className={cn(
          workbookHeaderActionButtonClass({ active: controller.isOpen, iconOnly: true }),
          controller.isOpen
            ? 'border-transparent bg-[var(--color-mauve-100)] text-[var(--color-mauve-900)] shadow-none'
            : 'border-transparent bg-transparent text-[var(--color-mauve-700)] shadow-none hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]',
          'max-[420px]:hidden',
        )}
        data-testid="workbook-import-toggle"
        disabled={!enabled}
        title="Import workbook"
        type="button"
        onClick={() => {
          controller.toggle()
        }}
      >
        <Upload aria-hidden="true" className="size-4" strokeWidth={1.9} />
      </Button>
    ),
    [controller, enabled],
  )

  const importPanel = useMemo(
    () => (
      <WorkbookImportPanel
        enabled={enabled}
        isImporting={controller.isImporting}
        isOpen={controller.isOpen}
        isPreviewing={controller.isPreviewing}
        stagedPreview={controller.stagedPreview}
        onClose={controller.close}
        onFileSelected={(file) => {
          void controller.stageFile(file)
        }}
        onImportAsNew={() => {
          void controller.importStagedFile('create')
        }}
        onReplaceCurrent={() => {
          void controller.importStagedFile('replace')
        }}
      />
    ),
    [controller, enabled],
  )

  return {
    clearImportError: controller.clearError,
    importError: controller.error,
    importPanel,
    importToggle,
  }
}
