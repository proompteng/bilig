// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookLoadedResponse } from '@bilig/agent-api'
import { CSV_CONTENT_TYPE, XLSX_CONTENT_TYPE } from '@bilig/agent-api'
import type { ImportedWorkbookPreview } from '@bilig/excel-import'
import type { ToastT, ToastToDismiss } from 'sonner'
import { toast } from 'sonner'
import { WorkbookToastRegion } from '../WorkbookToastRegion.js'
import { resolveWorkbookImportContentType } from '../workbook-import-client.js'
import { useWorkbookImportPane } from '../use-workbook-import-pane.js'

async function flushToasts(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

function findActiveToast(id: string): ToastT | null {
  return toast.getToasts().find((entry: ToastT | ToastToDismiss): entry is ToastT => !('dismiss' in entry) && entry.id === id) ?? null
}

function setInputFiles(input: HTMLInputElement, files: readonly File[]): void {
  Object.defineProperty(input, 'files', {
    configurable: true,
    value: {
      0: files[0],
      length: files.length,
      item: (index: number) => files[index] ?? null,
    },
  })
}

function createPreview(overrides: Partial<ImportedWorkbookPreview> = {}): ImportedWorkbookPreview {
  return {
    fileName: 'metrics.csv',
    contentType: CSV_CONTENT_TYPE,
    fileSizeBytes: 24,
    workbookName: 'metrics',
    sheetCount: 1,
    sheets: [
      {
        name: 'metrics',
        rowCount: 2,
        columnCount: 2,
        nonEmptyCellCount: 4,
        previewRows: [
          ['Name', 'Value'],
          ['alpha', '12'],
        ],
      },
    ],
    warnings: [],
    ...overrides,
  }
}

function ImportHarness(props: {
  readonly previewFile?: Parameters<typeof useWorkbookImportPane>[0]['previewFile']
  readonly finalizeImport?: Parameters<typeof useWorkbookImportPane>[0]['finalizeImport']
  readonly navigateToWorkbook?: (result: WorkbookLoadedResponse) => void
}) {
  const { clearImportError, importError, importPanel, importToggle } = useWorkbookImportPane({
    currentDocumentId: 'doc-1',
    enabled: true,
    previewFile: props.previewFile,
    finalizeImport: props.finalizeImport,
    navigateToWorkbook: props.navigateToWorkbook,
  })

  return (
    <div>
      <WorkbookToastRegion
        toasts={
          importError
            ? [
                {
                  id: 'import-error',
                  tone: 'error',
                  message: importError,
                  onDismiss: clearImportError,
                },
              ]
            : []
        }
      />
      {importToggle}
      {importPanel}
    </div>
  )
}

afterEach(() => {
  toast.dismiss()
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

describe('workbook import', () => {
  it('classifies supported csv and xlsx workbook files', () => {
    expect(resolveWorkbookImportContentType(new File(['alpha'], 'metrics.csv', { type: CSV_CONTENT_TYPE }))).toBe(CSV_CONTENT_TYPE)
    expect(resolveWorkbookImportContentType(new File(['alpha'], 'model.xlsx', { type: XLSX_CONTENT_TYPE }))).toBe(XLSX_CONTENT_TYPE)
    expect(resolveWorkbookImportContentType(new File(['alpha'], 'notes.txt'))).toBeNull()
  })

  it('stages a local preview and imports a new workbook through the authoritative loader', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const previewFile = vi.fn(async () => createPreview())
    const finalizeImport = vi.fn(async () => ({
      kind: 'workbookLoaded',
      id: 'load-1',
      documentId: 'csv:abc123',
      sessionId: 'csv:abc123:browser-import',
      workbookName: 'metrics',
      sheetNames: ['metrics'],
      serverUrl: 'http://127.0.0.1:4321',
      browserUrl: 'http://127.0.0.1:4321/?document=csv%3Aabc123',
      warnings: [],
    }))
    const navigateToWorkbook = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ImportHarness finalizeImport={finalizeImport} navigateToWorkbook={navigateToWorkbook} previewFile={previewFile} />)
    })

    const importToggle = host.querySelector<HTMLButtonElement>("[data-testid='workbook-import-toggle']")
    expect(importToggle?.getAttribute('aria-label')).toBe('Import workbook')
    expect(importToggle?.textContent?.trim()).toBe('')
    expect(importToggle?.getAttribute('class')).toContain('border-transparent')
    expect(importToggle?.getAttribute('class')).toContain('shadow-none')
    expect(importToggle?.getAttribute('class')).toContain('max-[360px]:hidden')

    await act(async () => {
      host.querySelector("[data-testid='workbook-import-toggle']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const importDialog = host.querySelector<HTMLElement>("[data-testid='workbook-import-dialog']")
    expect(importDialog?.getAttribute('class')).toContain('max-w-[28rem]')
    expect(host.querySelector("[data-testid='workbook-import-preview-list']")).toBeNull()

    const input = host.querySelector<HTMLInputElement>("[data-testid='workbook-import-file']")
    const file = new File(['Name,Value\nalpha,12'], 'metrics.csv', { type: CSV_CONTENT_TYPE })
    setInputFiles(input!, [file])

    await act(async () => {
      input?.dispatchEvent(new Event('change', { bubbles: true }))
    })

    expect(previewFile).toHaveBeenCalledWith({
      file,
      contentType: CSV_CONTENT_TYPE,
    })
    expect(host.querySelector<HTMLElement>("[data-testid='workbook-import-dialog']")?.getAttribute('class')).toContain('max-w-[72rem]')
    expect(host.querySelector("[data-testid='workbook-import-preview-list']")).not.toBeNull()
    expect(host.textContent).toContain('metrics')

    await act(async () => {
      host.querySelector("[data-testid='workbook-import-create']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(finalizeImport).toHaveBeenCalledWith({
      file,
      contentType: CSV_CONTENT_TYPE,
      openMode: 'create',
    })
    expect(navigateToWorkbook).toHaveBeenCalledWith(
      expect.objectContaining({
        documentId: 'csv:abc123',
      }),
    )

    await act(async () => {
      root.unmount()
    })
  })

  it('routes replace-current imports through the current document id and surfaces unsupported files', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const previewFile = vi.fn(async () => createPreview())
    const finalizeImport = vi.fn(async () => ({
      kind: 'workbookLoaded',
      id: 'load-2',
      documentId: 'doc-1',
      sessionId: 'doc-1:browser-import',
      workbookName: 'metrics',
      sheetNames: ['metrics'],
      serverUrl: 'http://127.0.0.1:4321',
      warnings: [],
    }))
    const navigateToWorkbook = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ImportHarness finalizeImport={finalizeImport} navigateToWorkbook={navigateToWorkbook} previewFile={previewFile} />)
    })

    await act(async () => {
      host.querySelector("[data-testid='workbook-import-toggle']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const input = host.querySelector<HTMLInputElement>("[data-testid='workbook-import-file']")
    setInputFiles(input!, [new File(['hello'], 'notes.txt', { type: 'text/plain' })])

    await act(async () => {
      input?.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await flushToasts()

    expect(findActiveToast('import-error')?.title).toBe('Only local CSV and XLSX files can be staged for workbook import.')

    const file = new File(['Name,Value\nalpha,12'], 'metrics.csv', { type: CSV_CONTENT_TYPE })
    setInputFiles(input!, [file])

    await act(async () => {
      input?.dispatchEvent(new Event('change', { bubbles: true }))
    })

    await act(async () => {
      host.querySelector("[data-testid='workbook-import-replace']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(finalizeImport).toHaveBeenCalledWith({
      file,
      contentType: CSV_CONTENT_TYPE,
      openMode: 'replace',
      documentId: 'doc-1',
    })
    expect(navigateToWorkbook).toHaveBeenCalledWith(
      expect.objectContaining({
        documentId: 'doc-1',
      }),
    )

    await act(async () => {
      root.unmount()
    })
  })
})
