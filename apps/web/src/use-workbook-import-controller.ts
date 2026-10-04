import type { WorkbookLoadedResponse } from '@bilig/agent-api'
import type { ImportedWorkbookPreview } from '@bilig/excel-import/browser'
import { useCallback, useEffect, useRef, useState } from 'react'
import { createLatestRequestGate } from './latest-request.js'
import { resolveWorkbookImportContentType } from './workbook-import-client.js'
import type { finalizeWorkbookImport, previewWorkbookImport } from './workbook-import-client.js'

interface StagedWorkbookImport {
  file: File
  preview: ImportedWorkbookPreview
}

export function useWorkbookImportController(input: {
  readonly currentDocumentId: string
  readonly enabled: boolean
  readonly previewFile: typeof previewWorkbookImport
  readonly finalizeImport: typeof finalizeWorkbookImport
  readonly navigateToWorkbook: (result: WorkbookLoadedResponse) => void
}) {
  const { currentDocumentId, enabled, finalizeImport, navigateToWorkbook, previewFile } = input
  const [isOpen, setIsOpen] = useState(false)
  const [isPreviewing, setIsPreviewing] = useState(false)
  const [isImporting, setIsImporting] = useState(false)
  const [stagedImport, setStagedImport] = useState<StagedWorkbookImport | null>(null)
  const [error, setError] = useState<string | null>(null)
  const mountedRef = useRef(true)
  const previewGateRef = useRef(createLatestRequestGate())
  const importInFlightRef = useRef(false)

  useEffect(() => {
    const previewGate = previewGateRef.current
    mountedRef.current = true
    return () => {
      mountedRef.current = false
      previewGate.invalidate()
    }
  }, [])

  useEffect(() => {
    if (enabled) {
      return
    }
    previewGateRef.current.invalidate()
    setIsPreviewing(false)
  }, [enabled])

  const stageFile = useCallback(
    async (file: File | null) => {
      const requestId = previewGateRef.current.begin()
      if (!enabled || file === null) {
        setStagedImport(null)
        setIsPreviewing(false)
        return
      }
      const contentType = resolveWorkbookImportContentType(file)
      if (!contentType) {
        setStagedImport(null)
        setError('Only local Bilig backups, CSV, XLSX, and XLSM files can be staged for workbook import.')
        setIsOpen(true)
        setIsPreviewing(false)
        return
      }
      setIsOpen(true)
      setStagedImport(null)
      setError(null)
      setIsPreviewing(true)
      try {
        const preview = await previewFile({ file, contentType })
        if (mountedRef.current && previewGateRef.current.isCurrent(requestId)) {
          setStagedImport({ file, preview })
        }
      } catch (nextError) {
        if (mountedRef.current && previewGateRef.current.isCurrent(requestId)) {
          setStagedImport(null)
          setError(nextError instanceof Error ? nextError.message : String(nextError))
        }
      } finally {
        if (mountedRef.current && previewGateRef.current.isCurrent(requestId)) {
          setIsPreviewing(false)
        }
      }
    },
    [enabled, previewFile],
  )

  const importStagedFile = useCallback(
    async (openMode: 'create' | 'replace') => {
      if (!enabled || !stagedImport || importInFlightRef.current) {
        return
      }
      importInFlightRef.current = true
      setError(null)
      setIsImporting(true)
      try {
        const result = await finalizeImport({
          file: stagedImport.file,
          contentType: stagedImport.preview.contentType,
          openMode,
          ...(openMode === 'replace' ? { documentId: currentDocumentId } : {}),
        })
        if (mountedRef.current) {
          navigateToWorkbook(result)
        }
      } catch (nextError) {
        if (mountedRef.current) {
          setError(nextError instanceof Error ? nextError.message : String(nextError))
        }
      } finally {
        importInFlightRef.current = false
        if (mountedRef.current) {
          setIsImporting(false)
        }
      }
    },
    [currentDocumentId, enabled, finalizeImport, navigateToWorkbook, stagedImport],
  )

  return {
    clearError: useCallback(() => setError(null), []),
    open: useCallback(() => setIsOpen(true), []),
    close: useCallback(() => setIsOpen(false), []),
    error,
    importStagedFile,
    isImporting,
    isOpen,
    isPreviewing,
    stageFile,
    stagedPreview: stagedImport?.preview ?? null,
    toggle: useCallback(() => setIsOpen((current) => !current), []),
  }
}
