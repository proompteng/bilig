// @vitest-environment jsdom
import { act, useEffect, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { EditSelectionBehavior } from '@bilig/grid'
import { useWorkbookEditorConflict } from '../use-workbook-editor-conflict.js'
import { emptyCellSnapshot, type EditingMode, type ParsedEditorInput, type WorkbookEditorConflict } from '../worker-workbook-app-model.js'
import type { WorkerRuntimeSelection } from '../runtime-session.js'

interface ConflictHarnessSpies {
  readonly applyParsedInput: ReturnType<typeof vi.fn<(sheetName: string, address: string, parsed: ParsedEditorInput) => Promise<void>>>
  readonly finishEditingWithAuthoritative: ReturnType<typeof vi.fn<(targetSelection: WorkerRuntimeSelection) => void>>
  readonly resetEditorConflictTracking: ReturnType<typeof vi.fn<(nextSelection?: WorkerRuntimeSelection) => void>>
  readonly reportRuntimeError: ReturnType<typeof vi.fn<(error: unknown) => void>>
}

function ConflictHarness(props: {
  selectedCell: CellSnapshot
  authoritativeCell: CellSnapshot
  selection: WorkerRuntimeSelection
  baseSnapshot: CellSnapshot
  spies: ConflictHarnessSpies
}) {
  const [editorConflict, setEditorConflict] = useState<WorkbookEditorConflict | null>(null)
  const [editingMode, setEditingMode] = useState<EditingMode>('formula')
  const [editorSelectionBehavior, setEditorSelectionBehavior] = useState<EditSelectionBehavior>('caret-end')
  const [editorValue] = useState('local-draft')
  const editorValueRef = useRef(editorValue)
  const editorTargetRef = useRef(props.selection)
  const editorBaseSnapshotRef = useRef(props.baseSnapshot)
  const editingModeRef = useRef(editingMode)

  useEffect(() => {
    editorValueRef.current = editorValue
  }, [editorValue])

  useEffect(() => {
    editorTargetRef.current = props.selection
  }, [props.selection])

  useEffect(() => {
    editingModeRef.current = editingMode
  }, [editingMode])

  const banner = useWorkbookEditorConflict({
    editingMode,
    editorValue,
    editorConflict,
    setEditorConflict,
    selectedCell: props.selectedCell,
    selection: props.selection,
    editorValueRef,
    editorTargetRef,
    editorBaseSnapshotRef,
    editingModeRef,
    cloneLiveSelectedCell: () => props.authoritativeCell,
    completeEditNavigation: (targetSelection) => targetSelection,
    finishEditingWithAuthoritative: (targetSelection) => {
      props.spies.finishEditingWithAuthoritative(targetSelection)
      setEditorConflict(null)
      setEditingMode('idle')
    },
    resetEditorConflictTracking: (nextSelection) => {
      props.spies.resetEditorConflictTracking(nextSelection)
      setEditorConflict(null)
    },
    applyParsedInput: props.spies.applyParsedInput,
    reportRuntimeError: props.spies.reportRuntimeError,
    setEditorSelectionBehavior,
    setEditingMode,
  })

  return (
    <div>
      <div data-testid="editing-mode">{editingMode}</div>
      <div data-testid="selection-behavior">{editorSelectionBehavior}</div>
      {banner}
    </div>
  )
}

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

describe('workbook editor conflict', () => {
  it('surfaces a stale same-cell draft and applies the local draft through the compare flow', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const selection = { sheetName: 'Sheet1', address: 'A1' } satisfies WorkerRuntimeSelection
    const baseSnapshot = emptyCellSnapshot('Sheet1', 'A1')
    const authoritativeSnapshot = stringCell('Sheet1', 'A1', 'remote', 2)
    const spies = createConflictHarnessSpies()

    await act(async () => {
      root.render(
        <ConflictHarness
          authoritativeCell={baseSnapshot}
          baseSnapshot={baseSnapshot}
          selectedCell={baseSnapshot}
          selection={selection}
          spies={spies}
        />,
      )
    })

    await act(async () => {
      root.render(
        <ConflictHarness
          authoritativeCell={authoritativeSnapshot}
          baseSnapshot={baseSnapshot}
          selectedCell={authoritativeSnapshot}
          selection={selection}
          spies={spies}
        />,
      )
    })

    const badgeBanner = host.querySelector("[data-testid='editor-conflict-banner']")
    expect(badgeBanner?.textContent).toContain('Remote update detected in Sheet1!A1 while you were editing.')

    await act(async () => {
      host.querySelector("[data-testid='editor-conflict-review']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(host.querySelector("[data-testid='editor-conflict-apply-mine']")).toBeTruthy()
    expect(host.querySelector("[data-testid='editor-conflict-banner']")?.textContent).toContain('Sheet1!A1 changed while you were editing')
    expect(host.querySelector("[data-testid='editor-conflict-banner']")?.textContent).toContain('remote')

    await act(async () => {
      host.querySelector("[data-testid='editor-conflict-apply-mine']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })

    expect(spies.applyParsedInput).toHaveBeenCalledWith('Sheet1', 'A1', {
      kind: 'value',
      value: 'local-draft',
    })
    expect(spies.resetEditorConflictTracking).toHaveBeenCalledWith(selection)
    expect(host.querySelector("[data-testid='editing-mode']")?.textContent).toBe('idle')
    expect(host.querySelector("[data-testid='selection-behavior']")?.textContent).toBe('select-all')
    expect(host.querySelector("[data-testid='editor-conflict-banner']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('clears the conflict when the authoritative value converges to the local draft', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const selection = { sheetName: 'Sheet1', address: 'A1' } satisfies WorkerRuntimeSelection
    const baseSnapshot = emptyCellSnapshot('Sheet1', 'A1')
    const remoteSnapshot = stringCell('Sheet1', 'A1', 'remote', 1)
    const convergedSnapshot = stringCell('Sheet1', 'A1', 'local-draft', 2)
    const spies = createConflictHarnessSpies()

    await act(async () => {
      root.render(
        <ConflictHarness
          authoritativeCell={baseSnapshot}
          baseSnapshot={baseSnapshot}
          selectedCell={baseSnapshot}
          selection={selection}
          spies={spies}
        />,
      )
    })

    await act(async () => {
      root.render(
        <ConflictHarness
          authoritativeCell={remoteSnapshot}
          baseSnapshot={baseSnapshot}
          selectedCell={remoteSnapshot}
          selection={selection}
          spies={spies}
        />,
      )
    })

    expect(host.querySelector("[data-testid='editor-conflict-banner']")).toBeTruthy()

    await act(async () => {
      root.render(
        <ConflictHarness
          authoritativeCell={convergedSnapshot}
          baseSnapshot={baseSnapshot}
          selectedCell={convergedSnapshot}
          selection={selection}
          spies={spies}
        />,
      )
    })

    expect(host.querySelector("[data-testid='editor-conflict-banner']")).toBeNull()
    expect(spies.applyParsedInput).not.toHaveBeenCalled()
    expect(spies.reportRuntimeError).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })
})

function createConflictHarnessSpies(): ConflictHarnessSpies {
  return {
    applyParsedInput: vi.fn(async () => undefined),
    finishEditingWithAuthoritative: vi.fn(),
    resetEditorConflictTracking: vi.fn(),
    reportRuntimeError: vi.fn(),
  }
}

function stringCell(sheetName: string, address: string, value: string, version: number): CellSnapshot {
  return {
    ...emptyCellSnapshot(sheetName, address),
    input: value,
    value: {
      tag: ValueTag.String,
      value,
      stringId: 0,
    },
    version,
  }
}
