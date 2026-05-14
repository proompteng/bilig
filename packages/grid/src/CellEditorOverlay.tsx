import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties } from 'react'
import type { EditMovement, EditTargetSelection } from './SheetGridView.js'

function normalizeNumpadKey(key: string, code: string): string | null {
  if (!code.startsWith('Numpad')) {
    return null
  }
  const suffix = code.slice('Numpad'.length)
  if (/^\d$/.test(suffix)) {
    return suffix
  }
  if (suffix === 'Decimal') {
    return '.'
  }
  if (suffix === 'Add') {
    return '+'
  }
  if (suffix === 'Subtract') {
    return '-'
  }
  if (suffix === 'Multiply') {
    return '*'
  }
  if (suffix === 'Divide') {
    return '/'
  }
  return key.length === 1 ? key : null
}

interface CellEditorOverlayProps {
  label: string
  value: string
  resolvedValue: string
  selectionBehavior?: 'select-all' | 'caret-end'
  textAlign?: 'left' | 'right'
  backgroundColor?: string
  color?: string
  font?: string
  fontSize?: number
  underline?: boolean
  targetSelection: EditTargetSelection
  onChange(this: void, next: string): void
  onCommit(this: void, movement?: EditMovement, valueOverride?: string, targetSelectionOverride?: EditTargetSelection): void
  onCancel(this: void): void
  style?: CSSProperties
}

export function CellEditorOverlay({
  label,
  value,
  resolvedValue: _resolvedValue,
  selectionBehavior = 'select-all',
  textAlign = 'left',
  backgroundColor = '#ffffff',
  color = '#202124',
  font,
  fontSize = 13,
  underline = false,
  targetSelection,
  onChange,
  onCommit,
  onCancel,
  style,
}: CellEditorOverlayProps) {
  const targetAddress = targetSelection.address
  const targetSheetName = targetSelection.sheetName
  const overlayRef = useRef<HTMLDivElement | null>(null)
  const inputRef = useRef<HTMLTextAreaElement | null>(null)
  const completionRef = useRef<'idle' | 'commit' | 'cancel'>('idle')
  const blurArmedRef = useRef(false)
  const pendingBlurCommitRef = useRef<number | null>(null)
  const pendingParentSyncRef = useRef<number | null>(null)
  const pendingParentSyncValueRef = useRef(value)
  const targetSelectionRef = useRef(targetSelection)
  const [isCompleting, setIsCompleting] = useState(false)
  const [draftValue, setDraftValue] = useState(value)
  const MAX_EDITOR_HEIGHT = 220

  const cancelPendingBlurCommit = () => {
    const pendingFrame = pendingBlurCommitRef.current
    if (pendingFrame === null) {
      return
    }
    pendingBlurCommitRef.current = null
    window.cancelAnimationFrame(pendingFrame)
  }

  const cancelPendingParentSync = () => {
    const pendingFrame = pendingParentSyncRef.current
    if (pendingFrame === null) {
      return
    }
    pendingParentSyncRef.current = null
    window.cancelAnimationFrame(pendingFrame)
  }

  const flushParentSync = (nextValue = pendingParentSyncValueRef.current) => {
    cancelPendingParentSync()
    onChange(nextValue)
  }

  const updateDraftValue = (nextValue: string) => {
    pendingParentSyncValueRef.current = nextValue
    setDraftValue(nextValue)
    if (pendingParentSyncRef.current !== null) {
      return
    }
    pendingParentSyncRef.current = window.requestAnimationFrame(() => {
      pendingParentSyncRef.current = null
      onChange(pendingParentSyncValueRef.current)
    })
  }

  const beginCompletion = (nextState: 'commit' | 'cancel') => {
    completionRef.current = nextState
    setIsCompleting(true)
    overlayRef.current?.style.setProperty('pointer-events', 'none')
  }

  useLayoutEffect(() => {
    blurArmedRef.current = false
    inputRef.current?.focus()
    if (selectionBehavior === 'select-all') {
      inputRef.current?.select()
    } else {
      const caretPosition = value.length
      inputRef.current?.setSelectionRange(caretPosition, caretPosition)
    }
    const blurArm = window.requestAnimationFrame(() => {
      blurArmedRef.current = true
    })

    return () => {
      window.cancelAnimationFrame(blurArm)
    }
  }, [selectionBehavior, value.length])

  useEffect(() => cancelPendingBlurCommit, [])
  useEffect(() => cancelPendingParentSync, [])

  useEffect(() => {
    targetSelectionRef.current = {
      address: targetAddress,
      sheetName: targetSheetName,
    }
    pendingParentSyncValueRef.current = value
    setDraftValue(value)
  }, [targetAddress, targetSheetName, value])

  useEffect(() => {
    const textarea = inputRef.current
    if (!textarea) {
      return
    }
    textarea.style.height = '0px'
    const measuredHeight = Math.min(Math.max(textarea.scrollHeight, fontSize + 16), MAX_EDITOR_HEIGHT)
    textarea.style.height = `${measuredHeight}px`
    textarea.style.overflowY = textarea.scrollHeight > MAX_EDITOR_HEIGHT ? 'auto' : 'hidden'
  }, [draftValue, fontSize])

  const commit = (movement?: EditMovement) => {
    if (movement) {
      cancelPendingBlurCommit()
    }
    const nextValue = inputRef.current?.value ?? draftValue
    flushParentSync(nextValue)
    if (completionRef.current !== 'idle') {
      if (movement && completionRef.current === 'commit') {
        onCommit(movement, nextValue, targetSelectionRef.current)
      }
      return
    }
    cancelPendingBlurCommit()
    beginCompletion('commit')
    onCommit(movement, nextValue, targetSelectionRef.current)
  }

  const commitAfterBlur = () => {
    if (!blurArmedRef.current || completionRef.current !== 'idle' || pendingBlurCommitRef.current !== null) {
      return
    }
    pendingBlurCommitRef.current = window.requestAnimationFrame(() => {
      pendingBlurCommitRef.current = null
      commit()
    })
  }

  const cancel = () => {
    if (completionRef.current !== 'idle') {
      return
    }
    cancelPendingBlurCommit()
    beginCompletion('cancel')
    onCancel()
  }

  return (
    <div
      className="cell-editor-overlay box-border overflow-hidden border border-[var(--wb-accent)] bg-[var(--wb-surface)]"
      data-completing={isCompleting ? 'true' : undefined}
      data-testid="cell-editor-overlay"
      ref={overlayRef}
      style={isCompleting ? { ...style, backgroundColor, pointerEvents: 'none' } : { ...style, backgroundColor }}
    >
      <textarea
        aria-label={`${label} editor`}
        className="w-full resize-none border-0 bg-transparent px-2 py-1.5 leading-tight outline-none"
        data-testid="cell-editor-input"
        ref={inputRef}
        rows={1}
        style={{
          color,
          font,
          fontSize,
          minHeight: '100%',
          textAlign,
          textDecorationLine: underline ? 'underline' : undefined,
        }}
        value={draftValue}
        onBlur={commitAfterBlur}
        onChange={(event) => updateDraftValue(event.target.value)}
        onKeyDown={(event) => {
          const normalizedNumpadKey = normalizeNumpadKey(event.key, event.code)
          if (normalizedNumpadKey !== null && event.key !== normalizedNumpadKey) {
            event.preventDefault()
            const input = event.currentTarget
            const currentValue = input.value
            const selectionStart = input.selectionStart ?? currentValue.length
            const selectionEnd = input.selectionEnd ?? currentValue.length
            const nextValue = `${currentValue.slice(0, selectionStart)}${normalizedNumpadKey}${currentValue.slice(selectionEnd)}`
            input.value = nextValue
            updateDraftValue(nextValue)
            window.requestAnimationFrame(() => {
              const caretPosition = selectionStart + normalizedNumpadKey.length
              inputRef.current?.setSelectionRange(caretPosition, caretPosition)
            })
            return
          }
          if (!event.nativeEvent.isComposing && event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
            event.preventDefault()
            const input = event.currentTarget
            const currentValue = input.value
            const selectionStart = input.selectionStart ?? currentValue.length
            const selectionEnd = input.selectionEnd ?? currentValue.length
            const nextValue = `${currentValue.slice(0, selectionStart)}${event.key}${currentValue.slice(selectionEnd)}`
            input.value = nextValue
            updateDraftValue(nextValue)
            window.requestAnimationFrame(() => {
              const caretPosition = selectionStart + event.key.length
              inputRef.current?.setSelectionRange(caretPosition, caretPosition)
            })
            return
          }
          if (event.key === 'Enter') {
            if (event.altKey) {
              return
            }
            event.preventDefault()
            commit([0, event.shiftKey ? -1 : 1])
            return
          }
          if (event.key === 'Tab') {
            event.preventDefault()
            commit([event.shiftKey ? -1 : 1, 0])
            return
          }
          if (event.key === 'Escape') {
            event.preventDefault()
            cancel()
          }
        }}
      />
    </div>
  )
}
