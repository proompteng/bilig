import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties } from 'react'
import type { EditMovement, EditTargetSelection } from './SheetGridView.js'
import { WORKBOOK_DEFAULT_FONT_SIZE, workbookFontPointSizeToCssPx } from './workbookTheme.js'

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

function moveCaretToBoundary(input: HTMLTextAreaElement, boundary: 'start' | 'end', extendSelection: boolean) {
  const nextPosition = boundary === 'start' ? 0 : input.value.length
  const anchor = boundary === 'start' ? input.selectionEnd : input.selectionStart
  if (extendSelection) {
    input.setSelectionRange(boundary === 'start' ? nextPosition : anchor, boundary === 'start' ? anchor : nextPosition)
    return
  }
  input.setSelectionRange(nextPosition, nextPosition)
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
  fontSize = workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE),
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
  const pendingSelectionRestoreRef = useRef<{
    readonly direction: 'backward' | 'forward' | 'none'
    readonly end: number
    readonly start: number
  } | null>(null)
  const pendingKeyboardSelectionRef = useRef<{
    readonly direction: 'backward' | 'forward' | 'none'
    readonly end: number
    readonly start: number
  } | null>(null)
  const caretWriteSequenceRef = useRef(0)
  const targetSelectionRef = useRef(targetSelection)
  const draftValueRef = useRef(value)
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

  const updateDraftValue = (
    nextValue: string,
    selection?: {
      readonly direction: 'backward' | 'forward' | 'none'
      readonly end: number
      readonly start: number
    },
  ) => {
    if (selection) {
      pendingSelectionRestoreRef.current = selection
    }
    draftValueRef.current = nextValue
    setDraftValue(nextValue)
    onChange(nextValue)
  }

  const preserveCaretSelection = (input: HTMLTextAreaElement) => {
    const selection = {
      direction: input.selectionDirection ?? 'none',
      end: input.selectionEnd ?? input.value.length,
      start: input.selectionStart ?? input.value.length,
    }
    pendingSelectionRestoreRef.current = selection
    caretWriteSequenceRef.current += 1
    const sequence = caretWriteSequenceRef.current
    window.requestAnimationFrame(() => {
      if (caretWriteSequenceRef.current !== sequence || document.activeElement !== input) {
        return
      }
      inputRef.current?.setSelectionRange(selection.start, selection.end, selection.direction)
    })
  }

  const beginCompletion = (nextState: 'commit' | 'cancel') => {
    completionRef.current = nextState
    setIsCompleting(true)
    overlayRef.current?.style.setProperty('pointer-events', 'none')
  }

  const readCurrentDraftValue = () => {
    const input = inputRef.current
    if (!input) {
      return draftValueRef.current
    }
    return pendingKeyboardSelectionRef.current ? draftValueRef.current : input.value
  }

  const insertTextAtSelection = (input: HTMLTextAreaElement, text: string) => {
    const pendingSelection = pendingKeyboardSelectionRef.current
    const currentValue = pendingSelection ? draftValueRef.current : input.value
    const selectionStart = Math.min(pendingSelection?.start ?? input.selectionStart ?? currentValue.length, currentValue.length)
    const selectionEnd = Math.min(pendingSelection?.end ?? input.selectionEnd ?? currentValue.length, currentValue.length)
    const nextValue = `${currentValue.slice(0, selectionStart)}${text}${currentValue.slice(selectionEnd)}`
    const caretPosition = selectionStart + text.length
    const nextSelection = {
      direction: 'none',
      end: caretPosition,
      start: caretPosition,
    } as const
    pendingKeyboardSelectionRef.current = nextSelection
    updateDraftValue(nextValue, nextSelection)
    caretWriteSequenceRef.current += 1
    const sequence = caretWriteSequenceRef.current
    window.requestAnimationFrame(() => {
      if (caretWriteSequenceRef.current !== sequence || document.activeElement !== input) {
        return
      }
      inputRef.current?.setSelectionRange(caretPosition, caretPosition)
    })
  }

  useLayoutEffect(() => {
    blurArmedRef.current = false
    const input = inputRef.current
    input?.focus()
    if (selectionBehavior === 'select-all') {
      input?.select()
    } else {
      const caretPosition = input?.value.length ?? 0
      input?.setSelectionRange(caretPosition, caretPosition)
    }
    const blurArm = window.requestAnimationFrame(() => {
      blurArmedRef.current = true
    })

    return () => {
      window.cancelAnimationFrame(blurArm)
    }
  }, [selectionBehavior, targetAddress, targetSheetName])

  useEffect(() => cancelPendingBlurCommit, [])

  useEffect(() => {
    draftValueRef.current = draftValue
  }, [draftValue])

  useLayoutEffect(() => {
    const restore = pendingSelectionRestoreRef.current
    if (!restore) {
      return
    }
    pendingSelectionRestoreRef.current = null
    pendingKeyboardSelectionRef.current = null
    const input = inputRef.current
    if (!input || document.activeElement !== input) {
      return
    }
    const start = Math.min(restore.start, input.value.length)
    const end = Math.min(restore.end, input.value.length)
    input.setSelectionRange(start, end, restore.direction)
  }, [draftValue])

  useEffect(() => {
    const input = inputRef.current
    const targetChanged = targetSelectionRef.current.address !== targetAddress || targetSelectionRef.current.sheetName !== targetSheetName
    targetSelectionRef.current = {
      address: targetAddress,
      sheetName: targetSheetName,
    }
    const localValue = input?.value ?? draftValueRef.current
    const editorHasFocusedDraft = input && document.activeElement === input && localValue !== value
    if (!targetChanged && editorHasFocusedDraft) {
      return
    }
    if (input && document.activeElement === input) {
      pendingSelectionRestoreRef.current = {
        direction: input.selectionDirection ?? 'none',
        end: input.selectionEnd ?? input.value.length,
        start: input.selectionStart ?? input.value.length,
      }
    }
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
    const nextValue = readCurrentDraftValue()
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
    const nextValue = readCurrentDraftValue()
    const nextTargetSelection = targetSelectionRef.current
    beginCompletion('commit')
    pendingBlurCommitRef.current = window.requestAnimationFrame(() => {
      pendingBlurCommitRef.current = null
      onCommit(undefined, nextValue, nextTargetSelection)
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
        className="w-full resize-none border-0 bg-transparent px-2 py-[3px] leading-[1.2] outline-none"
        data-testid="cell-editor-input"
        ref={inputRef}
        readOnly={isCompleting}
        rows={1}
        style={{
          color,
          font,
          fontSize,
          minHeight: '100%',
          opacity: isCompleting ? 0 : undefined,
          textRendering: 'auto',
          textAlign,
          textDecorationLine: underline ? 'underline' : undefined,
          WebkitFontSmoothing: 'auto',
        }}
        value={draftValue}
        onBlur={commitAfterBlur}
        onChange={(event) =>
          updateDraftValue(event.target.value, {
            direction: event.currentTarget.selectionDirection ?? 'none',
            end: event.currentTarget.selectionEnd ?? event.currentTarget.value.length,
            start: event.currentTarget.selectionStart ?? event.currentTarget.value.length,
          })
        }
        onKeyDown={(event) => {
          const normalizedNumpadKey = normalizeNumpadKey(event.key, event.code)
          if (normalizedNumpadKey !== null && event.key !== normalizedNumpadKey) {
            event.preventDefault()
            insertTextAtSelection(event.currentTarget, normalizedNumpadKey)
            return
          }
          if (!event.nativeEvent.isComposing && event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
            event.preventDefault()
            insertTextAtSelection(event.currentTarget, event.key)
            return
          }
          if ((event.key === 'Home' || event.key === 'End') && !event.ctrlKey && !event.metaKey && !event.altKey) {
            event.preventDefault()
            moveCaretToBoundary(event.currentTarget, event.key === 'Home' ? 'start' : 'end', event.shiftKey)
            preserveCaretSelection(event.currentTarget)
            return
          }
          if (event.key === 'Enter') {
            if (event.altKey) {
              event.preventDefault()
              insertTextAtSelection(event.currentTarget, '\n')
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
