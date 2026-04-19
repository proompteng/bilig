import { useEffect, useLayoutEffect, useRef, type CSSProperties } from 'react'
import type { EditMovement } from './SheetGridView.js'

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
  onChange(this: void, next: string): void
  onCommit(this: void, movement?: EditMovement): void
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
  onChange,
  onCommit,
  onCancel,
  style,
}: CellEditorOverlayProps) {
  const inputRef = useRef<HTMLTextAreaElement | null>(null)
  const completionRef = useRef<'idle' | 'commit' | 'cancel'>('idle')
  const blurArmedRef = useRef(false)
  const MAX_EDITOR_HEIGHT = 220

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

  useEffect(() => {
    const textarea = inputRef.current
    if (!textarea) {
      return
    }
    textarea.style.height = '0px'
    const measuredHeight = Math.min(Math.max(textarea.scrollHeight, fontSize + 16), MAX_EDITOR_HEIGHT)
    textarea.style.height = `${measuredHeight}px`
    textarea.style.overflowY = textarea.scrollHeight > MAX_EDITOR_HEIGHT ? 'auto' : 'hidden'
  }, [fontSize, value])

  const commit = (movement?: EditMovement) => {
    if (completionRef.current !== 'idle') {
      return
    }
    completionRef.current = 'commit'
    onCommit(movement)
  }

  const cancel = () => {
    if (completionRef.current !== 'idle') {
      return
    }
    completionRef.current = 'cancel'
    onCancel()
  }

  return (
    <div
      className="cell-editor-overlay overflow-hidden border border-[var(--wb-accent)] bg-[var(--wb-surface)]"
      data-testid="cell-editor-overlay"
      style={{ ...style, backgroundColor }}
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
        value={value}
        onBlur={() => {
          if (!blurArmedRef.current) {
            return
          }
          commit()
        }}
        onChange={(event) => onChange(event.target.value)}
        onKeyDown={(event) => {
          const normalizedNumpadKey = normalizeNumpadKey(event.key, event.code)
          if (normalizedNumpadKey !== null && event.key !== normalizedNumpadKey) {
            event.preventDefault()
            const input = event.currentTarget
            const selectionStart = input.selectionStart ?? value.length
            const selectionEnd = input.selectionEnd ?? value.length
            const nextValue = `${value.slice(0, selectionStart)}${normalizedNumpadKey}${value.slice(selectionEnd)}`
            onChange(nextValue)
            window.requestAnimationFrame(() => {
              const caretPosition = selectionStart + normalizedNumpadKey.length
              inputRef.current?.setSelectionRange(caretPosition, caretPosition)
            })
            return
          }
          if (!event.nativeEvent.isComposing && event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
            event.preventDefault()
            const input = event.currentTarget
            const selectionStart = input.selectionStart ?? value.length
            const selectionEnd = input.selectionEnd ?? value.length
            const nextValue = `${value.slice(0, selectionStart)}${event.key}${value.slice(selectionEnd)}`
            onChange(nextValue)
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
